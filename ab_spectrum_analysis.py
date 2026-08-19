"""Independent event-level EEG spectrum analysis for Figure A and Figure B.

This module intentionally does not import or depend on the fatigue-prediction
pipeline. It reads manifest-selected raw EDF files and folder-level eyeblink.dat
files, excludes eye-contaminated pre-deviation windows, computes one-sided PSD
values, and writes per-recording plus recording-balanced aggregate products.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Sequence

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyedflib
from scipy.signal import butter, sosfiltfilt


# ---------------------------------------------------------------------------
# Central analysis settings
# ---------------------------------------------------------------------------

CHANNEL = "FP2"
STATUS_CHANNEL = "Status"
RAW_EDF_STEM_SUFFIX = "_raw"

PRE_EVENT_START = -1.0
PRE_EVENT_END = 0.0

PSD_FREQ_MIN = 1.0
PSD_FREQ_MAX = 30.0
PSD_NFFT = 512
PSD_WINDOW = "boxcar"
PSD_DETREND = False

RT_MIN_VALID = 0.2
RT_THRESHOLD = 1.6
RT_CAP_FOR_HEATMAP = 5.0
RT_BIN_WIDTH = 0.1

DROWSY_EARLY_FRACTION = 0.20

EPS = 1e-12

EEG_HIGHPASS_HZ = 1.0
EEG_LOWPASS_HZ = 30.0
EEG_FILTER_ORDER = 4

THETA_BAND = (4.0, 7.0)
ALPHA_BAND = (8.0, 12.0)
BETA_BAND = (13.0, 20.0)
BANDS = {
    "theta": THETA_BAND,
    "alpha": ALPHA_BAND,
    "beta": BETA_BAND,
}

DEVIATION_START_CODES = frozenset({251, 252})
RESPONSE_START_CODE = 253
RESPONSE_END_CODE = 254

OUTPUT_DIR = Path(__file__).resolve().parent / "output_ab_spectrum"
TRAIN_DATA_PATH = Path(__file__).resolve().parent / "train_data"
FIGURE_DPI = 200


@dataclass
class ReactionTimeEvent:
    """One 251/252 -> 253 lane-deviation event from the EDF Status channel."""

    event_index: int
    deviation_status: int
    deviation_onset_time: float
    response_onset_time: float
    reaction_time_raw: float
    reaction_time: float
    response_end_time: float | None = None


@dataclass(frozen=True)
class EyeInformation:
    """Eye events loaded exclusively from folder-level eyeblink.dat."""

    source: str
    seconds: tuple[int, ...]
    source_path: Path | None = None


@dataclass
class EventPSDRecord:
    """Auditable event-level spectrum record."""

    event: ReactionTimeEvent
    rt_valid: bool
    rt_bin: str | None
    is_alert: bool
    is_drowsy: bool
    window_start_time: float
    window_end_time: float
    eye_contaminated: bool
    used_for_eeg_analysis: bool
    exclusion_reason: str
    psd: np.ndarray | None = None
    log_psd: np.ndarray | None = None
    band_power: dict[str, float] = field(default_factory=dict)
    band_log_power: dict[str, float] = field(default_factory=dict)
    is_early_drowsy: bool = False
    is_selected_alert: bool = False


@dataclass(frozen=True)
class RecordingAnalysis:
    """All products needed to export one recording."""

    edf_path: Path
    channel_name: str
    sample_rate: float
    source_unit: str
    eye_information: EyeInformation
    frequencies: np.ndarray
    records: tuple[EventPSDRecord, ...]
    early_drowsy_n: int
    first_valid_drowsy_time: float | None


@dataclass(frozen=True)
class TrainingInput:
    """One manifest-selected dataset folder after strict preflight validation."""

    recording_id: str
    folder_path: Path
    edf_path: Path
    eyeblink_path: Path


def round_half_up_one_decimal(value: float) -> float:
    """Round an RT to 0.1 seconds using conventional half-up rounding."""

    if not math.isfinite(value):
        raise ValueError(f"Reaction Time must be finite: {value}")
    return float(Decimal(str(value)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def find_channel_index(labels: Sequence[str], target: str) -> int:
    """Find a channel case-insensitively, preferring an exact label match."""

    normalized_target = str(target).strip().casefold()
    for index, label in enumerate(labels):
        if str(label).strip().casefold() == normalized_target:
            return index
    for index, label in enumerate(labels):
        if normalized_target in str(label).strip().casefold():
            return index
    raise ValueError(f"Channel '{target}' was not found. Available channels: {list(labels)}")


def signal_to_microvolts(signal: np.ndarray, physical_dimension: str) -> np.ndarray:
    """Convert a physical EDF signal to microvolts."""

    unit = str(physical_dimension).strip().casefold().replace("μ", "u").replace("µ", "u")
    factors = {
        "uv": 1.0,
        "mv": 1_000.0,
        "v": 1_000_000.0,
        "nv": 0.001,
    }
    if unit not in factors:
        raise ValueError(
            f"Unsupported EEG physical dimension '{physical_dimension}'. "
            "Expected V, mV, uV, µV, μV, or nV."
        )
    return np.asarray(signal, dtype=float) * factors[unit]


def parse_reaction_time_events(
    status_signal: np.ndarray, sample_rate: float
) -> list[ReactionTimeEvent]:
    """Parse transition runs in a Status signal into RT events."""

    if sample_rate <= 0:
        raise ValueError(f"Invalid Status sample rate: {sample_rate}")
    rounded_status = np.rint(np.asarray(status_signal)).astype(int)
    if rounded_status.size == 0:
        return []

    change_indices = np.flatnonzero(
        np.r_[True, rounded_status[1:] != rounded_status[:-1]]
    )
    events: list[ReactionTimeEvent] = []
    deviation_status: int | None = None
    deviation_sample: int | None = None
    pending_completion: ReactionTimeEvent | None = None

    for sample_index_value in change_indices:
        sample_index = int(sample_index_value)
        status_code = int(rounded_status[sample_index])

        if status_code in DEVIATION_START_CODES:
            deviation_status = status_code
            deviation_sample = sample_index
            pending_completion = None
            continue

        if status_code == RESPONSE_START_CODE:
            if deviation_status is None or deviation_sample is None:
                continue
            raw_rt = float(sample_index - deviation_sample) / float(sample_rate)
            if raw_rt < 0:
                deviation_status = None
                deviation_sample = None
                continue
            event = ReactionTimeEvent(
                event_index=len(events) + 1,
                deviation_status=deviation_status,
                deviation_onset_time=float(deviation_sample) / float(sample_rate),
                response_onset_time=float(sample_index) / float(sample_rate),
                reaction_time_raw=raw_rt,
                reaction_time=round_half_up_one_decimal(raw_rt),
            )
            events.append(event)
            pending_completion = event
            deviation_status = None
            deviation_sample = None
            continue

        if status_code == RESPONSE_END_CODE and pending_completion is not None:
            pending_completion.response_end_time = float(sample_index) / float(sample_rate)
            pending_completion = None

    return events


def read_recording(
    edf_path: Path, channel: str
) -> tuple[np.ndarray, float, str, str, list[ReactionTimeEvent]]:
    """Read the configured EEG channel plus independently sampled Status events."""

    reader = pyedflib.EdfReader(str(edf_path))
    try:
        labels = reader.getSignalLabels()
        eeg_index = find_channel_index(labels, channel)
        status_index = find_channel_index(labels, STATUS_CHANNEL)
        eeg_signal = reader.readSignal(eeg_index).astype(float)
        eeg_sample_rate = float(reader.getSampleFrequency(eeg_index))
        eeg_unit = str(reader.getPhysicalDimension(eeg_index))
        resolved_channel = str(labels[eeg_index])
        status_signal = reader.readSignal(status_index)
        status_sample_rate = float(reader.getSampleFrequency(status_index))
    finally:
        reader.close()

    signal_uv = signal_to_microvolts(eeg_signal, eeg_unit)
    events = parse_reaction_time_events(status_signal, status_sample_rate)
    return signal_uv, eeg_sample_rate, eeg_unit, resolved_channel, events


def apply_record_alpha_filter(
    signal: np.ndarray,
    sample_rate: float,
    *,
    highpass_hz: float,
    lowpass_hz: float,
    order: int,
) -> np.ndarray:
    """Apply separate Butterworth high/low filters with SOS zero-phase filtering."""

    fs = float(sample_rate)
    if fs <= 0:
        raise ValueError(f"Invalid sampling rate: {sample_rate}")
    nyquist = fs / 2.0
    if not 0 < highpass_hz < lowpass_hz < nyquist:
        raise ValueError(
            f"Invalid filter range {highpass_hz}-{lowpass_hz} Hz for fs={fs} Hz"
        )
    values = np.asarray(signal, dtype=float)
    high_sos = butter(order, highpass_hz, btype="highpass", fs=fs, output="sos")
    filtered = sosfiltfilt(high_sos, values)
    low_sos = butter(order, lowpass_hz, btype="lowpass", fs=fs, output="sos")
    return sosfiltfilt(low_sos, filtered)


def parse_eyeblink_dat(dat_path: Path) -> tuple[int, ...]:
    """Read and validate the project's count,second,... eyeblink.dat format."""

    text = dat_path.read_text(encoding="utf-8-sig").strip()
    if not text:
        raise ValueError(f"eyeblink.dat is empty: {dat_path}")
    tokens = [token for token in re.split(r"[,\s]+", text) if token]
    try:
        values = [int(token) for token in tokens]
    except ValueError as exc:
        raise ValueError(f"eyeblink.dat contains a non-integer value: {dat_path}") from exc
    if not values:
        raise ValueError(f"eyeblink.dat does not contain a declared count: {dat_path}")
    declared_count, seconds = values[0], values[1:]
    if declared_count != len(seconds):
        raise ValueError(
            f"eyeblink.dat count mismatch in {dat_path}: "
            f"declared {declared_count}, found {len(seconds)}"
        )
    invalid_seconds = [second for second in seconds if second < 1]
    if invalid_seconds:
        raise ValueError(
            f"eyeblink.dat contains {len(invalid_seconds)} seconds below 1: {dat_path}"
        )
    return tuple(sorted(set(seconds)))


def find_eyeblink_dat(folder_path: Path) -> Path | None:
    """Return the folder-level eyeblink.dat using a case-insensitive exact name."""

    matches = [
        path
        for path in folder_path.iterdir()
        if path.is_file() and path.name.casefold() == "eyeblink.dat"
    ]
    if len(matches) > 1:
        raise ValueError(f"Multiple case-variant eyeblink.dat files in {folder_path}")
    return matches[0] if matches else None


def build_eye_information(eyeblink_dat_path: Path) -> EyeInformation:
    """Load eye seconds exclusively from a file named eyeblink.dat."""

    if eyeblink_dat_path.name.casefold() != "eyeblink.dat":
        raise ValueError(
            f"Only folder-level eyeblink.dat is allowed, got: {eyeblink_dat_path.name}"
        )
    seconds = parse_eyeblink_dat(eyeblink_dat_path)
    return EyeInformation(
        source="eyeblink_dat",
        seconds=seconds,
        source_path=eyeblink_dat_path,
    )


def parse_training_manifest(manifest_path: Path) -> list[str]:
    """Parse recording IDs separated by Chinese commas, commas, or whitespace."""

    if not manifest_path.is_file():
        raise FileNotFoundError(f"train_data manifest not found: {manifest_path}")
    text = manifest_path.read_text(encoding="utf-8-sig").strip()
    recording_ids = [
        token.strip()
        for token in re.split(r"[\u3001\uff0c,\s]+", text)
        if token.strip()
    ]
    if not recording_ids:
        raise ValueError(f"train_data manifest is empty: {manifest_path}")
    normalized = [recording_id.casefold() for recording_id in recording_ids]
    duplicates = sorted(
        {
            recording_ids[index]
            for index, value in enumerate(normalized)
            if normalized.count(value) > 1
        }
    )
    if duplicates:
        raise ValueError(
            "train_data contains duplicate recording IDs: " + ", ".join(duplicates)
        )
    return recording_ids


def normalized_dataset_folder_name(folder_name: str) -> str:
    name = str(folder_name).strip()
    if name.casefold().endswith(".set"):
        name = name[:-4]
    return name.casefold()


def validate_edf_for_analysis(edf_path: Path, channel: str) -> tuple[float, float]:
    """Validate required channels and return (EEG sampling rate, duration)."""

    reader = pyedflib.EdfReader(str(edf_path))
    try:
        labels = reader.getSignalLabels()
        eeg_index = find_channel_index(labels, channel)
        find_channel_index(labels, STATUS_CHANNEL)
        sample_rate = float(reader.getSampleFrequency(eeg_index))
        sample_count = int(reader.getNSamples()[eeg_index])
        physical_dimension = str(reader.getPhysicalDimension(eeg_index))
    finally:
        reader.close()
    signal_to_microvolts(np.zeros(1, dtype=float), physical_dimension)
    if sample_rate <= 0:
        raise ValueError(f"Invalid sampling rate in {edf_path}: {sample_rate}")
    samples_per_window = int(round((PRE_EVENT_END - PRE_EVENT_START) * sample_rate))
    if samples_per_window > PSD_NFFT:
        raise ValueError(
            f"One-second window has {samples_per_window} samples but PSD_NFFT={PSD_NFFT}"
        )
    return sample_rate, float(sample_count) / sample_rate


def preflight_training_inputs(
    input_root: Path,
    manifest_path: Path,
    *,
    channel: str = CHANNEL,
) -> tuple[list[TrainingInput], list[dict[str, object]]]:
    """Resolve every manifest entry and validate raw EDF plus eyeblink.dat."""

    recording_ids = parse_training_manifest(manifest_path)
    child_folders = [path for path in input_root.iterdir() if path.is_dir()]
    folder_map: dict[str, list[Path]] = {}
    for folder in child_folders:
        folder_map.setdefault(normalized_dataset_folder_name(folder.name), []).append(folder)

    resolved: list[TrainingInput] = []
    rows: list[dict[str, object]] = []
    for manifest_index, recording_id in enumerate(recording_ids, start=1):
        row: dict[str, object] = {
            "manifest_index": manifest_index,
            "recording_id": recording_id,
            "folder_path": None,
            "edf_path": None,
            "eyeblink_path": None,
            "preflight_status": "FAILED",
            "preflight_error": None,
            "analysis_status": "NOT_RUN",
            "analysis_error": None,
            "parsed_events": None,
            "used_for_eeg_analysis": None,
            "selected_alert_n": None,
            "early_drowsy_n": None,
        }
        try:
            folder_matches = folder_map.get(recording_id.casefold(), [])
            if len(folder_matches) != 1:
                raise FileNotFoundError(
                    f"Expected exactly one folder named '{recording_id}' or "
                    f"'{recording_id}.set', found {len(folder_matches)}"
                )
            folder = folder_matches[0]
            row["folder_path"] = str(folder.resolve())

            raw_edfs = sorted(
                path
                for path in folder.iterdir()
                if path.is_file()
                and path.suffix.casefold() == ".edf"
                and path.stem.casefold().endswith(RAW_EDF_STEM_SUFFIX.casefold())
            )
            if len(raw_edfs) != 1:
                raise FileNotFoundError(
                    f"Expected exactly one *_raw.EDF, found {len(raw_edfs)}"
                )
            edf_path = raw_edfs[0]
            row["edf_path"] = str(edf_path.resolve())

            eyeblink_path = find_eyeblink_dat(folder)
            if eyeblink_path is None:
                raise FileNotFoundError(f"Missing eyeblink.dat in {folder}")
            row["eyeblink_path"] = str(eyeblink_path.resolve())
            eye_seconds = parse_eyeblink_dat(eyeblink_path)

            _, duration = validate_edf_for_analysis(edf_path, channel)
            out_of_range = [second for second in eye_seconds if second > math.ceil(duration)]
            if out_of_range:
                raise ValueError(
                    f"eyeblink.dat contains {len(out_of_range)} seconds beyond "
                    f"EDF duration {duration:.3f}s"
                )

            row["preflight_status"] = "READY"
            resolved.append(
                TrainingInput(
                    recording_id=recording_id,
                    folder_path=folder,
                    edf_path=edf_path,
                    eyeblink_path=eyeblink_path,
                )
            )
        except Exception as exc:
            row["preflight_error"] = str(exc)
        rows.append(row)
    return resolved, rows


def eye_overlaps_window(
    eye_information: EyeInformation, window_start: float, window_end: float
) -> bool:
    """Return whether a coarse eyeblink.dat second overlaps [start, end)."""

    # eyeblink.dat values are ceil(peak_time), so second s represents a peak in
    # the coarse interval (s-1, s]. Conservatively reject any overlapping window.
    return any(
        float(second) >= window_start and float(second - 1) < window_end
        for second in eye_information.seconds
    )


def nearest_frequency_slice(
    frequencies: np.ndarray, low_hz: float, high_hz: float
) -> slice:
    if low_hz > high_hz:
        raise ValueError("low_hz must be <= high_hz")
    low_index = int(np.argmin(np.abs(frequencies - float(low_hz))))
    high_index = int(np.argmin(np.abs(frequencies - float(high_hz))))
    if low_index > high_index:
        low_index, high_index = high_index, low_index
    return slice(low_index, high_index + 1)


def compute_one_sided_psd(
    segment_uv: np.ndarray,
    sample_rate: float,
    *,
    nfft: int = PSD_NFFT,
) -> tuple[np.ndarray, np.ndarray]:
    """Compute a boxcar, non-detrended, one-sided PSD in uV^2/Hz."""

    segment = np.asarray(segment_uv, dtype=float)
    if segment.ndim != 1 or segment.size == 0:
        raise ValueError("PSD segment must be a non-empty one-dimensional array")
    if not np.all(np.isfinite(segment)):
        raise ValueError("PSD segment contains non-finite values")
    if sample_rate <= 0:
        raise ValueError(f"Invalid sampling rate: {sample_rate}")
    if nfft < segment.size:
        raise ValueError(
            f"PSD_NFFT={nfft} is smaller than the one-second window "
            f"({segment.size} samples at fs={sample_rate} Hz)."
        )

    spectrum = np.fft.rfft(segment, n=nfft)
    # Boxcar window energy is N. This is scipy.signal.periodogram's density
    # normalization without detrending, written explicitly to keep the FFT
    # relationship auditable.
    psd = np.abs(spectrum).astype(float, copy=False) ** 2
    psd /= float(sample_rate) * float(segment.size)
    if nfft % 2 == 0:
        psd[1:-1] *= 2.0
    else:
        psd[1:] *= 2.0
    frequencies = np.fft.rfftfreq(nfft, d=1.0 / float(sample_rate))
    selected = nearest_frequency_slice(frequencies, PSD_FREQ_MIN, PSD_FREQ_MAX)
    return frequencies[selected], psd[selected]


def compute_band_power(
    psd: np.ndarray,
    frequencies: np.ndarray,
    low_hz: float,
    high_hz: float,
) -> float:
    """Integrate selected FFT density bins using their uniform bin width."""

    selected = nearest_frequency_slice(frequencies, low_hz, high_hz)
    selected_psd = np.asarray(psd[selected], dtype=float)
    if frequencies.size < 2:
        raise ValueError("At least two frequency bins are required")
    bin_width = float(frequencies[1] - frequencies[0])
    return float(np.sum(selected_psd) * bin_width)


def rt_bin_label(reaction_time: float) -> str:
    if reaction_time >= RT_CAP_FOR_HEATMAP:
        return f">={RT_CAP_FOR_HEATMAP:.1f}"
    return f"{reaction_time:.1f}"


def classify_rounded_rt(reaction_time: float) -> tuple[bool, bool, bool]:
    """Return (valid, alert, drowsy) after RT has been rounded to 0.1 s."""

    rt_valid = float(reaction_time) > RT_MIN_VALID
    is_drowsy = rt_valid and float(reaction_time) >= RT_THRESHOLD
    is_alert = rt_valid and float(reaction_time) < RT_THRESHOLD
    return rt_valid, is_alert, is_drowsy


def heatmap_bin_labels() -> list[str]:
    start = Decimal(str(RT_MIN_VALID)) + Decimal(str(RT_BIN_WIDTH))
    cap = Decimal(str(RT_CAP_FOR_HEATMAP))
    width = Decimal(str(RT_BIN_WIDTH))
    labels: list[str] = []
    value = start
    while value < cap:
        labels.append(f"{float(value):.1f}")
        value += width
    labels.append(f">={RT_CAP_FOR_HEATMAP:.1f}")
    return labels


def assign_figure_b_groups(
    records: Sequence[EventPSDRecord],
) -> tuple[int, float | None]:
    """Mark earliest 20% valid Drowsy and all valid Alert before its onset."""

    for record in records:
        record.is_early_drowsy = False
        record.is_selected_alert = False

    valid_drowsy = sorted(
        (
            record
            for record in records
            if record.used_for_eeg_analysis and record.is_drowsy
        ),
        key=lambda record: (record.event.deviation_onset_time, record.event.event_index),
    )
    early_drowsy_n = (
        max(1, math.ceil(len(valid_drowsy) * DROWSY_EARLY_FRACTION))
        if valid_drowsy
        else 0
    )
    for record in valid_drowsy[:early_drowsy_n]:
        record.is_early_drowsy = True

    first_valid_drowsy_time = (
        valid_drowsy[0].event.deviation_onset_time if valid_drowsy else None
    )
    if first_valid_drowsy_time is not None:
        for record in records:
            record.is_selected_alert = bool(
                record.used_for_eeg_analysis
                and record.is_alert
                and record.event.deviation_onset_time < first_valid_drowsy_time
            )
    return early_drowsy_n, first_valid_drowsy_time


def analyze_recording(
    edf_path: Path,
    *,
    eyeblink_dat_path: Path,
    channel: str = CHANNEL,
) -> RecordingAnalysis:
    signal_uv, sample_rate, source_unit, resolved_channel, events = read_recording(
        edf_path, channel
    )
    if not events:
        raise ValueError("No 251/252 -> 253 reaction-time events were found")

    samples_per_window = int(round((PRE_EVENT_END - PRE_EVENT_START) * sample_rate))
    if samples_per_window <= 0:
        raise ValueError("PRE_EVENT_END must be greater than PRE_EVENT_START")
    if samples_per_window > PSD_NFFT:
        raise ValueError(
            f"One-second window has {samples_per_window} samples, but PSD_NFFT={PSD_NFFT}."
        )

    eye_information = build_eye_information(eyeblink_dat_path)
    filtered_uv = apply_record_alpha_filter(
        signal_uv,
        sample_rate,
        highpass_hz=EEG_HIGHPASS_HZ,
        lowpass_hz=EEG_LOWPASS_HZ,
        order=EEG_FILTER_ORDER,
    )

    frequencies, _ = compute_one_sided_psd(
        np.zeros(samples_per_window, dtype=float), sample_rate
    )
    records: list[EventPSDRecord] = []

    for event in events:
        rounded_rt = float(event.reaction_time)
        rt_valid, is_alert, is_drowsy = classify_rounded_rt(rounded_rt)
        window_start = event.deviation_onset_time + PRE_EVENT_START
        window_end = event.deviation_onset_time + PRE_EVENT_END
        eye_contaminated = eye_overlaps_window(
            eye_information, window_start, window_end
        )
        reasons: list[str] = []
        if not rt_valid:
            reasons.append(f"reaction_time<={RT_MIN_VALID:.1f}s")

        start_index = int(round(window_start * sample_rate))
        end_index = start_index + samples_per_window
        within_bounds = start_index >= 0 and end_index <= len(filtered_uv)
        if not within_bounds:
            reasons.append("eeg_window_out_of_bounds")
        if eye_contaminated:
            reasons.append("eye_contaminated")

        psd: np.ndarray | None = None
        log_psd: np.ndarray | None = None
        band_power: dict[str, float] = {}
        band_log_power: dict[str, float] = {}
        if rt_valid and within_bounds and not eye_contaminated:
            segment = filtered_uv[start_index:end_index]
            try:
                event_frequencies, psd = compute_one_sided_psd(segment, sample_rate)
                if not np.allclose(event_frequencies, frequencies):
                    raise ValueError("Inconsistent PSD frequency grid")
                log_psd = 10.0 * np.log10(psd + EPS)
                for band_name, (low_hz, high_hz) in BANDS.items():
                    power = compute_band_power(psd, frequencies, low_hz, high_hz)
                    band_power[band_name] = power
                    band_log_power[band_name] = 10.0 * math.log10(power + EPS)
            except ValueError as exc:
                reasons.append(f"psd_error:{exc}")
                psd = None
                log_psd = None

        used = psd is not None and log_psd is not None
        records.append(
            EventPSDRecord(
                event=event,
                rt_valid=rt_valid,
                rt_bin=rt_bin_label(rounded_rt) if rt_valid else None,
                is_alert=is_alert,
                is_drowsy=is_drowsy,
                window_start_time=window_start,
                window_end_time=window_end,
                eye_contaminated=eye_contaminated,
                used_for_eeg_analysis=used,
                exclusion_reason=";".join(reasons),
                psd=psd,
                log_psd=log_psd,
                band_power=band_power,
                band_log_power=band_log_power,
            )
        )

    early_drowsy_n, first_valid_drowsy_time = assign_figure_b_groups(records)

    return RecordingAnalysis(
        edf_path=edf_path,
        channel_name=resolved_channel,
        sample_rate=sample_rate,
        source_unit=source_unit,
        eye_information=eye_information,
        frequencies=frequencies,
        records=tuple(records),
        early_drowsy_n=early_drowsy_n,
        first_valid_drowsy_time=first_valid_drowsy_time,
    )


def frequency_column(prefix: str, frequency: float) -> str:
    return f"{prefix}_{frequency:.6f}Hz"


def event_level_dataframe(analysis: RecordingAnalysis) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    frequencies = analysis.frequencies
    for record in analysis.records:
        event = record.event
        row: dict[str, object] = {
            "event_index": event.event_index,
            "deviation_status": event.deviation_status,
            "deviation_onset_time": event.deviation_onset_time,
            "response_onset_time": event.response_onset_time,
            "response_end_time": event.response_end_time,
            "reaction_time_raw": event.reaction_time_raw,
            "reaction_time": event.reaction_time,
            "rt_valid": record.rt_valid,
            "eye_contaminated": record.eye_contaminated,
            "used_for_eeg_analysis": record.used_for_eeg_analysis,
            "exclusion_reason": record.exclusion_reason,
            "window_start_time": record.window_start_time,
            "window_end_time": record.window_end_time,
            "rt_bin": record.rt_bin,
            "is_alert": record.is_alert,
            "is_drowsy": record.is_drowsy,
            "is_early_drowsy": record.is_early_drowsy,
            "is_selected_alert": record.is_selected_alert,
        }
        for band_name in BANDS:
            row[f"{band_name}_power_uV2"] = record.band_power.get(band_name, np.nan)
            row[f"{band_name}_log_power_dB"] = record.band_log_power.get(
                band_name, np.nan
            )
        for index, frequency in enumerate(frequencies):
            row[frequency_column("PSD_uV2_per_Hz", frequency)] = (
                float(record.psd[index]) if record.psd is not None else np.nan
            )
            row[frequency_column("logPSD_dB", frequency)] = (
                float(record.log_psd[index])
                if record.log_psd is not None
                else np.nan
            )
        rows.append(row)
    return pd.DataFrame(rows)


def metadata_rows(analysis: RecordingAnalysis) -> list[tuple[str, object]]:
    records = analysis.records
    return [
        ("recording", analysis.edf_path.stem),
        ("edf_path", str(analysis.edf_path.resolve())),
        ("channel", analysis.channel_name),
        ("edf_source_unit", analysis.source_unit),
        ("analysis_unit", "microvolts"),
        ("sampling_rate_hz", analysis.sample_rate),
        ("event_window", f"[{PRE_EVENT_START}, {PRE_EVENT_END}) s relative to deviation"),
        ("rt_rounding", "ROUND_HALF_UP to 0.1 s before exclusion/classification/binning"),
        ("rt_exclusion", f"RT <= {RT_MIN_VALID:.1f} s"),
        ("alert_definition", f"{RT_MIN_VALID:.1f} < RT < {RT_THRESHOLD:.1f} s"),
        ("drowsy_definition", f"RT >= {RT_THRESHOLD:.1f} s"),
        ("eeg_filter", f"separate order-{EEG_FILTER_ORDER} Butterworth sosfiltfilt, {EEG_HIGHPASS_HZ}-{EEG_LOWPASS_HZ} Hz"),
        ("psd", f"one-sided FFT density, {PSD_WINDOW}, detrend={PSD_DETREND}, nfft={PSD_NFFT}"),
        ("psd_unit", "uV^2/Hz; log PSD is dB re 1 uV^2/Hz"),
        ("nominal_psd_frequency_range_hz", f"{PSD_FREQ_MIN}-{PSD_FREQ_MAX}"),
        ("actual_first_frequency_bin_hz", float(analysis.frequencies[0])),
        ("actual_last_frequency_bin_hz", float(analysis.frequencies[-1])),
        ("eye_source", analysis.eye_information.source),
        (
            "eye_source_path",
            str(analysis.eye_information.source_path.resolve())
            if analysis.eye_information.source_path
            else None,
        ),
        ("eye_event_seconds", len(analysis.eye_information.seconds)),
        ("parsed_events", len(records)),
        ("valid_rt_events", sum(record.rt_valid for record in records)),
        ("eye_contaminated_events", sum(record.eye_contaminated for record in records)),
        (
            "used_for_eeg_analysis",
            sum(record.used_for_eeg_analysis for record in records),
        ),
        (
            "valid_drowsy_events",
            sum(record.used_for_eeg_analysis and record.is_drowsy for record in records),
        ),
        (
            "drowsy_selection",
            f"earliest {DROWSY_EARLY_FRACTION:.0%} of EEG-valid Drowsy events",
        ),
        ("early_drowsy_n", analysis.early_drowsy_n),
        ("first_valid_drowsy_time", analysis.first_valid_drowsy_time),
        (
            "selected_alert_n",
            sum(record.is_selected_alert for record in records),
        ),
    ]


def metadata_dataframe(analysis: RecordingAnalysis) -> pd.DataFrame:
    return pd.DataFrame(metadata_rows(analysis), columns=["parameter", "value"])


def style_excel_writer(writer: pd.ExcelWriter) -> None:
    for worksheet in writer.book.worksheets:
        worksheet.freeze_panes = "A2"
        if worksheet.max_row >= 1 and worksheet.max_column >= 1:
            worksheet.auto_filter.ref = worksheet.dimensions
        for cells in worksheet.iter_cols(
            min_col=1,
            max_col=min(worksheet.max_column, 20),
            min_row=1,
            max_row=min(worksheet.max_row, 200),
        ):
            width = min(
                max((len(str(cell.value)) if cell.value is not None else 0) for cell in cells)
                + 2,
                38,
            )
            worksheet.column_dimensions[cells[0].column_letter].width = max(width, 10)


def write_event_level_workbook(analysis: RecordingAnalysis, output_path: Path) -> None:
    event_frame = event_level_dataframe(analysis)
    frequency_frame = pd.DataFrame(
        {
            "frequency_index": np.arange(len(analysis.frequencies)),
            "frequency_hz": analysis.frequencies,
        }
    )
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        event_frame.to_excel(writer, sheet_name="events", index=False)
        frequency_frame.to_excel(writer, sheet_name="frequencies", index=False)
        metadata_dataframe(analysis).to_excel(writer, sheet_name="metadata", index=False)
        style_excel_writer(writer)


def build_figure_a_matrix(
    analysis: RecordingAnalysis,
) -> tuple[list[str], np.ndarray, np.ndarray]:
    labels = heatmap_bin_labels()
    matrix = np.full((len(labels), len(analysis.frequencies)), np.nan, dtype=float)
    counts = np.zeros(len(labels), dtype=int)
    for row_index, label in enumerate(labels):
        matching = [
            record.log_psd
            for record in analysis.records
            if record.used_for_eeg_analysis
            and record.rt_bin == label
            and record.log_psd is not None
        ]
        counts[row_index] = len(matching)
        if matching:
            matrix[row_index] = np.mean(np.vstack(matching), axis=0)
    return labels, counts, matrix


def write_figure_a_workbook(
    analysis: RecordingAnalysis,
    labels: list[str],
    counts: np.ndarray,
    matrix: np.ndarray,
    output_path: Path,
) -> None:
    data: dict[str, object] = {
        "rt_bin": labels,
        "event_count": counts,
    }
    for index, frequency in enumerate(analysis.frequencies):
        data[frequency_column("mean_logPSD_dB", frequency)] = matrix[:, index]
    matrix_frame = pd.DataFrame(data)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        matrix_frame.to_excel(writer, sheet_name="rt_frequency_matrix", index=False)
        metadata_dataframe(analysis).to_excel(writer, sheet_name="metadata", index=False)
        style_excel_writer(writer)


def frequency_edges(frequencies: np.ndarray) -> np.ndarray:
    if frequencies.size < 2:
        return np.asarray([frequencies[0] - 0.5, frequencies[0] + 0.5])
    half_step = float(frequencies[1] - frequencies[0]) / 2.0
    return np.r_[frequencies - half_step, frequencies[-1] + half_step]


def save_figure_a(
    analysis: RecordingAnalysis,
    labels: list[str],
    counts: np.ndarray,
    matrix: np.ndarray,
    output_path: Path,
) -> None:
    fig, ax = plt.subplots(figsize=(13, 10))
    finite_values = matrix[np.isfinite(matrix)]
    if finite_values.size:
        image = ax.pcolormesh(
            frequency_edges(analysis.frequencies),
            np.arange(len(labels) + 1),
            np.ma.masked_invalid(matrix),
            shading="flat",
            cmap="viridis",
        )
        colorbar = fig.colorbar(image, ax=ax, pad=0.02)
        colorbar.set_label("Mean Log PSD (dB re 1 μV²/Hz)")
    else:
        ax.text(
            0.5,
            0.5,
            "No eye-clean events with RT > 0.2 s",
            transform=ax.transAxes,
            ha="center",
            va="center",
        )

    tick_indices = list(range(0, len(labels), 5))
    if len(labels) - 1 not in tick_indices:
        tick_indices.append(len(labels) - 1)
    ax.set_yticks([index + 0.5 for index in tick_indices])
    ax.set_yticklabels(
        [f"{labels[index]} (n={counts[index]})" for index in tick_indices]
    )
    ax.set_ylim(0, len(labels))
    ax.set_xlim(analysis.frequencies[0], analysis.frequencies[-1])
    for boundary in sorted({value for band in BANDS.values() for value in band}):
        ax.axvline(boundary, color="white", linestyle=":", linewidth=0.7, alpha=0.65)
    ax.set_xlabel("Frequency (Hz)")
    ax.set_ylabel("Rounded Reaction Time bin (s)")
    ax.set_title(f"{analysis.edf_path.stem} — Figure A: RT × Frequency Heatmap")
    fig.tight_layout()
    fig.savefig(output_path, dpi=FIGURE_DPI)
    plt.close(fig)


def selected_group_records(
    analysis: RecordingAnalysis,
) -> tuple[list[EventPSDRecord], list[EventPSDRecord]]:
    alerts = [record for record in analysis.records if record.is_selected_alert]
    early_drowsy = [record for record in analysis.records if record.is_early_drowsy]
    return alerts, early_drowsy


def mean_log_spectrum(records: Sequence[EventPSDRecord], size: int) -> np.ndarray:
    values = [record.log_psd for record in records if record.log_psd is not None]
    if not values:
        return np.full(size, np.nan, dtype=float)
    return np.mean(np.vstack(values), axis=0)


def summarize_values(values: Iterable[float]) -> tuple[int, float, float, float]:
    array = np.asarray(list(values), dtype=float)
    count = int(array.size)
    if count == 0:
        return 0, np.nan, np.nan, np.nan
    mean = float(np.mean(array))
    standard_deviation = float(np.std(array, ddof=1)) if count > 1 else np.nan
    standard_error = standard_deviation / math.sqrt(count) if count > 1 else np.nan
    return count, mean, standard_deviation, standard_error


def build_band_summary(
    alerts: Sequence[EventPSDRecord], early_drowsy: Sequence[EventPSDRecord]
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for band_name, (low_hz, high_hz) in BANDS.items():
        alert_stats = summarize_values(
            record.band_log_power[band_name]
            for record in alerts
            if band_name in record.band_log_power
        )
        drowsy_stats = summarize_values(
            record.band_log_power[band_name]
            for record in early_drowsy
            if band_name in record.band_log_power
        )
        rows.append(
            {
                "band": band_name.capitalize(),
                "low_hz": low_hz,
                "high_hz": high_hz,
                "alert_n": alert_stats[0],
                "alert_mean_log_power_dB": alert_stats[1],
                "alert_sd_dB": alert_stats[2],
                "alert_sem_dB": alert_stats[3],
                "early_drowsy_n": drowsy_stats[0],
                "early_drowsy_mean_log_power_dB": drowsy_stats[1],
                "early_drowsy_sd_dB": drowsy_stats[2],
                "early_drowsy_sem_dB": drowsy_stats[3],
                "early_drowsy_minus_alert_dB": (
                    drowsy_stats[1] - alert_stats[1]
                    if math.isfinite(drowsy_stats[1]) and math.isfinite(alert_stats[1])
                    else np.nan
                ),
            }
        )
    return pd.DataFrame(rows)


def figure_b_selected_events_dataframe(
    alerts: Sequence[EventPSDRecord], early_drowsy: Sequence[EventPSDRecord]
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for group_name, records in (("Alert", alerts), ("Early Drowsy", early_drowsy)):
        for record in records:
            rows.append(
                {
                    "group": group_name,
                    "event_index": record.event.event_index,
                    "deviation_onset_time": record.event.deviation_onset_time,
                    "reaction_time": record.event.reaction_time,
                    "rt_bin": record.rt_bin,
                }
            )
    return pd.DataFrame(
        rows,
        columns=[
            "group",
            "event_index",
            "deviation_onset_time",
            "reaction_time",
            "rt_bin",
        ],
    )


def write_figure_b_workbook(analysis: RecordingAnalysis, output_path: Path) -> None:
    alerts, early_drowsy = selected_group_records(analysis)
    alert_mean = mean_log_spectrum(alerts, len(analysis.frequencies))
    drowsy_mean = mean_log_spectrum(early_drowsy, len(analysis.frequencies))
    spectrum_frame = pd.DataFrame(
        {
            "frequency_hz": analysis.frequencies,
            "alert_mean_logPSD_dB": alert_mean,
            "early_drowsy_mean_logPSD_dB": drowsy_mean,
            "early_drowsy_minus_alert_dB": drowsy_mean - alert_mean,
        }
    )
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        spectrum_frame.to_excel(writer, sheet_name="mean_spectrum", index=False)
        build_band_summary(alerts, early_drowsy).to_excel(
            writer, sheet_name="band_summary", index=False
        )
        figure_b_selected_events_dataframe(alerts, early_drowsy).to_excel(
            writer, sheet_name="selected_events", index=False
        )
        metadata_dataframe(analysis).to_excel(writer, sheet_name="metadata", index=False)
        style_excel_writer(writer)


def save_figure_b(analysis: RecordingAnalysis, output_path: Path) -> None:
    alerts, early_drowsy = selected_group_records(analysis)
    fig, ax = plt.subplots(figsize=(12, 7))
    band_colors = {"theta": "#8dd3c7", "alpha": "#ffffb3", "beta": "#bebada"}
    for band_name, (low_hz, high_hz) in BANDS.items():
        ax.axvspan(
            low_hz,
            high_hz,
            color=band_colors[band_name],
            alpha=0.18,
            label=f"{band_name.capitalize()} {low_hz:g}–{high_hz:g} Hz",
        )

    if alerts and early_drowsy:
        alert_mean = mean_log_spectrum(alerts, len(analysis.frequencies))
        drowsy_mean = mean_log_spectrum(early_drowsy, len(analysis.frequencies))
        ax.plot(
            analysis.frequencies,
            alert_mean,
            color="#1f77b4",
            linewidth=2.2,
            label=f"Alert (n={len(alerts)})",
        )
        ax.plot(
            analysis.frequencies,
            drowsy_mean,
            color="#d62728",
            linewidth=2.2,
            label=f"Early Drowsy 20% (n={len(early_drowsy)})",
        )
    else:
        missing = []
        if not alerts:
            missing.append("Alert before first valid Drowsy")
        if not early_drowsy:
            missing.append("Early Drowsy")
        ax.text(
            0.5,
            0.5,
            "Figure B unavailable\nMissing: " + ", ".join(missing),
            transform=ax.transAxes,
            ha="center",
            va="center",
        )

    ax.set_xlim(analysis.frequencies[0], analysis.frequencies[-1])
    ax.set_xlabel("Frequency (Hz)")
    ax.set_ylabel("Mean Log PSD (dB re 1 μV²/Hz)")
    ax.set_title(
        f"{analysis.edf_path.stem} — Figure B: Alert vs Early Drowsy (20%) Mean Spectrum"
    )
    ax.grid(True, alpha=0.25)
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(handles, labels, loc="best")
    fig.tight_layout()
    fig.savefig(output_path, dpi=FIGURE_DPI)
    plt.close(fig)


def output_paths(recording: str, output_dir: Path) -> dict[str, Path]:
    return {
        "event_level": output_dir / f"{recording}_event_level_psd.xlsx",
        "figure_a_png": output_dir / f"{recording}_figureA_rt_frequency_heatmap.png",
        "figure_a_xlsx": output_dir / f"{recording}_figureA_rt_frequency_matrix.xlsx",
        "figure_b_png": output_dir / f"{recording}_figureB_alert_vs_early_drowsy.png",
        "figure_b_xlsx": output_dir / f"{recording}_figureB_alert_vs_early_drowsy.xlsx",
    }


def export_recording(
    analysis: RecordingAnalysis,
    output_dir: Path,
    *,
    recording_name: str | None = None,
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = output_paths(recording_name or analysis.edf_path.stem, output_dir)
    labels, counts, matrix = build_figure_a_matrix(analysis)
    write_event_level_workbook(analysis, paths["event_level"])
    write_figure_a_workbook(
        analysis, labels, counts, matrix, paths["figure_a_xlsx"]
    )
    save_figure_a(analysis, labels, counts, matrix, paths["figure_a_png"])
    write_figure_b_workbook(analysis, paths["figure_b_xlsx"])
    save_figure_b(analysis, paths["figure_b_png"])
    return paths


def batch_summary_dataframe(rows: Sequence[dict[str, object]]) -> pd.DataFrame:
    return pd.DataFrame(rows)


def write_batch_summary(rows: Sequence[dict[str, object]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        batch_summary_dataframe(rows).to_excel(writer, sheet_name="batch_summary", index=False)
        style_excel_writer(writer)


def assert_common_frequency_grid(
    analyses: Sequence[tuple[str, RecordingAnalysis]],
) -> np.ndarray:
    if not analyses:
        raise ValueError("No successful recordings are available for aggregation")
    reference = analyses[0][1].frequencies
    for recording_id, analysis in analyses[1:]:
        if reference.shape != analysis.frequencies.shape or not np.allclose(
            reference, analysis.frequencies
        ):
            raise ValueError(
                f"Recording '{recording_id}' has an incompatible PSD frequency grid"
            )
    return reference


def aggregate_metadata_dataframe(
    analyses: Sequence[tuple[str, RecordingAnalysis]],
    *,
    eligible_figure_b_ids: Sequence[str] = (),
) -> pd.DataFrame:
    rows: list[tuple[str, object]] = [
        ("aggregation", "recording-balanced; every recording has equal weight"),
        ("recording_count", len(analyses)),
        ("recording_ids", ",".join(recording_id for recording_id, _ in analyses)),
        ("channel", analyses[0][1].channel_name if analyses else CHANNEL),
        ("rt_exclusion", f"RT <= {RT_MIN_VALID:.1f} s after 0.1-s rounding"),
        ("drowsy_definition", f"RT >= {RT_THRESHOLD:.1f} s"),
        ("early_drowsy_fraction", DROWSY_EARLY_FRACTION),
        ("figure_b_recording_count", len(eligible_figure_b_ids)),
        ("figure_b_recording_ids", ",".join(eligible_figure_b_ids)),
    ]
    return pd.DataFrame(rows, columns=["parameter", "value"])


def build_aggregate_figure_a(
    analyses: Sequence[tuple[str, RecordingAnalysis]],
) -> tuple[
    np.ndarray,
    list[str],
    np.ndarray,
    np.ndarray,
    np.ndarray,
    np.ndarray,
    pd.DataFrame,
]:
    """Average within recording/RT-bin first, then equally across recordings."""

    frequencies = assert_common_frequency_grid(analyses)
    labels = heatmap_bin_labels()
    recording_matrices: list[np.ndarray] = []
    recording_counts: list[np.ndarray] = []
    count_rows: list[dict[str, object]] = []
    for recording_id, analysis in analyses:
        analysis_labels, counts, matrix = build_figure_a_matrix(analysis)
        if analysis_labels != labels:
            raise ValueError(f"RT-bin mismatch in recording '{recording_id}'")
        recording_matrices.append(matrix)
        recording_counts.append(counts)
        for label, count in zip(labels, counts):
            count_rows.append(
                {
                    "recording_id": recording_id,
                    "rt_bin": label,
                    "event_count": int(count),
                }
            )

    matrix_stack = np.stack(recording_matrices)
    count_stack = np.stack(recording_counts)
    mean_matrix = np.full(matrix_stack.shape[1:], np.nan, dtype=float)
    sem_matrix = np.full(matrix_stack.shape[1:], np.nan, dtype=float)
    recording_count = np.sum(count_stack > 0, axis=0).astype(int)
    event_count = np.sum(count_stack, axis=0).astype(int)
    for bin_index in range(len(labels)):
        valid_rows = matrix_stack[count_stack[:, bin_index] > 0, bin_index, :]
        if valid_rows.size == 0:
            continue
        mean_matrix[bin_index] = np.mean(valid_rows, axis=0)
        if valid_rows.shape[0] > 1:
            sem_matrix[bin_index] = np.std(valid_rows, axis=0, ddof=1) / math.sqrt(
                valid_rows.shape[0]
            )
    return (
        frequencies,
        labels,
        recording_count,
        event_count,
        mean_matrix,
        sem_matrix,
        pd.DataFrame(count_rows),
    )


def matrix_dataframe(
    labels: Sequence[str],
    recording_count: np.ndarray,
    event_count: np.ndarray,
    matrix: np.ndarray,
    frequencies: np.ndarray,
    *,
    prefix: str,
) -> pd.DataFrame:
    data: dict[str, object] = {
        "rt_bin": list(labels),
        "recording_count": recording_count,
        "event_count": event_count,
    }
    for index, frequency in enumerate(frequencies):
        data[frequency_column(prefix, frequency)] = matrix[:, index]
    return pd.DataFrame(data)


def save_aggregate_figure_a(
    frequencies: np.ndarray,
    labels: Sequence[str],
    recording_count: np.ndarray,
    event_count: np.ndarray,
    matrix: np.ndarray,
    recording_total: int,
    output_path: Path,
) -> None:
    fig, ax = plt.subplots(figsize=(13, 10))
    finite_values = matrix[np.isfinite(matrix)]
    if finite_values.size:
        image = ax.pcolormesh(
            frequency_edges(frequencies),
            np.arange(len(labels) + 1),
            np.ma.masked_invalid(matrix),
            shading="flat",
            cmap="viridis",
        )
        colorbar = fig.colorbar(image, ax=ax, pad=0.02)
        colorbar.set_label("Recording-balanced Mean Log PSD (dB re 1 μV²/Hz)")
    else:
        ax.text(0.5, 0.5, "No valid RT-bin spectra", transform=ax.transAxes, ha="center")
    tick_indices = list(range(0, len(labels), 5))
    if len(labels) - 1 not in tick_indices:
        tick_indices.append(len(labels) - 1)
    ax.set_yticks([index + 0.5 for index in tick_indices])
    ax.set_yticklabels(
        [
            f"{labels[index]} (R={recording_count[index]}, E={event_count[index]})"
            for index in tick_indices
        ]
    )
    ax.set_ylim(0, len(labels))
    ax.set_xlim(frequencies[0], frequencies[-1])
    for boundary in sorted({value for band in BANDS.values() for value in band}):
        ax.axvline(boundary, color="white", linestyle=":", linewidth=0.7, alpha=0.65)
    ax.set_xlabel("Frequency (Hz)")
    ax.set_ylabel("Rounded Reaction Time bin (s)")
    ax.set_title(
        f"Train Aggregate (n={recording_total}) — Figure A: RT × Frequency Heatmap"
    )
    fig.tight_layout()
    fig.savefig(output_path, dpi=FIGURE_DPI)
    plt.close(fig)


def recording_mean_band_log_power(
    records: Sequence[EventPSDRecord], band_name: str
) -> float:
    values = [
        record.band_log_power[band_name]
        for record in records
        if band_name in record.band_log_power
    ]
    return float(np.mean(values)) if values else np.nan


def mean_and_sem_axis_zero(values: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    if values.ndim != 2 or values.shape[0] == 0:
        raise ValueError("Expected a non-empty recordings × frequencies matrix")
    mean = np.mean(values, axis=0)
    sem = (
        np.std(values, axis=0, ddof=1) / math.sqrt(values.shape[0])
        if values.shape[0] > 1
        else np.full(values.shape[1], np.nan, dtype=float)
    )
    return mean, sem


def build_aggregate_figure_b(
    analyses: Sequence[tuple[str, RecordingAnalysis]],
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    list[str],
]:
    frequencies = assert_common_frequency_grid(analyses)
    eligible: list[
        tuple[str, list[EventPSDRecord], list[EventPSDRecord], np.ndarray, np.ndarray]
    ] = []
    count_rows: list[dict[str, object]] = []
    for recording_id, analysis in analyses:
        alerts, early_drowsy = selected_group_records(analysis)
        valid_drowsy_n = sum(
            record.used_for_eeg_analysis and record.is_drowsy
            for record in analysis.records
        )
        eligible_for_figure_b = bool(alerts and early_drowsy)
        count_rows.append(
            {
                "recording_id": recording_id,
                "alert_event_count": len(alerts),
                "valid_drowsy_event_count": valid_drowsy_n,
                "early_drowsy_event_count": len(early_drowsy),
                "included_in_aggregate_figure_b": eligible_for_figure_b,
            }
        )
        if eligible_for_figure_b:
            eligible.append(
                (
                    recording_id,
                    alerts,
                    early_drowsy,
                    mean_log_spectrum(alerts, len(frequencies)),
                    mean_log_spectrum(early_drowsy, len(frequencies)),
                )
            )
    if not eligible:
        raise ValueError("No recording has both selected Alert and Early Drowsy events")

    alert_stack = np.vstack([item[3] for item in eligible])
    drowsy_stack = np.vstack([item[4] for item in eligible])
    delta_stack = drowsy_stack - alert_stack
    alert_mean, alert_sem = mean_and_sem_axis_zero(alert_stack)
    drowsy_mean, drowsy_sem = mean_and_sem_axis_zero(drowsy_stack)
    delta_mean, delta_sem = mean_and_sem_axis_zero(delta_stack)
    spectrum_frame = pd.DataFrame(
        {
            "frequency_hz": frequencies,
            "alert_recording_mean_logPSD_dB": alert_mean,
            "alert_recording_sem_dB": alert_sem,
            "early_drowsy_recording_mean_logPSD_dB": drowsy_mean,
            "early_drowsy_recording_sem_dB": drowsy_sem,
            "early_drowsy_minus_alert_mean_dB": delta_mean,
            "early_drowsy_minus_alert_sem_dB": delta_sem,
        }
    )
    per_recording_spectrum: dict[str, object] = {"frequency_hz": frequencies}
    for recording_id, _, _, alert_spectrum, drowsy_spectrum in eligible:
        per_recording_spectrum[f"{recording_id}_alert"] = alert_spectrum
        per_recording_spectrum[f"{recording_id}_early_drowsy"] = drowsy_spectrum

    band_rows: list[dict[str, object]] = []
    for recording_id, alerts, early_drowsy, _, _ in eligible:
        for band_name, (low_hz, high_hz) in BANDS.items():
            alert_value = recording_mean_band_log_power(alerts, band_name)
            drowsy_value = recording_mean_band_log_power(early_drowsy, band_name)
            band_rows.append(
                {
                    "recording_id": recording_id,
                    "band": band_name.capitalize(),
                    "low_hz": low_hz,
                    "high_hz": high_hz,
                    "alert_mean_log_power_dB": alert_value,
                    "early_drowsy_mean_log_power_dB": drowsy_value,
                    "early_drowsy_minus_alert_dB": drowsy_value - alert_value,
                }
            )
    band_by_recording = pd.DataFrame(band_rows)
    aggregate_band_rows: list[dict[str, object]] = []
    for band_name in (name.capitalize() for name in BANDS):
        selected = band_by_recording[band_by_recording["band"] == band_name]
        alert_stats = summarize_values(selected["alert_mean_log_power_dB"])
        drowsy_stats = summarize_values(selected["early_drowsy_mean_log_power_dB"])
        delta_stats = summarize_values(selected["early_drowsy_minus_alert_dB"])
        aggregate_band_rows.append(
            {
                "band": band_name,
                "recording_count": alert_stats[0],
                "alert_mean_dB": alert_stats[1],
                "alert_sd_dB": alert_stats[2],
                "alert_sem_dB": alert_stats[3],
                "early_drowsy_mean_dB": drowsy_stats[1],
                "early_drowsy_sd_dB": drowsy_stats[2],
                "early_drowsy_sem_dB": drowsy_stats[3],
                "difference_mean_dB": delta_stats[1],
                "difference_sd_dB": delta_stats[2],
                "difference_sem_dB": delta_stats[3],
            }
        )
    return (
        spectrum_frame,
        pd.DataFrame(per_recording_spectrum),
        pd.DataFrame(count_rows),
        pd.DataFrame(aggregate_band_rows),
        band_by_recording,
        [item[0] for item in eligible],
    )


def save_aggregate_figure_b(
    spectrum_frame: pd.DataFrame,
    eligible_recording_ids: Sequence[str],
    output_path: Path,
) -> None:
    frequencies = spectrum_frame["frequency_hz"].to_numpy(dtype=float)
    alert_mean = spectrum_frame["alert_recording_mean_logPSD_dB"].to_numpy(dtype=float)
    alert_sem = spectrum_frame["alert_recording_sem_dB"].to_numpy(dtype=float)
    drowsy_mean = spectrum_frame[
        "early_drowsy_recording_mean_logPSD_dB"
    ].to_numpy(dtype=float)
    drowsy_sem = spectrum_frame["early_drowsy_recording_sem_dB"].to_numpy(dtype=float)
    fig, ax = plt.subplots(figsize=(12, 7))
    band_colors = {"theta": "#8dd3c7", "alpha": "#ffffb3", "beta": "#bebada"}
    for band_name, (low_hz, high_hz) in BANDS.items():
        ax.axvspan(
            low_hz,
            high_hz,
            color=band_colors[band_name],
            alpha=0.18,
            label=f"{band_name.capitalize()} {low_hz:g}–{high_hz:g} Hz",
        )
    ax.plot(frequencies, alert_mean, color="#1f77b4", linewidth=2.2, label="Alert")
    ax.plot(
        frequencies,
        drowsy_mean,
        color="#d62728",
        linewidth=2.2,
        label="Early Drowsy 20%",
    )
    if np.any(np.isfinite(alert_sem)):
        ax.fill_between(
            frequencies,
            alert_mean - alert_sem,
            alert_mean + alert_sem,
            color="#1f77b4",
            alpha=0.18,
        )
    if np.any(np.isfinite(drowsy_sem)):
        ax.fill_between(
            frequencies,
            drowsy_mean - drowsy_sem,
            drowsy_mean + drowsy_sem,
            color="#d62728",
            alpha=0.18,
        )
    ax.set_xlim(frequencies[0], frequencies[-1])
    ax.set_xlabel("Frequency (Hz)")
    ax.set_ylabel("Recording-balanced Mean Log PSD (dB re 1 μV²/Hz)")
    ax.set_title(
        "Train Aggregate — Figure B: Alert vs Early Drowsy (20%) "
        f"(recordings n={len(eligible_recording_ids)})"
    )
    ax.grid(True, alpha=0.25)
    ax.legend(loc="best")
    fig.tight_layout()
    fig.savefig(output_path, dpi=FIGURE_DPI)
    plt.close(fig)


def aggregate_output_paths(output_dir: Path) -> dict[str, Path]:
    return {
        "event_level": output_dir / "train_event_level_psd.xlsx",
        "figure_a_png": output_dir / "train_figureA_rt_frequency_heatmap.png",
        "figure_a_xlsx": output_dir / "train_figureA_rt_frequency_matrix.xlsx",
        "figure_b_png": output_dir / "train_figureB_alert_vs_early_drowsy.png",
        "figure_b_xlsx": output_dir / "train_figureB_alert_vs_early_drowsy.xlsx",
    }


def export_aggregate(
    analyses: Sequence[tuple[str, RecordingAnalysis]], output_dir: Path
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = aggregate_output_paths(output_dir)
    frequencies = assert_common_frequency_grid(analyses)
    combined_events: list[pd.DataFrame] = []
    for recording_id, analysis in analyses:
        frame = event_level_dataframe(analysis)
        frame.insert(0, "recording_id", recording_id)
        combined_events.append(frame)
    with pd.ExcelWriter(paths["event_level"], engine="openpyxl") as writer:
        pd.concat(combined_events, ignore_index=True).to_excel(
            writer, sheet_name="events", index=False
        )
        pd.DataFrame(
            {
                "frequency_index": np.arange(len(frequencies)),
                "frequency_hz": frequencies,
            }
        ).to_excel(writer, sheet_name="frequencies", index=False)
        aggregate_metadata_dataframe(analyses).to_excel(
            writer, sheet_name="metadata", index=False
        )
        style_excel_writer(writer)

    (
        frequencies,
        labels,
        recording_count,
        event_count,
        mean_matrix,
        sem_matrix,
        bin_counts,
    ) = build_aggregate_figure_a(analyses)
    with pd.ExcelWriter(paths["figure_a_xlsx"], engine="openpyxl") as writer:
        matrix_dataframe(
            labels,
            recording_count,
            event_count,
            mean_matrix,
            frequencies,
            prefix="recording_mean_logPSD_dB",
        ).to_excel(writer, sheet_name="mean_matrix", index=False)
        matrix_dataframe(
            labels,
            recording_count,
            event_count,
            sem_matrix,
            frequencies,
            prefix="recording_sem_dB",
        ).to_excel(writer, sheet_name="sem_matrix", index=False)
        bin_counts.to_excel(writer, sheet_name="recording_bin_counts", index=False)
        aggregate_metadata_dataframe(analyses).to_excel(
            writer, sheet_name="metadata", index=False
        )
        style_excel_writer(writer)
    save_aggregate_figure_a(
        frequencies,
        labels,
        recording_count,
        event_count,
        mean_matrix,
        len(analyses),
        paths["figure_a_png"],
    )

    (
        spectrum_frame,
        per_recording_spectrum,
        selection_counts,
        aggregate_band_summary,
        band_by_recording,
        eligible_ids,
    ) = build_aggregate_figure_b(analyses)
    with pd.ExcelWriter(paths["figure_b_xlsx"], engine="openpyxl") as writer:
        spectrum_frame.to_excel(writer, sheet_name="mean_spectrum", index=False)
        per_recording_spectrum.to_excel(
            writer, sheet_name="recording_spectra", index=False
        )
        selection_counts.to_excel(writer, sheet_name="selection_counts", index=False)
        aggregate_band_summary.to_excel(
            writer, sheet_name="band_summary", index=False
        )
        band_by_recording.to_excel(
            writer, sheet_name="recording_band_values", index=False
        )
        aggregate_metadata_dataframe(
            analyses, eligible_figure_b_ids=eligible_ids
        ).to_excel(writer, sheet_name="metadata", index=False)
        style_excel_writer(writer)
    save_aggregate_figure_b(spectrum_frame, eligible_ids, paths["figure_b_png"])
    return paths


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Create individual and recording-balanced aggregate EEG spectrum "
            "figures for every recording listed in train_data."
        )
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        type=Path,
        help=(
            "Root folder containing the dataset folders listed in train_data; "
            "each folder must contain exactly one *_raw.EDF and eyeblink.dat."
        ),
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=TRAIN_DATA_PATH,
        help=f"Recording manifest (default: {TRAIN_DATA_PATH}).",
    )
    parser.add_argument("--channel", default=CHANNEL, help=f"EEG channel (default: {CHANNEL}).")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    input_dir = args.input_dir.expanduser().resolve()
    if not input_dir.is_dir():
        print(f"Error: input directory does not exist: {input_dir}", file=sys.stderr)
        return 2

    manifest_path = args.manifest.expanduser().resolve()
    output_root = OUTPUT_DIR.resolve()
    summary_path = output_root / "batch_summary.xlsx"
    try:
        training_inputs, summary_rows = preflight_training_inputs(
            input_dir, manifest_path, channel=str(args.channel)
        )
    except Exception as exc:
        print(f"Error: preflight could not start: {exc}", file=sys.stderr)
        return 2

    write_batch_summary(summary_rows, summary_path)
    failed_preflight = [
        row for row in summary_rows if row["preflight_status"] != "READY"
    ]
    print(
        f"Manifest: {manifest_path} ({len(summary_rows)} recordings); "
        f"eye source: eyeblink.dat only"
    )
    if failed_preflight:
        print(
            "Preflight FAILED. No EEG analysis was started because every manifest "
            "recording must pass.",
            file=sys.stderr,
        )
        for row in failed_preflight:
            print(
                f"  {row['recording_id']}: {row['preflight_error']}",
                file=sys.stderr,
            )
        print(f"Preflight report: {summary_path}")
        return 2

    row_by_id = {
        str(row["recording_id"]).casefold(): row for row in summary_rows
    }
    individual_root = output_root / "individual"
    aggregate_root = output_root / "aggregate"
    analyses: list[tuple[str, RecordingAnalysis]] = []
    runtime_failures: list[tuple[str, str]] = []
    for index, training_input in enumerate(training_inputs, start=1):
        recording_id = training_input.recording_id
        row = row_by_id[recording_id.casefold()]
        print(
            f"[{index}/{len(training_inputs)}] Processing {recording_id}: "
            f"{training_input.edf_path.name}"
        )
        try:
            analysis = analyze_recording(
                training_input.edf_path,
                eyeblink_dat_path=training_input.eyeblink_path,
                channel=str(args.channel),
            )
            paths = export_recording(
                analysis,
                individual_root / recording_id,
                recording_name=recording_id,
            )
            used_count = sum(
                record.used_for_eeg_analysis for record in analysis.records
            )
            alert_count = sum(record.is_selected_alert for record in analysis.records)
            drowsy_count = sum(record.is_early_drowsy for record in analysis.records)
            row["analysis_status"] = "SUCCESS"
            row["parsed_events"] = len(analysis.records)
            row["used_for_eeg_analysis"] = used_count
            row["selected_alert_n"] = alert_count
            row["early_drowsy_n"] = drowsy_count
            analyses.append((recording_id, analysis))
            print(
                f"  Done: {used_count}/{len(analysis.records)} EEG-valid events; "
                f"Figure B Alert n={alert_count}, Early Drowsy 20% n={drowsy_count}"
            )
            for path in paths.values():
                print(f"    {path}")
        except Exception as exc:
            row["analysis_status"] = "FAILED"
            row["analysis_error"] = str(exc)
            runtime_failures.append((recording_id, str(exc)))
            print(f"  FAILED: {exc}", file=sys.stderr)

    write_batch_summary(summary_rows, summary_path)
    if runtime_failures:
        print(
            "Aggregate outputs were not generated because the batch was incomplete.",
            file=sys.stderr,
        )
        print(f"Batch summary: {summary_path}")
        return 1

    try:
        aggregate_paths = export_aggregate(analyses, aggregate_root)
    except Exception as exc:
        print(f"Aggregate export FAILED: {exc}", file=sys.stderr)
        print(f"Batch summary: {summary_path}")
        return 1

    print(f"Completed: {len(analyses)}/{len(training_inputs)} recordings")
    print(f"Batch summary: {summary_path}")
    print("Aggregate outputs:")
    for path in aggregate_paths.values():
        print(f"  {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
