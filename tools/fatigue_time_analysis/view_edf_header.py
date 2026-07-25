"""Display the fixed header and per-signal headers of an EDF file.

This script uses only Python's standard library.  It reads the EDF header
directly, so no package installation is required.

Examples
--------
    python view_edf_header.py
    python tools/fatigue_time_analysis/view_edf_header.py data/raw_edf/auxiliary/s01_060926_1n_car_position.EDF
"""

from __future__ import annotations

import argparse
import sys
from array import array
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CANDIDATES = (
    PROJECT_ROOT / "data/raw_edf/auxiliary/s01_060926_1n_car_position.EDF",
    PROJECT_ROOT / "EDF/s01_060926_1n_car_position.EDF",
)
DEFAULT_FILE = next(
    (candidate for candidate in DEFAULT_CANDIDATES if candidate.is_file()),
    DEFAULT_CANDIDATES[0],
)

# EDF stores each signal-header field for all signals together, in this order.
SIGNAL_FIELDS: tuple[tuple[str, int], ...] = (
    ("label", 16),
    ("transducer", 80),
    ("physical_dimension", 8),
    ("physical_min", 8),
    ("physical_max", 8),
    ("digital_min", 8),
    ("digital_max", 8),
    ("prefilter", 80),
    ("samples_per_record", 8),
    ("reserved", 32),
)


def decode_field(data: bytes) -> str:
    """Decode one fixed-width EDF ASCII field."""
    return data.decode("ascii", errors="replace").strip()


def as_int(value: str, name: str) -> int:
    try:
        return int(value)
    except ValueError as error:
        raise ValueError(f"Invalid EDF {name!r} field: {value!r}") from error


def as_float(value: str, name: str) -> float:
    try:
        return float(value)
    except ValueError as error:
        raise ValueError(f"Invalid EDF {name!r} field: {value!r}") from error


def read_edf_header(path: Path) -> tuple[dict[str, object], list[dict[str, object]]]:
    """Read EDF fixed and signal headers without reading the signal samples."""
    with path.open("rb") as file:
        fixed = file.read(256)
        if len(fixed) != 256:
            raise ValueError("The file is shorter than the 256-byte EDF fixed header.")

        fixed_header = {
            "version": decode_field(fixed[0:8]),
            "patient": decode_field(fixed[8:88]),
            "recording": decode_field(fixed[88:168]),
            "start_date": decode_field(fixed[168:176]),
            "start_time": decode_field(fixed[176:184]),
            "header_bytes": as_int(decode_field(fixed[184:192]), "header_bytes"),
            "reserved": decode_field(fixed[192:236]),
            "data_records": as_int(decode_field(fixed[236:244]), "data_records"),
            "record_duration_seconds": as_float(
                decode_field(fixed[244:252]), "record_duration_seconds"
            ),
            "signals": as_int(decode_field(fixed[252:256]), "signals"),
        }

        header_bytes = int(fixed_header["header_bytes"])
        signal_count = int(fixed_header["signals"])
        expected_header_bytes = 256 * (signal_count + 1)
        if header_bytes < 256:
            raise ValueError(f"Invalid header size: {header_bytes} bytes.")

        signal_data = file.read(header_bytes - 256)
        if len(signal_data) != header_bytes - 256:
            raise ValueError("The file ended before its complete EDF header was read.")

    if header_bytes != expected_header_bytes:
        print(
            "Warning: EDF header size is "
            f"{header_bytes}, but {signal_count} signals normally use "
            f"{expected_header_bytes} bytes."
        )

    signal_headers: list[dict[str, object]] = [
        {"index": index} for index in range(signal_count)
    ]
    offset = 0
    for field_name, field_width in SIGNAL_FIELDS:
        for index in range(signal_count):
            start = offset + index * field_width
            end = start + field_width
            signal_headers[index][field_name] = decode_field(signal_data[start:end])
        offset += signal_count * field_width

    duration = float(fixed_header["record_duration_seconds"])
    for signal in signal_headers:
        sample_count = as_int(str(signal["samples_per_record"]), "samples_per_record")
        signal["sample_frequency_hz"] = sample_count / duration if duration else None

    return fixed_header, signal_headers


def physical_value(digital_value: int, signal: dict[str, object]) -> float:
    """Convert one EDF digital value to the signal's physical unit."""
    digital_min = as_int(str(signal["digital_min"]), "digital_min")
    digital_max = as_int(str(signal["digital_max"]), "digital_max")
    physical_min = as_float(str(signal["physical_min"]), "physical_min")
    physical_max = as_float(str(signal["physical_max"]), "physical_max")

    return physical_min + (physical_max - physical_min) * (
        (digital_value - digital_min) / (digital_max - digital_min)
    )


def read_signal_data_ranges(
    path: Path,
    fixed_header: dict[str, object],
    signal_headers: list[dict[str, object]],
) -> None:
    """Scan the EDF samples and add actual ranges to each signal header.

    EDF samples are 16-bit signed integers, stored signal-by-signal inside
    each data record.  Only aggregate statistics are retained, so this works
    without loading the full recording into memory.
    """
    header_bytes = int(fixed_header["header_bytes"])
    samples_per_record = [
        as_int(str(signal["samples_per_record"]), "samples_per_record")
        for signal in signal_headers
    ]
    record_bytes = 2 * sum(samples_per_record)
    data_bytes = path.stat().st_size - header_bytes

    if data_bytes < 0 or data_bytes % record_bytes:
        raise ValueError(
            "The data section does not contain a whole number of EDF data records."
        )

    detected_records = data_bytes // record_bytes
    declared_records = int(fixed_header["data_records"])
    if declared_records >= 0 and declared_records != detected_records:
        print(
            "Warning: header declares "
            f"{declared_records} records, but file size contains {detected_records}."
        )

    stats: list[dict[str, int | None]] = []
    for signal in signal_headers:
        stats.append(
            {
                "minimum": None,
                "maximum": None,
                "digital_min_count": 0,
                "trailing_digital_min_count": 0,
                "non_digital_minimum": None,
                "non_digital_maximum": None,
            }
        )

    with path.open("rb") as file:
        file.seek(header_bytes)
        for _ in range(detected_records):
            record = file.read(record_bytes)
            if len(record) != record_bytes:
                raise ValueError("The file ended while reading an EDF data record.")

            values = array("h")
            values.frombytes(record)
            if sys.byteorder != "little":
                values.byteswap()

            offset = 0
            for index, sample_count in enumerate(samples_per_record):
                signal_values = values[offset : offset + sample_count]
                offset += sample_count
                signal_stats = stats[index]
                header_digital_min = as_int(
                    str(signal_headers[index]["digital_min"]), "digital_min"
                )

                segment_min = min(signal_values)
                segment_max = max(signal_values)
                if signal_stats["minimum"] is None or segment_min < signal_stats["minimum"]:
                    signal_stats["minimum"] = segment_min
                if signal_stats["maximum"] is None or segment_max > signal_stats["maximum"]:
                    signal_stats["maximum"] = segment_max

                for digital_value in signal_values:
                    if digital_value == header_digital_min:
                        signal_stats["digital_min_count"] += 1
                        signal_stats["trailing_digital_min_count"] += 1
                    else:
                        signal_stats["trailing_digital_min_count"] = 0
                        if (
                            signal_stats["non_digital_minimum"] is None
                            or digital_value < signal_stats["non_digital_minimum"]
                        ):
                            signal_stats["non_digital_minimum"] = digital_value
                        if (
                            signal_stats["non_digital_maximum"] is None
                            or digital_value > signal_stats["non_digital_maximum"]
                        ):
                            signal_stats["non_digital_maximum"] = digital_value

    for signal, signal_stats in zip(signal_headers, stats):
        actual_minimum = int(signal_stats["minimum"])
        actual_maximum = int(signal_stats["maximum"])
        signal["actual_digital_min"] = actual_minimum
        signal["actual_digital_max"] = actual_maximum
        signal["actual_physical_min"] = physical_value(actual_minimum, signal)
        signal["actual_physical_max"] = physical_value(actual_maximum, signal)
        signal["header_digital_min_count"] = signal_stats["digital_min_count"]
        signal["trailing_header_digital_min_count"] = signal_stats[
            "trailing_digital_min_count"
        ]

        # Report this alternate range only when every occurrence of the header
        # minimum is at the end of the signal. This avoids silently dropping a
        # legitimate clipped value from the middle of a recording.
        if (
            signal_stats["digital_min_count"]
            == signal_stats["trailing_digital_min_count"]
            and signal_stats["digital_min_count"]
            and signal_stats["non_digital_minimum"] is not None
        ):
            valid_minimum = int(signal_stats["non_digital_minimum"])
            valid_maximum = int(signal_stats["non_digital_maximum"])
            signal["range_excluding_trailing_header_min_digital"] = (
                valid_minimum,
                valid_maximum,
            )
            signal["range_excluding_trailing_header_min_physical"] = (
                physical_value(valid_minimum, signal),
                physical_value(valid_maximum, signal),
            )


def print_header(fixed_header: dict[str, object], signal_headers: list[dict[str, object]]) -> None:
    """Print the parsed header in a readable format."""
    print("=== EDF fixed header ===")
    for name, value in fixed_header.items():
        print(f"{name:24}: {value}")

    print("\n=== Signal headers ===")
    for signal in signal_headers:
        print(f"\n[Channel {signal['index']}] {signal['label']}")
        for name, value in signal.items():
            if name not in {
                "index",
                "label",
                "reserved",
                "actual_digital_min",
                "actual_digital_max",
                "actual_physical_min",
                "actual_physical_max",
                "header_digital_min_count",
                "trailing_header_digital_min_count",
                "range_excluding_trailing_header_min_digital",
                "range_excluding_trailing_header_min_physical",
            }:
                print(f"  {name:22}: {value}")

        if "actual_digital_min" in signal:
            unit = signal["physical_dimension"] or "physical unit"
            print("  actual_stored_data:")
            print(
                "    digital_range       : "
                f"{signal['actual_digital_min']} .. {signal['actual_digital_max']}"
            )
            print(
                "    physical_range      : "
                f"{signal['actual_physical_min']:.6f} .. "
                f"{signal['actual_physical_max']:.6f} {unit}"
            )
            print(
                "    header_min_count    : "
                f"{signal['header_digital_min_count']}"
            )
            print(
                "    trailing_min_count  : "
                f"{signal['trailing_header_digital_min_count']}"
            )
            if "range_excluding_trailing_header_min_digital" in signal:
                digital_range = signal[
                    "range_excluding_trailing_header_min_digital"
                ]
                physical_range = signal[
                    "range_excluding_trailing_header_min_physical"
                ]
                print(
                    "    excluding_trailing_header_min_digital: "
                    f"{digital_range[0]} .. {digital_range[1]}"
                )
                print(
                    "    excluding_trailing_header_min_physical: "
                    f"{physical_range[0]:.6f} .. {physical_range[1]:.6f} {unit}"
                )


def main() -> None:
    parser = argparse.ArgumentParser(description="Read and display an EDF header.")
    parser.add_argument(
        "edf_path",
        nargs="?",
        type=Path,
        default=DEFAULT_FILE,
        help=f"EDF file to inspect (default: {DEFAULT_FILE})",
    )
    parser.add_argument(
        "--header-only",
        action="store_true",
        help="Show only the EDF header; do not scan the signal samples.",
    )
    args = parser.parse_args()

    if not args.edf_path.is_file():
        parser.error(f"EDF file not found: {args.edf_path}")

    fixed_header, signal_headers = read_edf_header(args.edf_path)
    if not args.header_only:
        read_signal_data_ranges(args.edf_path, fixed_header, signal_headers)
    print(f"File: {args.edf_path.resolve()}\n")
    print_header(fixed_header, signal_headers)


if __name__ == "__main__":
    main()
