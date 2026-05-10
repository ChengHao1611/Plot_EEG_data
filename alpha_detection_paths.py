from __future__ import annotations

from pathlib import Path


def normalize_user_path(raw_path: str) -> Path:
    cleaned = raw_path.strip().strip('"').strip("'")
    return Path(cleaned)


def derive_prefix_from_raw_edf(edf_path: Path) -> str:
    stem = edf_path.stem
    if stem.lower().endswith("_raw"):
        return stem[:-4]
    return stem


def find_single_raw_edf(folder: Path) -> Path:
    candidates = sorted(
        path.resolve()
        for path in folder.iterdir()
        if path.is_file() and path.name.lower().endswith("_raw.edf")
    )
    if not candidates:
        raise FileNotFoundError(f"No *_raw.EDF file found in folder: {folder}")
    if len(candidates) > 1:
        joined = "\n".join(str(path) for path in candidates)
        raise ValueError(f"Multiple *_raw.EDF files found in folder:\n{joined}")
    return candidates[0]


def resolve_eye_dat_path(folder: Path, prefix: str) -> Path | None:
    candidates = [
        folder / f"{prefix}_raw_arousal info.dat",
        folder / f"{prefix}_arousal info.dat",
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def resolve_input_paths(input_path: Path) -> tuple[Path, Path, str, Path, Path | None]:
    resolved_input = input_path.resolve()
    if resolved_input.is_dir():
        folder = resolved_input
        edf_path = find_single_raw_edf(folder)
    else:
        if resolved_input.suffix.lower() != ".edf":
            raise ValueError(f"Input must be a folder or EDF file: {resolved_input}")
        folder = resolved_input.parent
        edf_path = resolved_input

    prefix = derive_prefix_from_raw_edf(edf_path)
    alpha_dat_path = folder / f"{prefix}_alpha.dat"
    if not alpha_dat_path.exists():
        raise FileNotFoundError(f"Alpha label file not found: {alpha_dat_path}")

    eye_dat_path = resolve_eye_dat_path(folder, prefix)
    return folder, edf_path, prefix, alpha_dat_path, eye_dat_path


def default_plot_dir(folder: Path, prefix: str) -> Path:
    return folder / f"{prefix}_alpha_detection_plots"


__all__ = [
    "default_plot_dir",
    "derive_prefix_from_raw_edf",
    "find_single_raw_edf",
    "normalize_user_path",
    "resolve_eye_dat_path",
    "resolve_input_paths",
]
