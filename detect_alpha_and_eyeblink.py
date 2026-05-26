import argparse
import os
import sys
import mne
from record_arousal import detect_eye_movements
from record_alpha import detect_alpha

DYNAMIC_PROMINENCE = 0.00006

def get_args():
    parser = argparse.ArgumentParser(description='Detect eye movements from an EDF file')
    parser.add_argument("--file", type=str, help='The path to the EDF file')
    return parser.parse_args()


if __name__ == "__main__":
    args = get_args()
    file_path = args.file
    if file_path is None:
        print('Hint: The --file parameter was not detected.')
        file_path = input('Please enter the path to the EDF file: ').strip().replace('"', '').replace("'", "")
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        sys.exit(1)
    file_dir = os.path.dirname(file_path)
    output_eyeblink_path = os.path.join(file_dir, "eyeblink.dat")
    output_alpha_path = os.path.join(file_dir, "Alpha.dat")
    raw = mne.io.read_raw_edf(file_path, preload=True)

    all_channels = raw.ch_names
    target_channels = [ch for ch in all_channels if 'fp1' in ch.lower() or 'fp2' in ch.lower()]

    if len(target_channels) < 2:
        print("找不到 FP1 或 FP2 通道")
    else:
        detect_eye_movements(raw, target_channels, output_eyeblink_path)
        detect_alpha(file_path, output_alpha_path, target_channels)
