import os
import sys
import argparse
import matplotlib.pyplot as plt
import pandas as pd

def get_args():
    parser = argparse.ArgumentParser(description='plot EEG data from a Excel file')
    parser.add_argument("--file", type=str, help='The path to the Excel file')
    parser.add_argument("--mode", type=int, choices=[1, 2], default=2, help='1:use α/total ratio, 2: use α/β ratio and α/θ ratio')
    return parser.parse_args()

def plot_eeg_data (xlsx_file, mode):
    df = pd.read_excel(xlsx_file)
    eeg_file = os.path.basename(xlsx_file).split('_raw_')[0]
    fig, ax1 = plt.subplots(figsize=(14, 7))

    events = df[df['react_time'] > 0]

    shaded_label_added = False
    for _, row in events.iterrows():
        start_t = row['second']
        duration = row['react_time']
        end_t = start_t + duration

        label = 'Event' if not shaded_label_added else ""
        p = ax1.axvspan(start_t, end_t, color='yellow', alpha=0.3, label=label, zorder = 1)
        if not shaded_label_added:
            event_patch = p
            shaded_label_added = True

    if mode == 2:
        line1, = ax1.plot(df['second'], df['alpha_beta'], label='α/β Ratio', color='blue', alpha=0.7, zorder = 3)
        line2, = ax1.plot(df['second'], df['alpha_theta'], label='α/θ Ratio', color='green', alpha=0.7, zorder = 3)
    else:
        line1, = ax1.plot(df['second'], df['alpha_total'], label='α/total Ratio', color='blue', alpha=0.7, zorder = 3)

    ax1.set_xlabel('Time(s)')
    ax1.set_ylabel('Ratio', color='black')
    ax1.tick_params(axis='y', labelcolor='black')
    ax1.grid(True, which='both', linestyle='--', linewidth=0.5, zorder = 0)

    ax2 = ax1.twinx()
    line3, = ax2.plot(df['second'], df['eyeblinking_count'], label='Eyeblink Count (per 30s)', 
                      color='red', linewidth=2, linestyle='-', zorder = 4)
    ax2.set_ylabel('Eyeblink Count (per 30s)', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    elements = []
    if event_patch:
        elements.append(event_patch)
    if mode == 2:
        elements.extend([line1, line2, line3])
    else:
        elements.extend([line1, line3])
    labels = [obj.get_label() for obj in elements]
    ax1.legend(elements, labels, loc='upper right')
    plt.title(eeg_file)
    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    args = get_args()
    xlsx_file = args.file
    if xlsx_file is None:
        print('Hint: The --file parameter was not detected.')
        xlsx_file = input('Please enter the path to the Excel file: ').strip().replace('"', '').replace("'", "")

    if not os.path.exists(xlsx_file):
        print(f"Error: The file '{xlsx_file}' does not exist.")
        sys.exit(1)

    plot_eeg_data(xlsx_file, args.mode)