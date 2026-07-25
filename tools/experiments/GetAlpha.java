
import com.github.psambit9791.jdsp.transform.FastFourier;

import java.util.Arrays;
//        GetAlpha.AlphaResult result = detectAlpha(dataArray, 250);
//       result.alphaMatrix// Alpha detection per second
//      result.movement


public class GetAlpha {

    public static double[] FFT_AMP_1s(double[] signal, int fs, int nfft) {
        // Ensure signal length matches the expected FFT input size
        double[] signalPadded = new double[nfft];
        System.arraycopy(signal, 0, signalPadded, 0, Math.min(signal.length, nfft));

        // Perform FFT using JDSP
        FastFourier fft = new FastFourier(signalPadded);
        fft.transform();
        double[] fftResult = fft.getMagnitude(false);  // Get the magnitude of FFT
        //Log.d("FFT", String.valueOf(nfft));
        // Compute the two-sided spectrum
        int fftLength = Math.min(fftResult.length, nfft);

        double[] P2 = new double[fftLength];
        for (int i = 0; i < fftLength; i++) {
            P2[i] = fftResult[i] / fs;
        }

        // Compute the one-sided spectrum
        int oneSidedLength = Math.min(fs / 2 + 1, P2.length);  // Ensure we don't exceed array bounds
        double[] P1 = new double[oneSidedLength];
        System.arraycopy(P2, 0, P1, 0, oneSidedLength);

        // Double the amplitude for all except the first and last values
        for (int i = 1; i < P1.length - 1; i++) {
            P1[i] = 2 * P1[i];
        }

        return P1;
    }


    // Main method for alpha detection
    public static AlphaResult detectAlpha(double[] data, int fs) {
        int epoch = fs * 30;
        int nfft = nextPowerOf2(fs);  // Power of 2 for FFT length

        int totalSeconds = data.length / fs;  // Total number of seconds to process
        int[] alphaMatrix = new int[totalSeconds];  // Store alpha detection per second
        double[] ll_eeg = new double[totalSeconds];  // Line length of EEG per second
        double alphaThreshold = 10.0;
        double[] CH1_alpha_power = new double[totalSeconds];  // Alpha power per second
        double[] CH1_all_power = new double[totalSeconds];  // Total power (1-30 Hz) per second

        // Step 1: Line length for artifact detection
        for (int i = 0; i < totalSeconds; i++) {
            int startIdx = i * fs;
            int endIdx = (i + 1) * fs;
            double[] eegWindow = Arrays.copyOfRange(data, startIdx, endIdx);
            ll_eeg[i] = calculateLineLength(eegWindow);
        }

        // Step 2: Compute median line length for artifact detection
        double eegMedian = calculateMedian(ll_eeg);

        // Step 3: Compute FFT for each 1-second window
        for (int i = 0; i < totalSeconds; i++) {
            int startIdx = i * fs;
            int endIdx = (i + 1) * fs;
            double[] eegWindow = Arrays.copyOfRange(data, startIdx, endIdx);

            // Perform FFT
            double[] fftAmp = FFT_AMP_1s(eegWindow, fs, nfft);
            //Log.d("P1", Arrays.toString(fftAmp));
            // Sum alpha band (8-13 Hz)
            double alphaPower = 0;//CH1_alpha_amp
            for (int j = 8; j <= 12; j++) {
                alphaPower += fftAmp[j];
            }
            CH1_alpha_power[i] = alphaPower;
            //Log.d("alphapower", String.valueOf(alphaPower));


            // Sum total power (1-30 Hz)
            double totalPower = 0;//CH1_all_amp
            for (int j = 1; j <= 30; j++) {
                totalPower += fftAmp[j];
            }
            CH1_all_power[i] = totalPower;
            //Log.d("totalpower", String.valueOf(totalPower));


            // Detect alpha if alpha power exceeds threshold
            if (alphaPower >= alphaThreshold) {
                alphaMatrix[i] = 1;  // Mark this second as having alpha waves
            }
        }

        // Step 4: Movement detection (artifact or movement based on line length)
        int[] movement = new int[totalSeconds / 30];
        for (int i = 0; i < movement.length; i++) {
            int movementCount = 0;
            for (int j = 0; j < 30; j++) {
                if (ll_eeg[i * 30 + j] > eegMedian) {
                    movementCount++;
                }
            }
            movement[i] = movementCount;  // Movement detection for each 30s epoch
        }

        return new AlphaResult(alphaMatrix, movement);
    }

    // Helper method to calculate the line length (simple artifact detection)
    public static double calculateLineLength(double[] data) {
        double lineLength = 0;
        for (int i = 1; i < data.length; i++) {
            lineLength += Math.abs(data[i] - data[i - 1]);
        }
        return lineLength;
    }

    // Helper method to calculate the median of an array
    public static double calculateMedian(double[] data) {
        double[] sortedData = Arrays.copyOf(data, data.length);
        Arrays.sort(sortedData);
        if (data.length % 2 == 0) {
            return (sortedData[data.length / 2 - 1] + sortedData[data.length / 2]) / 2.0;
        } else {
            return sortedData[data.length / 2];
        }
    }

    // Helper method to find the next power of 2 (useful for FFT efficiency)
    public static int nextPowerOf2(int n) {
        return (int) Math.pow(2, Math.ceil(Math.log(n) / Math.log(2)));
    }

    // Class to hold the result (alpha_sec matrix and movement detection)
    public static class AlphaResult {
        public int[] alphaMatrix;  // Alpha detection per second
        public int[] movement;  // Movement detection per 30-second epoch

        public AlphaResult(int[] alphaMatrix, int[] movement) {
            this.alphaMatrix = alphaMatrix;
            this.movement = movement;
        }
    }

    public static void main(String[] args) {
        // Example usage with synthetic data
        int fs = 250;  // Sampling rate
        double[] data = new double[7500];  // 30 seconds of sample data

        // Fill with synthetic EEG data (for demonstration)
        for (int i = 0; i < data.length; i++) {
            data[i] = Math.sin(2.0 * Math.PI * 10 * i / fs);  // Example 10 Hz sine wave
        }

        // Detect alpha waves and movement
        AlphaResult result = detectAlpha(data, fs);
//
//        System.out.println("Alpha detection matrix: " + Arrays.toString(result.alphaMatrix));
//        System.out.println("Movement detection: " + Arrays.toString(result.movement));
    }
}
