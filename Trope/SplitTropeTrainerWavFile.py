import os
from pydub import AudioSegment

# FORCED PATH FOR WINDOWS STORE PYTHON:
# Tells pydub exactly where ffmpeg.exe is sitting relative to this script
current_dir = os.path.dirname(os.path.abspath(__file__))
AudioSegment.converter = os.path.join(current_dir, "ffmpeg.exe")

# Define the precise file paths based on your project setup
MASTER_WAV = os.path.join(current_dir, "V'ahavta.wav")
# Change this line:
TIMESTAMP_FILE = os.path.join(current_dir, "V'ahavta", "V'ahavta_TS.txt")
OUTPUT_SUBFOLDER = os.path.join(current_dir, "V'ahavta")

def split_vahavta_audio():
    # 1. Double check that the master WAV file is present
    if not os.path.exists(MASTER_WAV):
        print(f"Error: Could not find master file '{MASTER_WAV}' in this folder.")
        return

    # 2. Automatically create the subfolder if it doesn't exist yet
    if not os.path.exists(OUTPUT_SUBFOLDER):
        os.makedirs(OUTPUT_SUBFOLDER)
        print(f"Created missing subfolder: {OUTPUT_SUBFOLDER}")

    # 3. Load the master audio file
    print(f"Loading {MASTER_WAV}...")
    audio = AudioSegment.from_wav(MASTER_WAV)
    
    # 4. Read the timestamps and extract the tracks
    line_number = 1
    with open(TIMESTAMP_FILE, "r") as file:
        for line in file:
            line = line.strip()
            if not line:
                continue # Skip any accidental blank lines
                
            try:
                # Split the line by the comma
                start_sec_str, end_sec_str = line.split(",")
                
                # Convert the decimal seconds directly to floats, then to milliseconds
                start_ms = int(float(start_sec_str.strip()) * 1000)
                end_ms = int(float(end_sec_str.strip()) * 1000)
                
                # Slice the audio segment
                audio_clip = audio[start_ms:end_ms]
                
                # Build the exact destination path: V'ahavta/V'ahavta_lineX.wav
                output_filename = f"V'ahavta_line{line_number}.wav"
                full_output_path = os.path.join(OUTPUT_SUBFOLDER, output_filename)
                
                # Export the uncompressed WAV file
                audio_clip.export(full_output_path, format="wav")
                print(f"Exported: {full_output_path} (Seconds: {start_sec_str.strip()} to {end_sec_str.strip()})")
                
                line_number += 1
                
            except ValueError:
                print(f"Skipping badly formatted line in text file: '{line}'")

    print("\nAudio splitting complete!")
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    split_vahavta_audio()
