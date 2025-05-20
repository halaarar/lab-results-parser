import os
import subprocess
import platform

def build_executable():
    # Determine the command based on the platform
    if platform.system() == "Windows":
        cmd = 'pyinstaller --onefile --windowed --name "Lab Results Parser" --icon=icon.ico lab_results_app.py'
    else:  # macOS, Linux
        cmd = 'pyinstaller --onefile --windowed --name "Lab Results Parser" lab_results_app.py'
    
    # Run the command
    subprocess.call(cmd, shell=True)
    
    # Print completion message
    print("Build completed! Executable can be found in the 'dist' folder.")

if __name__ == "__main__":
    build_executable()