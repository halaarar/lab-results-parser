import sys
import os
import pandas as pd
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                            QHBoxLayout, QFileDialog, QLabel, QWidget, QProgressBar, 
                            QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class ParserThread(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(list)
    error_signal = pyqtSignal(str)
    
    def __init__(self, input_file):
        super().__init__()
        self.input_file = input_file
        
    def run(self):
        try:
            results = self.parse_lab_results(self.input_file)
            self.finished_signal.emit(results)
        except Exception as e:
            self.error_signal.emit(str(e))
    
    def parse_lab_results(self, file_path):
        """Parse lab results from the text file and extract relevant information."""
        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            content = file.read()
        
        # Create a dictionary to store all information by specimen ID
        specimens_data = {}
        
        # Find all instances of specimen IDs (both SPEC #: and Specimen: formats)
        specimen_matches = list(re.finditer(r'(?:SPEC #:|Specimen:)\s*(\S+)', content))
        total_matches = len(specimen_matches)
        
        for i, match in enumerate(specimen_matches):
            # Update progress
            self.progress_signal.emit(int((i / total_matches) * 100))
            
            sample_id = match.group(1)
            start_pos = match.start()
            
            # If this is the first time we see this specimen ID, initialize its data
            if sample_id not in specimens_data:
                specimens_data[sample_id] = {
                    'run_date': "Unknown",
                    'age_sex': "Unknown",
                    'comp_date_time': "Unknown",
                    'detected_tests': []
                }
                
            # Extract section from this match to the next one (or end of file)
            section_end = len(content)
            for other_match in re.finditer(r'(?:SPEC #:|Specimen:)\s*(\S+)', content[start_pos + 1:]):
                section_end = start_pos + 1 + other_match.start()
                break
                
            section = content[start_pos:section_end]
            
            # Find RUN DATE near this specimen
            run_date_match = re.search(r'RUN DATE:\s*(\S+)', content[max(0, start_pos-1000):start_pos])
            if not run_date_match:
                run_date_match = re.search(r'RUN DATE:\s*(\S+)', section)
                
            if run_date_match and specimens_data[sample_id]['run_date'] == "Unknown":
                specimens_data[sample_id]['run_date'] = run_date_match.group(1)
                
            # Extract Age/Sex
            age_sex_match = re.search(r'AGE/SEX:\s*(\S+)', content[max(0, start_pos-1000):start_pos])
            if not age_sex_match:
                age_sex_match = re.search(r'AGE/SEX:\s*(\S+)', section)
                
            if age_sex_match and specimens_data[sample_id]['age_sex'] == "Unknown":
                specimens_data[sample_id]['age_sex'] = age_sex_match.group(1)
                
            # Extract Completion Date/Time
            comp_match = re.search(r'COMP:\s*(\S+)', section)
            if comp_match and specimens_data[sample_id]['comp_date_time'] == "Unknown":
                specimens_data[sample_id]['comp_date_time'] = comp_match.group(1)
                
            # Find detected tests in this section
            lines = section.split('\n')
            for i in range(len(lines) - 1):
                line = lines[i].strip()
                
                # Check if this is a test result line (contains "Final")
                if "Final" in line and not line.startswith("---"):
                    # Extract the test name (everything before "Final")
                    test_name = line.split("Final")[0].strip()
                    
                    # Get the next line which should contain the result
                    result_line = lines[i+1].strip() if i+1 < len(lines) else ""
                    
                    # Check if the result is "Detected" (but not "Not Detected")
                    if "Detected" in result_line and "Not Detected" not in result_line:
                        specimens_data[sample_id]['detected_tests'].append(test_name)
        
        # Convert the dictionary to the list of results
        results = []
        for sample_id, data in specimens_data.items():
            # Create result entry
            result = {
                "Date": data['run_date'],
                "Sample ID #": sample_id,
                "Age": data['age_sex'],
                "COMP DATE-Time": data['comp_date_time'],
                "Result": "Not detected" if not data['detected_tests'] else "; ".join([f"{test}: Detected" for test in data['detected_tests']])
            }
            results.append(result)
        
        self.progress_signal.emit(100)
        return results

class LabResultsApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # Set window properties
        self.setWindowTitle('Lab Results Parser')
        self.setGeometry(100, 100, 600, 300)
        
        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # Input file selection
        input_layout = QHBoxLayout()
        self.input_label = QLabel('No file selected')
        self.input_button = QPushButton('Select Input File')
        self.input_button.clicked.connect(self.select_input_file)
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_button)
        main_layout.addLayout(input_layout)
        
        # Output file selection
        output_layout = QHBoxLayout()
        self.output_label = QLabel('No output location selected')
        self.output_button = QPushButton('Select Output Location')
        self.output_button.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_button)
        main_layout.addLayout(output_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        main_layout.addWidget(self.progress_bar)
        
        # Parse button
        self.parse_button = QPushButton('Process File')
        self.parse_button.clicked.connect(self.process_file)
        self.parse_button.setEnabled(False)
        main_layout.addWidget(self.parse_button)
        
        # Status label
        self.status_label = QLabel('Ready')
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        # Help text
        help_text = """
        How to use:
        1. Click 'Select Input File' to choose your lab results text file
        2. Click 'Select Output Location' to choose where to save the Excel file
        3. Click 'Process File' to extract the lab results
        """
        help_label = QLabel(help_text)
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(help_label)
        
        # Initialize variables
        self.input_file = None
        self.output_file = None
        
    def select_input_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Select Lab Results Text File', '', 'Text Files (*.txt)')
        
        if file_path:
            self.input_file = file_path
            self.input_label.setText(os.path.basename(file_path))
            self.check_enable_parse()
            
    def select_output_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getSaveFileName(self, 'Save Excel File', '', 'Excel Files (*.xlsx)')
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_file = file_path
            self.output_label.setText(os.path.basename(file_path))
            self.check_enable_parse()
            
    def check_enable_parse(self):
        if self.input_file and self.output_file:
            self.parse_button.setEnabled(True)
        else:
            self.parse_button.setEnabled(False)
            
    def process_file(self):
        if not self.input_file or not self.output_file:
            return
            
        # Disable buttons during processing
        self.input_button.setEnabled(False)
        self.output_button.setEnabled(False)
        self.parse_button.setEnabled(False)
        self.status_label.setText('Processing...')
        
        # Start processing in a separate thread
        self.thread = ParserThread(self.input_file)
        self.thread.progress_signal.connect(self.update_progress)
        self.thread.finished_signal.connect(self.save_results)
        self.thread.error_signal.connect(self.show_error)
        self.thread.start()
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def save_results(self, results):
        try:
            # Create dataframe and save to Excel
            df = pd.DataFrame(results)
            df.to_excel(self.output_file, index=False)
            
            # Show success message
            QMessageBox.information(self, 'Success', 
                                   f'Successfully processed {len(results)} specimens.\n'
                                   f'Results saved to {self.output_file}')
            
            self.status_label.setText('Completed Successfully')
        except Exception as e:
            self.show_error(str(e))
        finally:
            # Re-enable buttons
            self.input_button.setEnabled(True)
            self.output_button.setEnabled(True)
            self.parse_button.setEnabled(True)
            
    def show_error(self, error_message):
        QMessageBox.critical(self, 'Error', f'An error occurred:\n{error_message}')
        self.status_label.setText('Error')
        
        # Re-enable buttons
        self.input_button.setEnabled(True)
        self.output_button.setEnabled(True)
        self.parse_button.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LabResultsApp()
    window.show()
    sys.exit(app.exec_())