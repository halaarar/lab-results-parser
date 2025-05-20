import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def extract_lab_data(file_path):
    # Read the file
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        content = file.read()
    
    # Split the content into separate lab reports if multiple exist
    reports = re.split(r'(?=RUN DATE:)', content)
    
    # Prepare a list to store extracted data
    data = []
    
    # Process each report
    for report in reports:
        if not report.strip():
            continue
            
        # Extract run date
        run_date_match = re.search(r'RUN DATE:\s*(\d{2}/\d{2}/\d{2})', report)
        run_date = run_date_match.group(1) if run_date_match else ""
        
        # Extract sample ID
        sample_id_match = re.search(r'SPEC #:\s*([^\s]+)', report)
        sample_id = sample_id_match.group(1) if sample_id_match else ""
        
        # Extract age/sex
        age_sex_match = re.search(r'AGE/SEX:\s*([^\s]+)', report)
        age_sex = age_sex_match.group(1) if age_sex_match else ""
        
        # Extract completion date-time
        comp_datetime_match = re.search(r'COMP:\s*(\d{2}/\d{2}/\d{2}-\d{4})', report)
        comp_datetime = comp_datetime_match.group(1) if comp_datetime_match else ""
        
        # Extract all test results
        test_pattern = r'([A-Za-z][A-Za-z\s\d\-]+(?:PCR|Result|Vir|Virus|Panel Interp))\s+Final\s+([^\n]+)'
        test_results = re.findall(test_pattern, report)
        
        # Check if any pathogens were detected
        detected_pathogens = []
        
        for test_name, result in test_results:
            if "Detected" in result and "Not Detected" not in result:
                # Clean up the test name to get the pathogen name
                pathogen = test_name.strip()
                detected_pathogens.append(pathogen)
        
        # Also check for panel interpretation which might have more specific information
        interp_match = re.search(r'Respiratory PCR Panel Interp\s+Final\s+([^\n]+)', report)
        if interp_match:
            interp_text = interp_match.group(1).strip()
            if "DETECTED" in interp_text:
                detected_pathogens.append(interp_text)
        
        # Determine the result
        if detected_pathogens:
            # Use the first detected pathogen as the result
            result = detected_pathogens[0]
        else:
            result = "Not detected"
        
        # Add the data to our list
        data.append({
            'Date': run_date,
            'Sample ID #': sample_id,
            'Age': age_sex,
            'COMP DATE-Time': comp_datetime,
            'Result': result
        })
    
    # Create a DataFrame
    df = pd.DataFrame(data)
    
    # Remove duplicates based on Sample ID
    df = df.drop_duplicates(subset=['Sample ID #'])
    
    return df

def process_file():
    # Ask user to select file
    file_path = filedialog.askopenfilename(
        title="Select the lab report text file",
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
    )
    
    if not file_path:
        return
    
    try:
        df = extract_lab_data(file_path)
        
        # Ask user where to save the Excel file
        output_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if not output_path:
            return
        
        # Export to Excel
        df.to_excel(output_path, index=False)
        
        messagebox.showinfo("Success", f"Data has been extracted and saved to:\n{output_path}")
        
        # Try to open the folder containing the saved file
        try:
            folder_path = os.path.dirname(output_path)
            os.startfile(folder_path)  # For Windows
        except:
            pass  # Silently fail if can't open folder
            
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    # Create the GUI window
    root = tk.Tk()
    root.title("Lab Report Data Extractor")
    root.geometry("400x200")
    
    # Add some padding
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Add a heading
    heading = tk.Label(frame, text="Lab Report Data Extractor", font=("Arial", 14, "bold"))
    heading.pack(pady=(0, 20))
    
    # Add instructions
    instructions = tk.Label(frame, text="Click the button below to select a lab report text file\nand convert it to Excel format.")
    instructions.pack(pady=(0, 20))
    
    # Add the process button
    process_button = tk.Button(frame, text="Select Lab Report File", command=process_file, height=2)
    process_button.pack()
    
    # Start the GUI event loop
    root.mainloop()

if __name__ == "__main__":
    main()