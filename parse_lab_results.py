import re
import pandas as pd
from pathlib import Path

def parse_lab_results(file_path):
    """Parse lab results from the text file and extract relevant information."""
    with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
        content = file.read()
    
    # Create a dictionary to store all information by specimen ID
    specimens_data = {}
    
    # Find all instances of specimen IDs (both SPEC #: and Specimen: formats)
    specimen_matches = re.finditer(r'(?:SPEC #:|Specimen:)\s*(\S+)', content)
    
    for match in specimen_matches:
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
    
    return results

def main():
    # Path to your text file
    file_path = "baba.txt"
    
    # Parse the lab results
    results = parse_lab_results(file_path)
    
    # Create a DataFrame
    df = pd.DataFrame(results)
    
    # Write to Excel
    excel_path = "lab_results.xlsx"
    df.to_excel(excel_path, index=False)
    
    print(f"Results exported to {excel_path}")
    print(f"Found {len(results)} specimens.")
    if results:
        print("First result example:")
        for key, value in results[0].items():
            print(f"{key}: {value}")

if __name__ == "__main__":
    main()