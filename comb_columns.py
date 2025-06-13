import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to process the selected Excel file
def process_excel(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)

        # Combine columns E to AX (columns 4 to 49 by index)
        # For each row, join all non-empty values from E to AX into a single string (separated by spaces)
        combined = df.iloc[:, 4:50].apply(
            lambda row: ' '.join(str(cell).strip() for cell in row if pd.notnull(cell) and str(cell).strip()),
            axis=1
        )
        # Store the combined string back in column E (index 4)
        df.iloc[:, 4] = combined

        # Drop columns F to AX (indices 5 to 49)
        df.drop(df.columns[5:50], axis=1, inplace=True)

        # Build output file name by appending '_output' before the extension
        output_path = file_path.replace('.xlsx', '_output.xlsx')
        # Save the resulting DataFrame to a new Excel file
        df.to_excel(output_path, index=False)
        # Show a success message with the output location
        messagebox.showinfo("Success", f"Processed and saved:\n{output_path}")
    except Exception as e:
        # Show an error message if any exception occurs
        messagebox.showerror("Error", str(e))

# Function to prompt the user to select a file and then process it
def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if file_path:
        process_excel(file_path)

# GUI Setup
root = tk.Tk()
root.title("Excel Combiner (E to AX)")
root.geometry("400x200")

# Instruction label
label = tk.Label(root, text="Drag and Drop or Click Below to Upload .xlsx File", wraplength=300, pady=20)
label.pack()

# Button to upload and process file
upload_btn = tk.Button(root, text="Select Excel File", command=upload_file)
upload_btn.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
