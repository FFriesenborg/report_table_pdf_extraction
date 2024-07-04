import tkinter as tk
from tkinter import filedialog, messagebox
import camelot
import pandas as pd
import os
import openpyxl

import re

def on_submit():
    pdf_file = entry_pdf.get()
    page_specification = entry_page.get()
    csv_file = entry_csv.get()

    try:
        # Split the page specification if it contains a dash
        if '-' in page_specification:
            start_page, end_page = map(int, page_specification.split('-'))
            pages = ','.join(map(str, range(start_page, end_page + 1)))
        
        #if none is specified the default pages are all pages between 3 & 20
        elif page_specification == "":
            pages = ','.join(map(str, range(3, 20 + 1)))
                                
        else:
            pages = page_specification

        wb = openpyxl.Workbook()
    
        for page_number in pages.split(','):
            
            #Specify table areas for pages where header is not detected since it's too far from table content
            #if page_number in ['5','6','9','10','11']:
                #table_areas = ['50,550,700,40']
            #else:
                #table_areas = None
            
            #define table areas so table + header are included in xlsx
            table_areas = ['50,550,740,40']
            
            #define row tolerance to ensure double row headers aren't separated for page 8&9
            if page_number in ['8','9']:
                row_tol = 9
            else:
                row_tol = 2

            #read pdf table
            tables = camelot.read_pdf(pdf_file, flavor='stream', pages=page_number, table_areas = table_areas, row_tol = row_tol)
            for i, table in enumerate(tables, start=1):
                df = table.df

                # Delete all'.' (they are just used as visual separator in German style reporting)
                df = df.applymap(lambda x: x.replace('.', '') if isinstance(x, str) else x)
                
                # Replace all ',' with '.' (to change from german decimal separator ',' to international style decimal separator '.')
                df = df.applymap(lambda x: x.replace(',', '.') if isinstance(x, str) else x)
                
                # Apply transformations to each cell in the DataFrame
                df = df.applymap(lambda x: adjust_negative_number(x))
                # Create a new worksheet for each table
                ws = wb.create_sheet(title=f"Page_{page_number}_Table_{i}")
                for row_data in df.values.tolist():
                    formatted_row_data = []
                    for cell in row_data:
                        try:
                            cell = float(cell)
                            formatted_row_data.append(cell)
                        except ValueError:
                            formatted_row_data.append(cell)
                    ws.append(formatted_row_data)
                # Add empty rows after each table
                for _ in range(5):
                    ws.append([''] * len(df.columns))

        # Remove the default sheet created by openpyxl
        wb.remove(wb["Sheet"])

        # Determine output filename
        if not csv_file.strip():
            base_name = os.path.splitext(os.path.basename(pdf_file))[0]  # Use input PDF filename without extension
            directory = os.path.dirname(pdf_file)
            csv_file = os.path.join(directory, base_name)

        # Check if the file already exists
        if os.path.exists(csv_file + '.xlsx'):
            # File exists, ask for confirmation before overwriting
            confirm = messagebox.askyesno("File Exists", "The output file already exists. Do you want to overwrite it?")
            if not confirm:
                result_label.config(text="Submission canceled.")
                return

        # Save the workbook
        wb.save(csv_file + '.xlsx')
        result_label.config(text=f"All tables extracted and saved to {csv_file}.xlsx")
    except Exception as e:
        result_label.config(text=f"Error: {e}")

#some of the values have unnecessary periods infron of '-'. Those are delted here
def adjust_negative_number(cell_value):
    if isinstance(cell_value, str) and re.match(r'\s*-.*', cell_value.strip()):
        return '-' + re.sub(r'\s+', '', cell_value.strip().lstrip('-'))
    return cell_value


# function for browsing PDF documents in the GUI file browser
def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    entry_pdf.delete(0, tk.END)
    entry_pdf.insert(0, file_path)


#----------------------Graphical User Interface (GUI)------------------------
# Create the main window
root = tk.Tk()
root.title("PDF Table Extractor")

# Create and pack widgets
label_pdf = tk.Label(root, text="PDF File:")
label_pdf.grid(row=0, column=0, padx=10, pady=5, sticky="E")

entry_pdf = tk.Entry(root, width=30)
entry_pdf.grid(row=0, column=1, padx=10, pady=5)

button_browse_pdf = tk.Button(root, text="Browse", command=browse_pdf)
button_browse_pdf.grid(row=0, column=2, pady=5)

label_page = tk.Label(root, text="Page Range: \n (default is pages 3-20)")
label_page.grid(row=1, column=0, padx=10, pady=5, sticky="E")

entry_page = tk.Entry(root, width=30)  # Allow input of page range in the format 'start_page-end_page'
entry_page.grid(row=1, column=1, padx=10, pady=5)

label_csv = tk.Label(root, text="Excel File: \n (default is pdf name with .xlsx extension)")
label_csv.grid(row=2, column=0, padx=10, pady=5, sticky="E")

entry_csv = tk.Entry(root, width=30)
entry_csv.grid(row=2, column=1, padx=10, pady=5)

button_submit = tk.Button(root, text="Submit", command=on_submit)
button_submit.grid(row=3, column=1, pady=10)

result_label = tk.Label(root, text="")
result_label.grid(row=4, column=0, columnspan=3, pady=10)

# Start the main loop
root.mainloop()
