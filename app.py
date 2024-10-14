import re

import streamlit as st
import pandas as pd
from io import BytesIO
import PyPDF2

st.title("PDF to Excel Converter")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
series_name = st.text_input("Series code name (e.g., KS-C5-):", "KS-C5-")
output_filename = st.text_input("Enter output filename (e.g., output.xlsx):", "output.xlsx")

if uploaded_file is not None:
    # Read the PDF file
    pdf_reader = PyPDF2.PdfReader(uploaded_file)

    # Extract the text from each page
    all_text = ''
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        all_text += page.extract_text()

    # Split the text into sections based on the "Lot No:" header
    lot_sections = all_text.split("Lot No:")

    # Initialize an empty list to store data for each lot
    all_data = []

    # Initialize a counter for unique codes
    unique_code_counter = 1
    seen_codes = set()

    for section in lot_sections[1:]:
        # Extract lot number from the first line
        try:
            lot_no = int(section.split()[0])
        except (ValueError, IndexError) as e:
            print(f"Error extracting lot number from section: {section[:50]}... - {e}")
            continue

        print(f"Processing Lot No: {lot_no}")

        # Initialize variables to track data within each lot
        lot_rough_wt = 0
        lot_polish_wt = 0
        lot_exp_carat = []
        shape = []
        # print(f"- section = {section}")
        # Split the section into parts based on the pattern
        pattern = r"(out of bound|\+[ ]?.*?)\s+(.+?)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(.+?)\s+(\d+\.\d+)\s+(\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2} [AP]M)"
        parts = re.findall(pattern, section)
        # print(f"- parts = {parts}")
        for part in parts:
            print(f"- parts = {part}")
            print(f" -part 1 = {part[0]}")
            try:
                rough_wt = float(part[2])
                exp_carat = float(part[4])
                rough_to_polish_pct = float(part[6]) / 100
                lot_rough_wt += rough_wt
                shape.append(part[5])
                lot_polish_wt += exp_carat
                lot_exp_carat.append(exp_carat)
                # Don't append to all_data here, it will be done later
                print(f"  - Data extracted: {rough_wt}, {exp_carat}, {rough_to_polish_pct}, {shape}")

            except (IndexError, ValueError) as e:
                print(f"  - Error processing part: {part} - {e}")

        # Sort exp_carat values in descending order for the lot
        lot_exp_carat.sort(reverse=True)
        if lot_rough_wt != 0:
            rough_to_polish_pct = float((lot_polish_wt * 100) / lot_rough_wt)
        else:
            rough_to_polish_pct = 0
        # Assign unique SR. NO. for each unique code
        code = series_name + str(lot_no)
        print(f"  - code no= {code}")
        if code not in seen_codes:
            sr_no = unique_code_counter
            unique_code_counter += 1

        # Add the data to the list with the correct sr_no and code (or blank if code is repeated)
        all_data.append(["", "", "", "", "", ""])
        for i, (shape_value, exp_carat_value) in enumerate(zip(shape, lot_exp_carat)):
            print(f"  - code no2= {code}")
            if code not in seen_codes:  # Check if code has been seen before
                seen_codes.add(code)
                all_data.append([sr_no, code, round(lot_rough_wt, 2), shape_value, round(exp_carat_value, 2), round(rough_to_polish_pct, 2)])
            else:
                all_data.append(["", "", "", shape_value, round(exp_carat_value, 2), ""])


    # Create a DataFrame from the extracted data
    df = pd.DataFrame(all_data, columns=["SR. NO.", "CODE", "ROUGH WT.", "SHAPE", "EST. POL WT.", "% ROUGH TO POLISH"])
    print(df.head)

    # Convert the DataFrame to Excel format
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Call save() on the ExcelWriter object (writer)
    writer.close()

    st.download_button(
        label="Download Excel file",
        data=output.getvalue(),
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )