# Install Required Libraries
# !pip install pandas PyPDF2 reportlab pdfrw

import pandas as pd
from pdfrw import PdfReader as PdfRd, PdfWriter as PdfWt, PageMerge
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

# Load the Excel Spreadsheet
excel_path = 'C:/2024_Spring/greenSheetsProject/sourceHours03.xlsx'
df = pd.read_excel(excel_path)

# Display the first few rows of the dataframe
print(df.head())

# Parse Data Based on Month and Employee
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')  # Convert Date column to datetime
df.dropna(subset=['Date'], inplace=True)  # Drop rows with invalid dates
df['Month'] = df['Date'].dt.month
df['Year'] = df['Date'].dt.year.astype(int)  # Convert year to integer

# Group Data by Employee and Month
grouped = df.groupby(['Name', 'Month', 'Year'])

# Month mapping
month_mapping = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

# Set Up PDF Template Handling
template_path = 'C:/2024_Spring/greenSheetsProject/GreenCertificatedTimesheet_v04.pdf'

# Function to Fill PDF Template with Basic Elements
def fill_pdf(data, details, template_path, output_path):
    reader = PdfRd(template_path)
    writer = PdfWt()
    tmp_output = 'C:/2024_Spring/greenSheetsProject/tmp_output.pdf'

    total_hours = data['Hours'].sum()  # Calculate total hours

    for page in reader.pages:
        # Create a new overlay page with ReportLab
        c = canvas.Canvas(tmp_output, pagesize=letter)

        # Draw the text in the specified positions
        c.setFont("Helvetica-Bold", 12)  # Set font to Helvetica-Bold with size 12
        c.setFillColorRGB(0, 0, 0)  # Set text color to black

        # Header information
        c.drawString(130, 723, details['Last Name'])     # Last Name
        c.drawString(240, 723, details['First Name'])    # First Name
        c.drawString(302, 723, details['Initial'])       # Initial
        c.drawString(412, 723, details['Month'])         # Month
        c.drawString(512, 723, str(details['Year']))     # Year

        # Set font for description and hours
        c.setFont("Helvetica", 8)  # Set font to Helvetica with size 8
        c.setFillColorRGB(0, 0, 0)  # Set text color to black

        # Fill in the description and hours based on the day
        for row in data.itertuples():
            day_offset = (row.Date.day - 1) * 15  # Adjust spacing between lines
            desc_y_position = 638 - day_offset
            hours_y_position = 638 - day_offset
            # changed 126 to 116 5/17/24 dH
            c.drawString(116, desc_y_position, row.Description)
            c.drawString(504, hours_y_position, f"{row.Hours}")

        # Render total hours at the bottom of the Hours column
        c.setFont("Helvetica-Bold", 12)  # Increase font size by 20% (from 10 to 12)

        # c.drawString(504, 623 - (31 * 15) + 5, f"{total_hours}")  # Adjust the position for total hours
        c.drawString(504, 165, f"{total_hours}")  # Adjust the position for total hours

        c.save()

        overlay = PdfRd(tmp_output)
        merge_page = PageMerge(page)
        merge_page.add(overlay.pages[0]).render()

        writer.addpage(page)

    with open(output_path, 'wb') as f:
        writer.write(f)

# Save the Filled PDFs
output_dir = "C:/2024_Spring/greenSheetsProject/GeneratedTimesheets"
os.makedirs(output_dir, exist_ok=True)

# Generate timesheets for all grouped data
for (name, month, year), group in grouped:
    details = {
        'Last Name': name.split()[-1],
        'First Name': name.split()[0],
        'Initial': name.split()[1] if len(name.split()) > 2 else '',
        'Month': month_mapping[month][:3],  # Use the first three characters of the month name
        'Year': year  # Ensure year is an integer
    }
    output_path = f"{output_dir}/{name}_{details['Month']}_{year}_Timesheet.pdf"
    fill_pdf(group, details, template_path, output_path)

print("Timesheets generated. Please check the output directory.")
