import os
import win32com.client as win32
import openpyxl

def fill_and_save_word_to_pdf(input_docx, output_pdf, participant_data):
    try:
        # Create a Word application
        word = win32.Dispatch('Word.Application')
        word.Visible = False  # Run in the background

        # Open the Word document
        doc = word.Documents.Open(input_docx)

        # Iterate through all the story ranges (main doc, headers, footers, etc.)
        for story_range in doc.StoryRanges:
            while story_range.Find.Execute('STUDENT'):
                story_range.Text = participant_data['name']

        # Save the modified Word document as a PDF
        temp_pdf = output_pdf + "_temp.pdf"
        doc.SaveAs(temp_pdf, FileFormat=17)
        
        # Iterate through all the story ranges (main doc, headers, footers, etc.)
        for story_range in doc.StoryRanges:
            while story_range.Find.Execute(participant_data['name']):
                story_range.text = 'STUDENT'


        # Close the Word document
        doc.Close()
        word.Quit()

        # Rename the PDF with the participant's name
        os.rename(temp_pdf, output_pdf)
        

        print(f"Certificate for {participant_data['name']} saved as {output_pdf}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# Function to generate certificates
def generate_certificates(input_excel_file, template_docx, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    workbook = openpyxl.load_workbook(input_excel_file)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, values_only=True):
        participant_data = {
            'name': row[0],
        }
        output_pdf = os.path.join(output_dir, f"{participant_data['name']}_certificate.pdf")
        fill_and_save_word_to_pdf(template_docx, output_pdf, participant_data)


# Usage
input_excel_file = r'C:\Users\SIVA PRASAD\OneDrive\mine\projects\CERTIFICATE GENERATOR\data.xlsx'
template_docx = r'C:\Users\SIVA PRASAD\OneDrive\mine\projects\CERTIFICATE GENERATOR\CERTIFICATE.docx'
output_directory = r'C:\Users\SIVA PRASAD\OneDrive\mine\projects\CERTIFICATE GENERATOR\output'

generate_certificates(input_excel_file, template_docx, output_directory)

