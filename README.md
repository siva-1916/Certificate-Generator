# **Certificate Generator** 🏆  

### **Automated Certificate Generation Using Python**  

This project automates the process of generating personalized certificates by dynamically filling in participant names from an Excel sheet into a Word document template and exporting them as PDFs.  

## **Features** 🚀  
✅ Reads participant names from an Excel file 📊  
✅ Replaces placeholders in a Word template (.docx) with participant names ✏️  
✅ Exports the final certificates as PDFs 📄  
✅ Batch processing for multiple participants at once ⏳  

## **Technologies Used** 🛠️  
- **Python** 🐍  
- **win32com.client (pywin32)** – For handling Microsoft Word automation  
- **openpyxl** – For reading Excel files  
- **os** – For file handling  

## **How It Works** 🔧  
1. Reads participant names from an Excel sheet (`data.xlsx`).  
2. Replaces the placeholder **"STUDENT"** in the Word template (`CERTIFICATE.docx`) with each participant's name.  
3. Saves the modified certificate as a PDF in the output directory.  
4. Resets the placeholder text in the template for the next iteration.  

## **Setup & Usage** 📌  
1. Install dependencies:  
   ```bash
   pip install pywin32 openpyxl
   ```  
2. Place your Excel file (`data.xlsx`) and Word template (`CERTIFICATE.docx`) in the project folder.  
3. Update the file paths in `generate_certificates()` accordingly.  
4. Run the script:  
   ```bash
   python certificate_generator.py
   ```  
5. Generated certificates will be saved in the **output** folder.  

## **Future Enhancements** 🌱  
🔹 Add support for dynamic text formatting  
🔹 Enable customization for multiple placeholders  
🔹 GUI-based interface for easy input  

### **Contributions & Feedback** 🤝  
Feel free to **fork, star ⭐, or contribute** to enhance this project! Any suggestions or improvements are welcome.  

### 📜 **License**  
MIT License
