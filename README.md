# ISM-Audit-Folder-Creation-Script
PowerShell Script for Automated Document Creation

This PowerShell script automates the creation of a directory structure and Word documents based on data from an Excel file. It utilizes the `PSWriteWord` module to simplify Word document generation.

## **Update**
Tested on latest edition of ISM March 2025 uploaded here: RFFR SoA based on ISM March 2025.xlsx
v1.1 - Set root folder to be 'Guideline' - Updated from root folder 'Section' in v1.0

## **Introduction**
For organizations handling Australian government information, understanding the relationship between the ACSC Information Security Manual (ISM) and the Protective Security Policy Framework (PSPF) is critical. The ISM is a dynamic resource that provides cybersecurity guidelines and practical advice for protecting information and systems. The PSPF establishes the Australian Government's protective security policy. The ISM is instrumental in enabling PSPF compliance, acting as a detailed implementation guide. While the PSPF specifies what security measures are required, the ISM outlines how to implement them effectively. By following the ISM's recommendations, agencies can demonstrate adherence to the PSPF, enhance their security posture, and ensure the confidentiality, integrity, and availability of government information.

![image](https://github.com/user-attachments/assets/c53fb7ce-47e9-4821-84c8-f6ff379828cd)
![image](https://github.com/user-attachments/assets/693cd417-ea48-427d-aa47-1cbedc9ce539)


## Prerequisites

* **PowerShell:** Windows PowerShell or PowerShell Core must be installed.
* **PSWriteWord Module:** Install the `PSWriteWord` module from the PowerShell Gallery:
    ```powershell
    Install-Module PSWriteWord
    ```
* **Microsoft Excel:** Microsoft Excel application must be installed (for COM object interaction).
* **Excel Data File:** You need an Excel file with data formatted as expected by the script.

## Script Functionality

The script performs the following actions:

1.  **Prompts for Input:**
    * Prompts the user to select an Excel file using a file selection dialog.
          ![image](https://github.com/user-attachments/assets/b57fd1de-6711-4ba7-951b-83b38c5689fb)

    * Prompts the user to enter the worksheet name from the selected Excel file.
          ![image](https://github.com/user-attachments/assets/fb25d891-3a14-4c6c-8bfd-6ecd22fa51e4)

    * Prompts the user to enter the base directory where folders will be created.
          ![image](https://github.com/user-attachments/assets/df44e4a2-4261-4d8f-829c-29fd026b1d2c)


2.  **Reads Excel Data:**
    * Opens the specified Excel file.
    * Accesses the specified worksheet.
    * Iterates through rows in the worksheet, extracting data from specific columns (B, D, and E).

3.  **Creates Directory Structure:**
    * For each row in the Excel sheet, it creates a directory path based on the extracted data from columns B and D. The directory path is constructed as `$baseDirectory\$columnB\ISM-$columnD`.
    * It creates the directory if it doesn't already exist.

4.  **Generates Word Documents:**
    * For each row, it creates a blank Word document in the created directory. The document filename is `ISM-$columnD.docx`.
    * It uses the `PSWriteWord` module to:
        * Create a new Word document.
        * Add text from column E as a heading (font size 21, bold).
        * Add empty line (font size 12).
        * Add text "Add Evidence *" (font size 12).
        * Save the Word document.

5.  **Cleans Up Excel Objects:**
    * Closes the Excel workbook.
    * Quits the Excel application.
    * Releases COM objects to free up resources.
    * Performs garbage collection.

6.  **Completion Message:**
    * Displays a "Script completed successfully." message.

## How to Use

1.  **Save the Script:** Save the PowerShell script as a `.ps1` file (e.g., `CreateDocsFromExcel.ps1`).

2.  **Prepare the Excel File:**
    * Ensure your Excel file has the data organized correctly.
    * The script expects data in columns B, D, and E.
    * Column B: Used for the first level subdirectory name.
    * Column D: Used for the second level subdirectory name and the document name.
    * Column E: Used for the heading text in the Word document.
    * The script starts processing from row 3.

3.  **Run the Script:**
    * Open PowerShell.
    * Navigate to the directory where you saved the script.
    * Execute the script: `.\CreateDocsFromExcel.ps1`

4.  **Follow the Prompts:**
    * The script will first display a file selection dialog. Choose your Excel file.
    * Next, it will prompt you for the worksheet name. Enter the correct sheet name.
    * Finally, it will prompt you for the base directory where the folders will be created.

5.  **Verify the Output:**
    * After the script finishes, check the specified base directory. You should see a folder structure created, with Word documents in the appropriate subfolders.

## Important Notes

* **Error Handling:** The script includes basic error handling for file existence. You might want to add more robust error handling for other potential issues (e.g., Excel errors, file write errors).
* **Customization:** You can customize the script to:
    * Change the columns used for data extraction.
    * Modify the directory structure or document naming conventions.
    * Add more content or formatting to the generated Word documents.
* **Security:** Be cautious when running scripts from untrusted sources. Review the code to ensure it's safe.
* **Excel COM Objects:** The script uses the Excel COM object, which can sometimes be resource-intensive.  The script includes COM object cleanup, but be mindful of potential performance implications with very large Excel files.

## Example Excel Data
![image](https://github.com/user-attachments/assets/f7b72438-e549-46d3-96bb-40518aac14b8)


## Output Directory Structure
1. High Level Folder Structure
![image](https://github.com/user-attachments/assets/c9443772-9be3-427e-a29b-08445ef5ecf3)

2. Inside each Folder structure/section:
![image](https://github.com/user-attachments/assets/8046be6e-fabd-4cc0-8c25-da7f467493fb)

3. Inside the Section folder structure:
![image](https://github.com/user-attachments/assets/88bca114-5a55-4b71-bdb8-ba8c5cf08f98)

3. Inside each ISM Docx file:
![image](https://github.com/user-attachments/assets/8e9b8c67-92ff-4a76-83b0-00ff783814d9)

## **Contributing** (Sharing is caring)

If you'd like to contribute to Project Title, here are some guidelines:

1. Fork the repository.
2. Create a new branch for your changes.
3. Make your changes.
4. Write tests to cover your changes.
5. Run the tests to ensure they pass.
6. Commit your changes.
7. Push your changes to your forked repository.
8. Submit a pull request.

## **License**

Project Title is released under the MIT License. See the **[LICENSE](https://www.blackbox.ai/share/LICENSE)** file for details.

## **Authors and Acknowledgment**

Project Title was created by **AUSK1LL9(https://github.com/UASK1LL9)**.

## **Code of Conduct**

Please note that this project is released with a Contributor Code of Conduct. By participating in this project, you agree to abide by its terms. See the **[CODE_OF_CONDUCT.md](https://www.blackbox.ai/share/CODE_OF_CONDUCT.md)** file for more information.

## **FAQ**

**Q:** What ISM does this work with?

**A:** Currently working on ISM December 2024. Also works on ISM March 2024. In theory this will work on any ISM format released by RFFR here: 
https://www.dewr.gov.au/right-fit-risk-cyber-security-accreditation/resources/rffr-statement-applicability-soa-template

**Q:** How do I use this script?

**A:** Follow the usage steps in the README file.

**Q:** How do I contribute to Project Title?

**A:** Follow the contributing guidelines in the README file.

**Q:** What license is Project Title released under?

**A:** Project Title is released under the MIT License. See the **[LICENSE](https://www.blackbox.ai/share/LICENSE)** file for details.

## **Changelog**

- **1.0** Initial release
- **1.1** Updated script with betetr error handling.
          Set root folder to be 'Guideline' - Updated from root folder 'Section' in v1.0

## **Contact**

If you have any questions or comments about Project Title, please contact **[AUSK1LL9](AUK1LL9@proton.me)**.
