# lotus_leap
python code to load a lotus notes nsf file and recursively extract all docs and atttachments





📂 Lotus Notes NSF Data Extractor
This Python script extracts documents and attachments from a Lotus Notes .nsf database and saves them into structured folders on your filesystem.

🚀 Features
✅ Connects to Lotus Notes NSF databases using COM automation.
✅ Extracts all documents and their fields into text files.
✅ Retrieves and saves embedded attachments.
✅ Organizes data into folders named after each document's UniversalID.
📁 Repository
🔗 GitHub: rossn-hc/lotus_leap

🛠️ Requirements
💻 Windows OS (for win32com.client)
📧 Lotus Notes Client installed (to access .nsf files)
🐍 Python 3.7+
📦 Python Dependencies
pywin32 — COM interface to interact with Lotus Notes.
⚙️ Setup Instructions
1️⃣ Clone the Repository
bash
Copy
Edit
git clone https://github.com/rossn-hc/lotus_leap.git
cd lotus_leap
2️⃣ Create and Activate a Virtual Environment
Windows CMD:
bash
Copy
Edit
python -m venv venv
venv\Scripts\activate
PowerShell:
bash
Copy
Edit
python -m venv venv
.\venv\Scripts\Activate.ps1
3️⃣ Install Dependencies
bash
Copy
Edit
pip install pywin32
4️⃣ Place the .nsf File
Ensure the .nsf file you want to query is located in the Lotus Notes Data folder, typically found at:

java
Copy
Edit
C:\Program Files (x86)\IBM\Lotus\Notes\Data\
⚠️ Note: Adjust the path if Lotus Notes is installed in a different directory.

📖 Usage
▶️ Run the Script
bash
Copy
Edit
python extract_nsf_data.py
⚙️ Script Parameters
Parameter	Description	Required	Default
password	Your Lotus Notes client password.	✅	N/A
nsf_path	Path to the .nsf file relative to the Lotus Notes Data dir.	✅	N/A
output_dir	Directory to save extracted data.	❌	output
💡 Example
python
Copy
Edit
if __name__ == '__main__':
    # Replace with your Lotus Notes password and NSF file name
    extract_nsf_data_to_folders("your_password", "FND-CHHAD-Reference-Library.nsf")
Run the script:

bash
Copy
Edit
python extract_nsf_data.py
📂 Output
The extracted data will be saved in an output directory (default: ./output). Each document from the .nsf file will have its own folder, named using the document's UniversalID.

🗂️ Output Folder Structure:
bash
Copy
Edit
output/
├── <UniversalID>/
│   ├── document.txt          # Document fields and values
│   └── attachment1.ext       # Extracted attachment (if any)
├── <UniversalID>/
│   ├── document.txt
│   └── attachment2.ext
...
document.txt — Contains all fields and values from the document.
Attachments — All embedded files extracted from the document.
💡 Custom Output Directory:
To specify a different output directory, use the output_dir parameter in the function call.

🛑 Troubleshooting
⚠️ Lotus Notes COM Errors:

Ensure Lotus Notes is installed and properly configured.
Run the script with the same user profile as Lotus Notes.
⚠️ Incorrect .nsf Path:

Verify the .nsf file exists in the Lotus Notes Data folder.
Use relative paths when specifying the .nsf file.
⚠️ Permission Issues:

Run the script with elevated privileges if necessary.
📄 License
MIT License

This README is now optimized for clarity and presentation. Let me know if you'd like further changes!