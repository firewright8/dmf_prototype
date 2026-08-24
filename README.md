A custom web application built for the District Administration, Raigarh (Government of Chhattisgarh) to automate project document formatting, prevent duplicate submissions, and enable public social audits via dynamically generated QR codes.

## Features

* **Strict Naming Convention:** Automatically renames uploaded documents (UCC, GEO, ADS, TES, PRP) to an immutable Base ID (`BLK_VIC_SEC_FY_SEQ`) to prevent broken links during status changes.
* **Duplicate Prevention:** Blocks submissions if the Work Order (WO) or Administrative Sanction (AS) number already exists in the database.
* **Dynamic Cascading Forms:** Village dropdowns automatically filter based on the selected Block Code using official District Code Registers.
* **Bulk State Sync:** Allows administrators to bulk-update project statuses (Submitted, Approved, Tendered, Ongoing, Completed, Dropped) by uploading a CSV/Excel file exported from the State DMF portal.
* **Social Audit QR Codes:** Automatically generates a printable site plaque containing a QR code. When scanned by citizens, it opens a mobile-friendly, read-only view of the official project documents.
* **In-Memory ZIP Packaging:** Bundles and downloads project files on the fly, dynamically appending the real-time project status to the filename without altering the server's physical storage.

## Project Structure

```text
dmf-raigarh-portal/
│
├── app.py                              # Core Flask application and routing logic
├── requirements.txt                    # Python dependencies required to run the app
├── DMF Raigarh - Code Register.xlsx    # Official master index of Block, Village, and Sector codes
│
├── static/                             # Publicly accessible files
│   └── qrcodes/                        # Auto-generated QR codes are saved here
│
├── templates/                          # HTML frontend templates
│   ├── form.html                       # Public submission form
│   ├── success.html                    # Post-submission confirmation screen
│   ├── plaque.html                     # Printable A4 social audit plaque
│   └── public_view.html                # Mobile-friendly read-only view for citizens
│
└── project_documents/                  # Physical uploads (PDFs, JPGs) are saved here (auto-generated)
```

## Setup & Installation (Windows)

These instructions are written for Windows users to set up the local development environment.

**Prerequisite:** Ensure [Python](https://www.python.org/downloads/) is installed on your computer. _Important: During installation, make sure the box that says "Add Python to PATH" is checked._

1. **Download the Project:** Download this repository to your computer and extract the folder.
    
2. **Open Command Prompt:** Press the Windows key, type `cmd`, and press Enter.
    
3. **Navigate to the Folder:** Use the `cd` command to move into the project folder. (Replace the path below with your actual folder path).
    
    DOS
    
    ```
    cd C:\Users\YourName\Downloads\dmf-raigarh-portal
    ```
    
4. **Create a Virtual Environment:** This isolates the project's packages from your main system.
    
    DOS
    
    ```
    python -m venv venv
    ```
    
5. **Activate the Virtual Environment:**
    
    DOS
    
    ```
    venv\Scripts\activate
    ```
    
    _(You should see `(venv)` appear at the start of your command prompt line)._
    
6. **Install Dependencies:**
    
    DOS
    
    ```
    pip install -r requirements.txt
    ```
    
7. **Run the Server:**
    
    DOS
    
    ```
    python app.py
    ```
    

## Using the Portal

Once the server is running, open any web browser (Chrome, Edge, etc.) and navigate to the following local addresses:

- **Submission Portal:** [http://127.0.0.1:5000/](https://www.google.com/search?q=http://127.0.0.1:5000/)
    
- **Admin Dashboard:** [http://127.0.0.1:5000/dashboard](https://www.google.com/search?q=http://127.0.0.1:5000/dashboard)
    

_To safely stop the server, go back to your Command Prompt and press `CTRL + C`._

## Important Deployment Notes

- **Database Generation:** The `dmf_trial.db` SQLite database is intentionally excluded from this repository. The Python script will automatically generate a fresh, blank database on your machine the first time you run it.
    
- **Security:** This prototype runs on a local development server. Before deploying to the live Raigarh NIC servers, the application must be migrated to a production WSGI server (like Gunicorn) and connected to a robust PostgreSQL database.
