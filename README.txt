# Tool Inventory Pro v3

Features: tool photos; maintenance and warranty dates; batteries/accessories; QR code PDF labels; dashboard charts; CSV/Excel reports; detailed storage area, cabinet, shelf, and bin.

## Windows setup
1. Install Python 3.10+ from https://www.python.org/downloads/ (enable Add Python to PATH).
2. Extract this ZIP to a permanent folder, e.g. C:\ToolInventory.
3. In File Explorer, open the folder, click address bar, type cmd, press Enter.
4. Run:
   py -m venv .venv
   .venv\Scripts\activate
   py -m pip install -r requirements.txt
   py -m streamlit run app.py
5. Open the local URL printed in the terminal, usually http://localhost:8501.

## Data
tools.xlsx contains Tools, Maintenance, and Accessories sheets. Photos are stored in photos\. QR images are in qr_labels\ and the printable PDF is generated as tool_qr_labels.pdf. Back up tools.xlsx and both folders. Keep the workbook closed in Excel while the app saves.

QR codes encode the tool ID only; they do not open the app automatically or provide remote access. Intended for local single-user use. Do not expose publicly without authentication and appropriate security.
