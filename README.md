# Digital Enablement Team - Daily Standup Dashboard (Dashboard Looker)

A lightweight, highly responsive, and data-driven dashboard application designed for the Digital Enablement Team. This tool visualizes team standups, bandwidth, leave statuses, and project blockers by pulling data directly from a centralized Excel file, Google Sheets, or SharePoint Excel link.

---

## 🚀 Features

### **Data Integration & Live Updates**
- **Flexible Data Sources:** By default, it reads from a local Excel file (`_Bandwidth Tracker.xlsx`). 
- **Cloud Connect:** Users can dynamically connect a **Google Sheet** (via a published export link) or a **SharePoint/OneDrive Excel** link directly from the UI using the "Data Source" modal.
- **Auto-Refresh:** The dashboard automatically polls the data source every 5 minutes to ensure the team is always looking at the latest standup data.

### **Standup & Delivery Tracking**
- **Smart "Yesterday & Today" Panels:** Automatically extracts and categorizes tasks into "What was done yesterday?" and "What will be done today?".
- **Deliverables Parsing:** A smart filter omits non-deliverable operational tasks (like daily syncs, calls, or meetings) from the main standup panels, focusing only on real work items. 
- **Task Merging:** Groups multiple tasks for the same person on the same project into a cleanly comma-separated single-line summary for readability.

### **Bandwidth & Team Availability**
- **Visual Bandwidth Overview:** Displays clear percentage-based visual bars representing the total capacity of each team member. 
- **Leave Status & Availability:** Automatically categorized tags indicate if a member is "Available", "Partial", on a "Half Day", or "On Leave" (Full Day). 
- **Team Detail Table:** Sortable and searchable table outlining team roles, leave status, and exact bandwidth.

### **Filtering & Search**
- **Global Search:** Find any team member, project name, or specific task instantly using the search bar.
- **Project Selection:** Filter the entire dashboard by selecting or deselecting specific ongoing projects from the right-hand sidebar.
- **Date Range Picker:** Drill down data to specific dates or sprints using the calendar filters.

### **Project Insights**
- **Left Sidebar Details:** Instantly displays specific Project Details, the Project Manager, MBR/QBR Dates, and critical Blockers/Dependencies based on the selected project context.

### **Presentation Mode**
- Enables an immersive fullscreen, clean UI ideal for sharing on a TV monitor or during a Zoom/Teams screen-share session.

---

## 🏗️ Architecture & Stack

### **Backend (`server.js`)**
- Node.js & Express server.
- Uses `xlsx` package to parse incoming local Excel files or downloaded Google Sheet buffers.
- **Endpoints:**
  - `/api/bandwidth`: Maps and cleans bandwidth, task, and leave status data.
  - `/api/qbr`: Extracts MBR/QBR dates and blockers.
  - `/api/dropdown`: Exposes lookup data (lookup tables for Managers and Project descriptions).
  - `/api/connect-source` & `/api/source-status`: Manages live cloud data source connections.

### **Frontend (`public/`)**
- Vanilla JavaScript (`app.js`), HTML5 (`index.html`), Vanilla CSS (`styles.css`).
- Clean, dependency-free frontend to ensure maximum performance and minimal footprint.
- Uses modern CSS Flexbox/Grid and variables for theming.

### **Mock Data Generator (`fill_recent_data.py`)**
- A handy Python script utilizing `openpyxl`.
- Interacts with `_Bandwidth Tracker.xlsx` to automatically generate bulk randomized log entries for "Yesterday" and "Today". 
- Extremely useful for testing UI changes or populating the board for a demonstration.

---

## ⚙️ Setup & Installation

### 1. Prerequisites
- **Node.js** (v14 or higher is recommended)
- **Python** (version 3, if you plan to use the mock data script)

### 2. Install Dependencies
Navigate to the root directory of the project and install the required Node modules (`express`, `cors`, `xlsx`):

```bash
npm install
```

*(Optional)* If you wish to use the Python test-data generator, install its dependency:
```bash
pip install openpyxl
```

### 3. Data File Placement
Ensure you have the master bandwidth tracker Excel file located at the required path. By default, the server expects it at:
`../pov/_Bandwidth Tracker.xlsx` 
*(Note: You can override this using the `EXCEL_PATH` environment variable).*

### 4. Start the Application
You can run the application using npm:

```bash
npm start
```
*(Alternatively: `npm run dev` or `node server.js`)*

The server will initialize on port `3000`. 
**Open your browser to: `http://localhost:3000`**

---

## 🛠️ How to Use the Mock Data Generator

If your Excel file is empty and you want to visualize the dashboard immediately, use the included Python script to fill it with reliable test data.

1. Ensure `EXCEL_PATH` inside `fill_recent_data.py` points to your active Excel file.
2. Run the script:
   ```bash
   python fill_recent_data.py
   ```
3. The script will automatically read the `Drop Down` lookup sheet, grab team members and projects, and append about 20 random tasks (both deliverables and ops) with realistic time tracking to the currently active dates (Yesterday and Today).
4. Refresh the web dashboard (or wait for the auto-refresh) to see the new data populated!
