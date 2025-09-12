# Excel Merger & Sorter Tool

A standalone utility for working with Excel reports generated from **XP Power’s DVT Automation** system.  
This tool simplifies the workflow by merging multiple individual reports into a consolidated workbook,  
and then sorting the worksheets according to the **Functional Test OMS Standard** sequence, including Vpsu and temperature ordering.

---

## ✨ Features
- **Merge Reports**  
  Combine multiple DVT Automation Excel outputs into a single workbook.
- **Custom Sort**  
  Worksheets are ordered based on:
  1. DVT functional test name (all functional tests are now supported).  
  2. Vpsu suffix values (higher percentages prioritized).  
  3. Temperature suffixes (increments of 5°C: ascending from 25°C to 70°C, then descending from 25°C to -40°C).  
  Ensures consistent alignment with validation standards.
- **Standalone Utility**  
  Runs independently, no need to manually open or modify Excel files.
- **Progress Feedback**  
  Includes UI feedback (progress bar) and task-safe locking of controls during long operations.

---

## 📦 Requirements & Dependencies
- **Windows OS** (tested on Windows 10/11)  
- **Microsoft Excel** (required for Excel Interop to function)  
- **.NET Framework 4.8** runtime or later  
- **Office Interop Assemblies**  
  - Usually installed with Microsoft Office.  
  - Alternatively, install via [Microsoft Office Primary Interop Assemblies (PIA)](https://learn.microsoft.com/en-us/visualstudio/vsto/installing-office-primary-interop-assemblies).  

⚠️ **Note:** This tool relies on Excel Interop, so Excel must be installed on the machine.  

---

## 🔧 Usage
1. Launch the tool (`.exe` build available in Releases).  
2. Select the base Excel file or folder of reports to merge.  
3. Click **Merge** to combine files into a consolidated workbook.  
4. Click **Sort** to reorder worksheets according to the OMS sequence, Vpsu, and temperature rules.

---

## 📋 Sorting Logic (As of September 12, 2025)

The sorter orders worksheets with a **three-level prioritization**:

1. **Functional Test Name**  
   Worksheets are grouped according to all DVT functional test names (case-insensitive).  
   Example tests include:  
   - Ripple and Noise  
   - Start Up Time  
   - Start Up Rise Time  
   - Hold-Up Time  
   - Turn Off Fall Time  
   
   *(Full list now fully implemented in [FunctionalTestSequence.vb](Excel%20Handling/Module/FunctionalTestSequence.vb).)*

2. **Vpsu Value**  
   Worksheets with `Vpsu = XX%` suffix are prioritized with higher percentages first.  
   Example order:  
   - Ripple and Noise Vpsu = 110%  
   - Ripple and Noise Vpsu = 80%  

3. **Temperature**  
   Worksheets with temperature suffixes (e.g., `25C`, `50C`, `-20C`) are ordered in increments of 5°C:  
   - Ascending from **25°C up to 70°C**  
   - Then descending from **25°C down to -40°C**  

Worksheets not matching any sequence are placed **after the prioritized sheets**, preserving their relative order.

---

## 🗂️ Project Structure
- **Excel Handler (VB.NET)** – Contains the worksheet sorting logic and merging logic.  
- **Excel Merger (C# / WinForms)** – A graphical interface built with MVP pattern, providing user interaction, task locking, and progress bar feedback.  

---

## ✅ Progress & Roadmap

- [x] Merge logic (VB.NET) to combine individual reports into a single workbook.  
- [x] Sort logic (VB.NET) based on Functional Test OMS Standard sequence, **Vpsu**, and **temperature ordering**.  
  - [x] All DVT functional tests are now supported.  
  - [x] Worksheets are grouped by functional test first, then ordered by descending Vpsu values, then temperature increments.  
- [ ] Add customizable sort sequences (via config file).  
- [ ] Introduce option to **create a new empty workbook** (based on a template file in the repository).  
  - [ ] Add a blank template file into the repository.  
  - [ ] Triggered via a **checkbox** or a **button**.  
  - [ ] Reuse "Select Base" step to select the directory where the new workbook will be saved.  
  - [ ] Workbook is created **after merge completes**; sorting is blocked until then.  
  - [ ] Introduce user selection for adding template worksheets, in the context of DVT Full Report (e.g., Title and Summary).  
    - ⚠️ Requires an additional template file containing the predefined Title and Summary sheets.
- [ ] Drag-and-drop support for reports.  
- [ ] Logging of merge/sort operations.  

---

## 🛠️ Build Instructions
To build from source:

1. Clone the repository:  
   ```bash
   git clone https://github.com/WofflesAbdul/ExcelMergerSolution.git
   ```

2. Open the solution in Visual Studio 2019 or later.
3. Ensure that .NET Framework 4.8 is installed.
4. Add a reference to Microsoft.Office.Interop.Excel (if missing).  
   - Right-click project → Add Reference → COM → Microsoft Excel XX.X Object Library
5. Build the solution (Ctrl+Shift+B).
6. Run the WinForms project to launch the tool.
