# Excel Merger & Sorter Tool

A standalone utility for working with Excel reports generated from **XP Power’s DVT Automation** system.  
This tool simplifies the workflow by merging multiple individual reports into a consolidated workbook,  
and then sorting the worksheets according to the **Functional Test OMS Standard** sequence.

---

## ✨ Features
- **Merge Reports**  
  Combine multiple DVT Automation Excel outputs into a single workbook.
- **Custom Sort**  
  Worksheets are ordered based on the official OMS Functional Test sequence,  
  ensuring consistent alignment with validation standards.
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
4. Click **Sort** to reorder worksheets according to OMS sequence.

---

## 📋 Sorting Logic (As of September 11, 2025)
The sorter prioritizes worksheet names containing key substrings defined in the OMS sequence list, i.e.:

1. Ripple and Noise  
2. Start Up Time  
3. Start Up Rise Time  
4. Holdup Time  
5. Turn Off Fall Time  

Worksheets containing these substrings (case-insensitive) are grouped and ordered first.  
All other worksheets follow after, preserving their relative order.

---

## 🗂️ Project Structure
- **Excel Handler (VB.NET)** – Contains the worksheet sorting logic and merging logic.  
- **Excel Merger (C# / WinForms)** – A graphical interface built with MVP pattern, providing user interaction, task locking, and progress bar feedback.  

---

## ✅ Current Progress
- Successfully introduced **merge logic** (VB.NET) to combine individual reports into a single workbook.  
- Added **sort logic** (VB.NET) based on Functional Test OMS Standard sequence.  
  - ⚠️ Note: current implementation does not yet cover all DVT functional tests.  

---

## 🚀 Roadmap
- [ ] Extend sorting logic to include **all DVT functional tests**.  
- [ ] Add customizable sort sequences (via config file).  
- [ ] Introduce option to **create a new empty workbook** (based on a template file in the repository).  
  - New mode triggered either via a **checkbox** or a **separate button**.  
  - Reuse "Select Base" step to select the directory where the new workbook will be saved.  
  - Workbook will only be created **after the merge process completes**.  
  - Sorting will be blocked off until the file exists.  
- [ ] Drag-and-drop support for reports.  
- [ ] Logging of merge/sort operations.  
- [ ] Export summary report (optional).  

---

## 🛠️ Build Instructions
To build from source:

1. Clone the repository:  
   ```bash
   git clone https://github.com/WofflesAbdul/ExcelMergerSolution.git```
   Open the solution in Visual Studio 2019 or later.

2. Ensure that .NET Framework 4.8 is installed.
3. Add a reference to Microsoft.Office.Interop.Excel (if missing).
4. Right-click project → Add Reference → COM → Microsoft Excel XX.X Object Library
5. Build the solution (Ctrl+Shift+B).
6. Run the WinForms project to launch the tool.
