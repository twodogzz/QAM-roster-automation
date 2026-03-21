# ✅ **VSCode Codex Requirement Prompt — QAM Roster Automation Refactor**

**Goal:**  
Refactor and enhance the existing **QAM‑Roster‑automation** Python application so it correctly processes roster events that may span **previous, current, and next months**, instead of only the month indicated in the DOCX file’s title.
---
## **🔧 Functional Requirements**
### **1. Multi‑Month Event Handling**
- The DOCX roster file now includes events from:
  - the previous month  
  - the current month  
  - the following month  
  - Review the docx file provided to understand the new format.
- Update the parsing logic so **all valid events are captured**, regardless of month.
- Ensure the script **does not rely solely on the document title** to determine the month.
- Implement a reliable method to infer the correct date for each event, using:
  - event day numbers  
  - context from the roster layout  
  - transitions between months in the table  

### **2. Maintain Simplicity**
- Keep the code readable and beginner‑friendly.
- Avoid unnecessary abstractions or complex patterns.
- Add comments explaining key logic, especially around date inference.
- Review and update the project.json information with the appropriate version increment. (new ver is "0.1.2" with today's date)

### **3. Backward Compatibility**
- Existing behaviour for single‑month rosters should still work.
- No breaking changes to the current CLI or file‑input workflow.
---
## **🧹 Code Quality Requirements**
- Improve structure and readability.
- Add small helper functions where appropriate.
- Ensure variable names are clear and consistent.
- Add basic error handling for malformed or unexpected DOCX formats.
- Add lightweight logging or print statements to help the user understand what the script is doing.
---
## **📘 Documentation Updates**
Review the existing user manual and update it to reflect:

### **New behaviour**
- The script now processes events across three months.
- How the script determines the correct month for each event.
- Any new assumptions or limitations.

### **Examples**
-  Review the docx file provided to understand the new format. 
- Expected output includes calendar events for months other than the titles e.g. roster with "Queensland Air Museum - Front Counter Roster March 2026 v9" may, but not necessarily include Februrary and April 2026 too.

### **Troubleshooting**
- Add notes for common issues (e.g., ambiguous dates, malformed tables).
---
## **❓ Clarification Needed**
Before proceeding, identify any ambiguous or incomplete instructions in the existing code or documentation and ask for clarification.