# ✅ **Follow‑Up Codex Requirement Prompt — “All Workers” Calendar Automation**

Extend the previous refactor requirements with the following enhancement:

## **🔧 New Functional Requirement — “All Workers” Calendar Events**
A portion of the existing script qam_docx_to_ics.py, already creates calendar events for **all workers**, not just for **Wayne**.  
Enhance the main Google Calendar automation flow to support this capability in a clean and optional way.

### **Requirements**
- Add a new function that generates Google Calendar events for **all workers**.
- This function must:
  - Use a **separate Google Calendar** from the one used for Wayne’s events.
  - Follow the same improved multi‑month event‑parsing logic defined earlier.
  - Remain simple and easy to understand.

### **User‑Controlled Execution**
- The “All workers” calendar update must **not** run automatically.
- Require an **explicit user instruction** (e.g., CLI flag, menu option, or prompt) before executing the “All workers” update.
- Ensure the default behaviour continues to update only Wayne’s calendar.

### **Code Quality**
- Keep the implementation lightweight and beginner‑friendly.
- Add comments explaining:
  - how the “All workers” flow differs from the Wayne‑only flow  
  - how the script selects the correct calendar  
  - how the user triggers the optional function  

---

## **📘 Documentation Updates**
Update the user manual to include:

- A description of the new “All workers” calendar update feature.
- Clear instructions on how the user explicitly triggers it.
- Notes about the separate Google Calendar used for these events.
- Any limitations or assumptions.

---

## **❓ Clarification Needed**
Identify any ambiguous or incomplete instructions related to:
- how workers are identified  
- how the “All workers” calendar should be structured  
- how the user should trigger the optional function  

Ask for clarification before implementing.