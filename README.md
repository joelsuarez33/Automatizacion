# SAP Automation with Python

Automate repetitive SAP tasks using Python + SAP GUI Scripting.  
This project helps you save time, reduce manual work, and standardize processes.

---

## 👋 About me

My name is **Joel Suárez**, I'm a Data Analyst focused on process automation and business intelligence.  
Native Spanish speaker 🇦🇷 | Advanced English 🗣️

📫 [Connect with me on LinkedIn](https://www.linkedin.com/in/joel-f-suarez/)

---

## 🛠️ Tools I use

- Python 🐍
- Visual Studio Code 💻
- SQL 💾
- SAP GUI Scripting 📘
- Excel 📊
- Power BI 📈
- Git & GitHub 🔧

---

## 📂 Project structure

- `Script_SAP.py`: base script to connect and interact with SAP via scripting.
- You can insert your own logic based on SAP recordings directly adaptated to Python.

---

## ✅ What you can automate

✔️ Mass extraction from transactions (e.g., **FBL1N**, **MB51**)  
✔️ Automated form filling and navigation  
✔️ Data input from **Excel to SAP**  
✔️ Report customization and download  
✔️ Daily/monthly routine operations (e.g., invoice processing)

---

## 🚀 How to use it

1. **Enable scripting in SAP GUI**  
   - Go to: `SAP GUI Options > Accessibility & Scripting > Scripting`
   - Ensure scripting is enabled both client-side and server-side.

2. **Record your process in SAP**  
   - Use: `Customizing > Script Recording and Playback`
   - Perform the task manually and export the `.vbs` script.

3. **Convert `.vbs` to Python**  
   - Use Copilot or any AI assistant to adapt the logic to Python (based on the `Script_SAP.py` format)
   - Replace static values with dynamic variables as needed

4. **Run the script from VS Code**

---

## 💡 Example: Automated loop in MIRO

Check the file `example_miro_loop.py` for a script that automates invoice creation in a loop, using data from Excel.

---

## 📎 Requirements

- Python 3.8+
- SAP GUI for Windows (with scripting enabled)
- `pywin32` installed (`pip install pywin32`)

---

## 📥 Clone this repo

```bash
git clone https://github.com/joelsuarez33/automatizacion.git
