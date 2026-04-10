
# ReportForge

![Python](https://img.shields.io/badge/Python-3.11-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey)

A local-first desktop application for pentest reporting, designed for security professionals who want full control over their data.

---

## 💡 Why this project matters

Most pentest reporting tools are either:

* Cloud-based (risk for sensitive data)
* Overcomplicated
* Expensive

**ReportForge solves this by being:**

* 🖥️ Fully local-first
* 🔐 Privacy-focused
* ⚡ Fast and simple to use

Built with real-world pentesting workflows in mind.

---

## 📸 Screenshot

![App Screenshot](assets/Dashboard.png)

---

## 🚀 Features

* 📁 Project-based pentest management
* 🐞 Structured findings with severity levels
* 📎 Evidence & attachment management
* 📊 Executive, technical and stakeholder summaries
* 📄 Automated report generation (DOCX/PDF)
* 🔗 CVE / CWE reference support
* 🧠 Methodology checklist integration

---

## 🛠️ Tech Stack

* **Python**
* **PySide6** (GUI)
* **SQLAlchemy** (ORM)
* **SQLite** (local database)

---

## ⚡ Quick Start (Windows 11)

### Prerequisites

* Python 3.11+
* Git

### Setup

```powershell
git clone https://github.com/serafinilorenzo87-png/ReportForge.git
cd ReportForge

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
python main.py
```

---

## ⚡ Quick Start (Linux)

### Prerequisites

* Python 3.11+
* Git

### Setup

```bash
git clone https://github.com/serafinilorenzo87-png/ReportForge.git
cd ReportForge

python3 -m venv .venv
source .venv/bin/activate

pip install -r requirements.txt
python3 main.py
```

---

## 🧠 How it works

* All data is stored **locally** (no cloud dependency)
* A SQLite database is created automatically on first run
* Evidence files are stored per project/finding
* Reports are generated from structured data

---

## 🔐 Security Philosophy

ReportForge is **local-first by design**:

* No external API calls
* No data exfiltration risk
* Ideal for handling sensitive pentest data

---

## 📂 Project Structure

```
app/                # Core application logic
assets/             # UI assets
tests/              # Test suite
examples/           # Sample reports
main.py             # Entry point
```

---

## 📦 Future Improvements

* Export templates customization
* Multi-user support
* Plugin system
* Optional cloud sync

---

## 🤝 Contributing

Pull requests are welcome.
For major changes, please open an issue first to discuss what you would like to change.

---

## 📜 License

MIT License

---

## 👤 Author

**Lorenzo Serafini**
Cybersecurity & Software Engineering
