# ⚡ ATS Resume Builder & Scorer — AI Powered

> Upload your resume (PDF/DOCX/TXT) → get ATS score out of 100 → see exactly what to fix. Powered by Claude AI or local engine.

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)
![AI](https://img.shields.io/badge/AI-Claude%20Sonnet-purple?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)

---

## ✨ Features

| Feature | Detail |
|---|---|
| 🤖 AI Mode | Paste Anthropic API key → Claude deeply reads & understands your resume |
| 🔧 Local Mode | Works without API key using built-in scoring engine |
| 📄 File Upload | PDF, DOCX, TXT — all supported |
| 📊 Score/100 | Full ATS compatibility score with grade |
| 📈 7 Categories | Individual scores with progress bars |
| 💡 Smart Suggestions | Prioritized: 🔴 Critical → 🟡 Important → 🟢 Tips |
| ⚡ Quick Wins | Top 3 fastest ways to improve your score |
| ✅ Strengths | What your resume already does well |
| 📝 Extracted Text | See exactly what ATS reads |
| 🖥️ Fullscreen UI | Dark purple theme, 4-tab layout |

---

## 🛠️ Installation

```bash
# Clone
git clone https://github.com/aman-codes777/ats-resume-builder.git
cd ats-resume-builder

# Install (PDF + DOCX support)
pip install pdfplumber python-docx
```

> `tkinter` comes built-in with Python — no install needed.

---

## ▶️ Run

```bash
python ats_builder.py
```

---

## 🤖 AI Mode Setup (Optional but Recommended)

1. Go to https://console.anthropic.com → create account → get API key
2. Open the app → paste API key in the top field → click **Apply**
3. Now Claude AI reads your resume and gives deep, personalized analysis

**Without API key:** Local engine still gives solid scoring across 7 categories.

---

## 🖥️ How to Use

1. Run `python ats_builder.py`
2. (Optional) Paste your Anthropic API key → Apply
3. Click **Browse File** → select your resume
4. Click **⚡ SCAN & SCORE MY RESUME**
5. App auto-switches to **Score Breakdown** tab
6. Check **Suggestions** tab for prioritized fixes
7. Check **Extracted Text** to see what ATS reads

---

## 📊 Scoring (100 pts)

| Category | Pts | What's Checked |
|---|---|---|
| 📱 Contact Info | 15 | Email, phone, LinkedIn, GitHub |
| 📂 Sections | 20 | Experience, Education, Skills, Projects, Certifications |
| 🎯 Action Verbs | 15 | Built, Led, Improved, Automated... |
| 💻 Tech Keywords | 20 | Python, SQL, Git, Docker, AWS... |
| 📊 Quantification | 15 | Numbers, %, metrics in bullets |
| 📐 Formatting | 10 | Word count, structure |
| 🤝 Soft Skills | 5 | Communication, teamwork... |

---

## 🔧 PDF Not Working?

```bash
pip install pdfplumber
```
If still failing → open your PDF → Select All → Copy → paste into Notepad → Save as `resume.txt` → upload the `.txt` file.

---

## 📁 Structure

```
ats-resume-builder/
├── ats_builder.py      ← Run this
├── requirements.txt
└── README.md
```

---

## 🧑‍💻 Author

**Mo Aman** | [GitHub](https://github.com/aman-codes777) | [LinkedIn](https://linkedin.com/in/amandeveloper777)

---

**GitHub Repo Name:** `ats-resume-builder`

**GitHub Description:**
```
⚡ AI-powered ATS Resume Scorer built in Python. Upload PDF/DOCX/TXT → get score/100 → 7-category breakdown → prioritized suggestions. Uses Claude AI (API key optional) or local engine. Dark GUI with Tkinter. by Mo Aman
```
