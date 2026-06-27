"""
ATS Resume Builder & Scorer — AI Powered
Uses Claude AI API to deeply analyze resumes
Author : Mo Aman | github.com/aman-codes777
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import re
import json
import urllib.request
import urllib.error
import base64

# ── Optional PDF/DOCX libs ───────────────────────────────────────────────────
try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    import docx as python_docx
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ════════════════════════════════════════════════════════════════════════════
#  CONFIG — PUT YOUR API KEY HERE
# ════════════════════════════════════════════════════════════════════════════
ANTHROPIC_API_KEY = ""   # <-- Paste your Anthropic API key here

# ════════════════════════════════════════════════════════════════════════════
#  COLORS & FONTS
# ════════════════════════════════════════════════════════════════════════════
BG      = "#0d0d0d"
CARD    = "#161616"
CARD2   = "#1e1e1e"
ACCENT  = "#6c63ff"
ACCENT2 = "#5a52e0"
GREEN   = "#00e676"
YELLOW  = "#ffab00"
RED     = "#ff5252"
TEXT    = "#f0f0f0"
SUB     = "#999999"
BORDER  = "#2a2a2a"

FH  = ("Segoe UI", 18, "bold")
FL  = ("Segoe UI", 13, "bold")
FM  = ("Segoe UI", 10)
FB  = ("Segoe UI", 10, "bold")
FSM = ("Segoe UI", 9)
FXL = ("Segoe UI", 36, "bold")

# ════════════════════════════════════════════════════════════════════════════
#  TEXT EXTRACTION
# ════════════════════════════════════════════════════════════════════════════

def extract_text(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".pdf":
        if PDF_OK:
            text = ""
            try:
                with pdfplumber.open(filepath) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            text += t + "\n"
                if text.strip():
                    return text
            except Exception:
                pass
        # Fallback: read raw bytes and extract printable text
        with open(filepath, "rb") as f:
            raw = f.read()
        # Extract readable text from PDF bytes
        text = ""
        for chunk in raw.split(b"BT"):
            if b"ET" in chunk:
                inner = chunk.split(b"ET")[0]
                for match in re.finditer(rb'\(([^\)]{1,200})\)', inner):
                    try:
                        part = match.group(1).decode("latin-1", errors="ignore")
                        part = part.strip()
                        if len(part) > 1 and any(c.isalpha() for c in part):
                            text += part + " "
                    except Exception:
                        pass
        # Also try to extract text between stream markers
        for match in re.finditer(rb'stream(.*?)endstream', raw, re.DOTALL):
            chunk = match.group(1)
            try:
                decoded = chunk.decode("latin-1", errors="ignore")
                words = re.findall(r'[A-Za-z0-9@.+\-_]{2,}', decoded)
                text += " ".join(words) + " "
            except Exception:
                pass
        return text if text.strip() else "ERROR:PDF_EXTRACT_FAILED"

    elif ext == ".docx":
        if DOCX_OK:
            try:
                doc = python_docx.Document(filepath)
                return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            except Exception:
                pass
        return "ERROR:NO_DOCX"

    elif ext in (".txt", ".text"):
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    return "ERROR:UNSUPPORTED"


def get_file_as_base64(filepath: str) -> tuple:
    """Return (base64_data, media_type) for image-based PDF sending."""
    with open(filepath, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode("utf-8")
    return b64, "application/pdf"


# ════════════════════════════════════════════════════════════════════════════
#  CLAUDE AI SCANNER
# ════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are an expert ATS (Applicant Tracking System) resume analyzer with 10+ years of experience in HR and recruitment. 
You analyze resumes deeply and return ONLY valid JSON — no markdown, no explanation, just the raw JSON object."""

ATS_PROMPT = """Analyze this resume text as an ATS system would. Return ONLY a JSON object with this exact structure:

{
  "total_score": <integer 0-100>,
  "grade": "<Excellent|Good|Average|Needs Work>",
  "grade_emoji": "<🏆|👍|⚠️|🔴>",
  "grade_message": "<one sentence about overall quality>",
  "categories": {
    "contact_info": {
      "score": <0-15>,
      "max": 15,
      "has_email": <true/false>,
      "has_phone": <true/false>,
      "has_linkedin": <true/false>,
      "has_github": <true/false>,
      "detail": "<what was found>"
    },
    "sections": {
      "score": <0-20>,
      "max": 20,
      "found": ["<section names found>"],
      "missing": ["<important sections missing>"],
      "detail": "<summary>"
    },
    "action_verbs": {
      "score": <0-15>,
      "max": 15,
      "count": <integer>,
      "found": ["<up to 8 verbs found>"],
      "detail": "<assessment>"
    },
    "tech_keywords": {
      "score": <0-20>,
      "max": 20,
      "count": <integer>,
      "found": ["<tech skills/tools found>"],
      "detail": "<assessment>"
    },
    "quantification": {
      "score": <0-15>,
      "max": 15,
      "count": <integer>,
      "examples": ["<examples of numbers/metrics found>"],
      "detail": "<assessment>"
    },
    "formatting": {
      "score": <0-10>,
      "max": 10,
      "word_count": <integer>,
      "detail": "<formatting assessment>"
    },
    "soft_skills": {
      "score": <0-5>,
      "max": 5,
      "found": ["<soft skills found>"],
      "detail": "<assessment>"
    }
  },
  "suggestions": [
    {
      "priority": "<Critical|Important|Tip>",
      "title": "<short title>",
      "description": "<detailed actionable advice — be specific and helpful>"
    }
  ],
  "strengths": ["<3-5 things the resume does well>"],
  "quick_wins": ["<3 fastest things to improve score immediately>"]
}

Be thorough, accurate, and genuinely helpful. The suggestions must be specific to THIS resume's actual content.

RESUME TEXT:
"""


def call_claude_api(resume_text: str, api_key: str) -> dict:
    """Call Anthropic API and return parsed JSON analysis."""
    url     = "https://api.anthropic.com/v1/messages"
    payload = {
        "model": "claude-sonnet-4-6",
        "max_tokens": 2000,
        "system": SYSTEM_PROMPT,
        "messages": [
            {
                "role": "user",
                "content": ATS_PROMPT + resume_text[:8000]
            }
        ]
    }
    body    = json.dumps(payload).encode("utf-8")
    req     = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type":      "application/json",
            "x-api-key":         api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    raw_text = data["content"][0]["text"].strip()
    # Strip markdown code fences if present
    raw_text = re.sub(r'^```(?:json)?\s*', '', raw_text)
    raw_text = re.sub(r'\s*```$', '', raw_text)
    return json.loads(raw_text)


# ════════════════════════════════════════════════════════════════════════════
#  FALLBACK LOCAL SCORER (when no API key)
# ════════════════════════════════════════════════════════════════════════════

def local_score(text: str) -> dict:
    tl = text.lower()
    ACTION_VERBS = ["developed","built","created","designed","implemented",
                    "managed","led","improved","increased","decreased","reduced",
                    "achieved","delivered","launched","optimized","automated",
                    "analyzed","collaborated","coordinated","generated","resolved",
                    "maintained","deployed","integrated","engineered","streamlined"]
    TECH = ["python","javascript","java","c++","sql","html","css","react","node",
            "django","flask","tensorflow","pytorch","machine learning","pandas",
            "numpy","matplotlib","git","github","docker","aws","azure","linux",
            "api","rest","excel","power bi","tableau","mongodb","mysql","ai"]
    SOFT = ["communication","teamwork","leadership","problem solving",
            "critical thinking","time management","adaptability","analytical"]
    SECTIONS = ["experience","education","skills","projects","certifications",
                "summary","objective","achievements"]

    has_email    = bool(re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text))
    has_phone    = bool(re.search(r'(\+?\d[\d\s\-]{7,}\d)', text))
    has_linkedin = "linkedin" in tl
    has_github   = "github" in tl
    c_score = (4 if has_email else 0)+(3 if has_phone else 0)+(4 if has_linkedin else 0)+(4 if has_github else 0)

    found_sec = [s for s in SECTIONS if s in tl]
    miss_sec  = [s for s in ["experience","education","skills","projects","certifications","summary"] if s not in tl]
    s_score   = min(20, len(found_sec)*2)

    found_v   = [v for v in ACTION_VERBS if v in tl]
    v_score   = min(15, len(found_v)*2)

    found_t   = [k for k in TECH if k in tl]
    t_score   = min(20, len(found_t)*2)

    quant_pat = [r'\d+\s*%', r'\d+\s*x\b', r'\$[\d,]+',
                 r'\d+\s*(users|clients|team|projects|hours)']
    quant_m   = []
    for p in quant_pat:
        quant_m += re.findall(p, tl)
    q_score   = min(15, len(quant_m)*3)

    wc        = len(text.split())
    f_score   = min(10, (5 if 300<=wc<=800 else 2)+(3 if len(text.splitlines())>=15 else 1)+2)

    found_s   = [s for s in SOFT if s in tl]
    ss_score  = min(5, len(found_s))

    total = c_score+s_score+v_score+t_score+q_score+f_score+ss_score

    if total >= 80:  grade,emoji,msg = "Excellent","🏆","Great ATS compatibility!"
    elif total >= 65: grade,emoji,msg = "Good","👍","Solid resume with room to improve."
    elif total >= 50: grade,emoji,msg = "Average","⚠️","May be filtered by strict ATS systems."
    else:             grade,emoji,msg = "Needs Work","🔴","High risk of ATS rejection."

    sugg = []
    if not has_email:    sugg.append({"priority":"Critical","title":"Add Email","description":"Your email address is missing. ATS systems require it to contact you. Add it at the top of your resume."})
    if not has_phone:    sugg.append({"priority":"Critical","title":"Add Phone","description":"Add your phone number. Recruiters need a direct way to reach you."})
    if not has_linkedin: sugg.append({"priority":"Important","title":"Add LinkedIn","description":"Add your LinkedIn profile URL. Most ATS systems scan for it and recruiters check it."})
    if not has_github:   sugg.append({"priority":"Important","title":"Add GitHub","description":"Add your GitHub URL. Essential for tech/developer roles to show real work."})
    for sec in miss_sec: sugg.append({"priority":"Critical","title":f"Add {sec.title()} Section","description":f"Missing '{sec.title()}' section. ATS systems look for this heading explicitly. Add it with relevant content."})
    if v_score < 10:     sugg.append({"priority":"Important","title":"Use More Action Verbs","description":f"Only {len(found_v)} action verbs found. Use: Built, Developed, Improved, Led, Automated, Achieved, Deployed. Start every bullet point with an action verb."})
    if t_score < 10:     sugg.append({"priority":"Important","title":"Add Tech Keywords","description":f"Only {len(found_t)} tech keywords detected. Add specific tools and languages you know — match the exact words from job descriptions."})
    if q_score < 6:      sugg.append({"priority":"Critical","title":"Add Numbers & Metrics","description":"No quantified achievements found! Add numbers: 'Improved performance by 40%', 'Built app used by 500+ users', 'Reduced load time by 2s'. Numbers are the #1 ATS booster."})
    if wc < 300:         sugg.append({"priority":"Important","title":"Expand Resume Content","description":f"Resume is too short ({wc} words). Aim for 400-700 words. Add a professional summary, expand project descriptions, and list more skills."})
    sugg.append({"priority":"Tip","title":"Use Standard Headings","description":"Use exact headings: 'Work Experience', 'Education', 'Skills', 'Projects'. ATS systems match these keywords precisely."})
    sugg.append({"priority":"Tip","title":"Avoid Tables & Columns","description":"ATS systems cannot read tables, columns, text boxes, or images. Use a simple single-column plain text format."})
    sugg.append({"priority":"Tip","title":"Tailor for Each Job","description":"Copy exact keywords from each job description and include them in your resume. ATS does exact keyword matching."})

    return {
        "total_score": min(100,total),
        "grade": grade, "grade_emoji": emoji, "grade_message": msg,
        "categories": {
            "contact_info":   {"score":c_score,"max":15,"has_email":has_email,"has_phone":has_phone,"has_linkedin":has_linkedin,"has_github":has_github,"detail":f"Email:{'✅' if has_email else '❌'} Phone:{'✅' if has_phone else '❌'} LinkedIn:{'✅' if has_linkedin else '❌'} GitHub:{'✅' if has_github else '❌'}"},
            "sections":       {"score":s_score,"max":20,"found":found_sec,"missing":miss_sec,"detail":f"Found {len(found_sec)} sections"},
            "action_verbs":   {"score":v_score,"max":15,"count":len(found_v),"found":found_v[:8],"detail":f"{len(found_v)} action verbs detected"},
            "tech_keywords":  {"score":t_score,"max":20,"count":len(found_t),"found":found_t,"detail":f"{len(found_t)} tech keywords found"},
            "quantification": {"score":q_score,"max":15,"count":len(quant_m),"examples":quant_m[:5],"detail":f"{len(quant_m)} quantified achievements"},
            "formatting":     {"score":f_score,"max":10,"word_count":wc,"detail":f"{wc} words — {'ideal' if 300<=wc<=800 else 'too short' if wc<300 else 'too long'}"},
            "soft_skills":    {"score":ss_score,"max":5,"found":found_s,"detail":f"{len(found_s)} soft skills detected"},
        },
        "suggestions": sugg,
        "strengths":   [f"Has {len(found_t)} tech skills listed","Shows project work"] if found_t else ["Resume uploaded successfully"],
        "quick_wins":  ["Add LinkedIn & GitHub URLs","Add numbers to bullet points","Add a Professional Summary section"],
    }


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════

class ATSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ATS Resume Scorer — AI Powered  |  Mo Aman")
        self.configure(bg=BG)
        self.state("zoomed")
        self.resizable(True, True)

        self._filepath = None
        self._text     = ""
        self._result   = {}

        self._build_ui()

    def _build_ui(self):
        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        # ── HEADER ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT, height=60)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚡  ATS Resume Builder & Scorer",
                 bg=ACCENT, fg=TEXT, font=FH, pady=14).pack(side="left", padx=22)
        self.ai_badge = tk.Label(hdr,
            text="🤖 AI MODE" if ANTHROPIC_API_KEY else "🔧 LOCAL MODE",
            bg="#4a42cc" if ANTHROPIC_API_KEY else "#333",
            fg=TEXT, font=FSM, padx=10, pady=4)
        self.ai_badge.pack(side="right", padx=12)
        tk.Label(hdr, text="by Mo Aman",
                 bg=ACCENT, fg="#ccccff", font=FSM).pack(side="right", padx=4)

        # ── NOTEBOOK ────────────────────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",     background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=CARD2, foreground=SUB,
                        font=FB, padding=[16, 8])
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", TEXT)])
        style.configure("ATS.Horizontal.TProgressbar",
                        troughcolor=BORDER, background=ACCENT,
                        bordercolor=CARD, lightcolor=ACCENT, darkcolor=ACCENT,
                        thickness=20)

        self.nb = ttk.Notebook(self)
        self.nb.grid(row=1, column=0, sticky="nsew")

        self.t_upload  = tk.Frame(self.nb, bg=BG)
        self.t_score   = tk.Frame(self.nb, bg=BG)
        self.t_suggest = tk.Frame(self.nb, bg=BG)
        self.t_text    = tk.Frame(self.nb, bg=BG)

        self.nb.add(self.t_upload,  text="  📄  Upload & Scan  ")
        self.nb.add(self.t_score,   text="  📊  Score Breakdown  ")
        self.nb.add(self.t_suggest, text="  💡  Suggestions  ")
        self.nb.add(self.t_text,    text="  📝  Extracted Text  ")

        self._tab_upload()
        self._tab_score()
        self._tab_suggest()
        self._tab_text()

    # ── SCROLLABLE HELPER ────────────────────────────────────────────────────
    def _scrollable(self, parent):
        outer = tk.Frame(parent, bg=BG)
        outer.pack(fill="both", expand=True, padx=30, pady=(10, 20))
        c = tk.Canvas(outer, bg=BG, highlightthickness=0)
        c.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=c.yview)
        vsb.pack(side="right", fill="y")
        c.configure(yscrollcommand=vsb.set)
        inner = tk.Frame(c, bg=BG)
        win   = c.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: c.configure(scrollregion=c.bbox("all")))
        c.bind("<Configure>",     lambda e: c.itemconfig(win, width=e.width))
        c.bind_all("<MouseWheel>",
                   lambda e: c.yview_scroll(int(-1*(e.delta/120)), "units"))
        return inner

    # ════════════════════════════════════════════════════════════════════════
    #  TAB 1 — UPLOAD
    # ════════════════════════════════════════════════════════════════════════
    def _tab_upload(self):
        t = self.t_upload

        # API Key row
        api_card = tk.Frame(t, bg=CARD2, padx=20, pady=12,
                            highlightbackground=BORDER, highlightthickness=1)
        api_card.pack(fill="x", padx=30, pady=(16,0))
        tk.Label(api_card, text="🔑  Anthropic API Key (optional — enables AI deep scan)",
                 bg=CARD2, fg=ACCENT, font=FB).pack(anchor="w")
        api_row = tk.Frame(api_card, bg=CARD2)
        api_row.pack(fill="x", pady=(6,0))
        self.api_entry = tk.Entry(api_row, bg="#252525", fg=TEXT,
                                   insertbackground=TEXT, relief="flat",
                                   font=FM, bd=0, show="•")
        self.api_entry.pack(side="left", fill="x", expand=True, ipady=7, ipadx=8)
        if ANTHROPIC_API_KEY:
            self.api_entry.insert(0, ANTHROPIC_API_KEY)
        tk.Button(api_row, text="Apply", bg=ACCENT, fg=TEXT,
                  relief="flat", font=FSM, cursor="hand2",
                  padx=12, pady=7,
                  command=self._apply_api_key).pack(side="left", padx=(6,0))
        tk.Label(api_card,
                 text="Without API key: uses local scoring engine  |  With API key: Claude AI reads & understands your resume deeply",
                 bg=CARD2, fg=SUB, font=FSM).pack(anchor="w", pady=(4,0))

        # Drop zone
        drop = tk.Frame(t, bg=CARD, pady=36, padx=40,
                        highlightbackground=ACCENT, highlightthickness=2)
        drop.pack(fill="x", padx=30, pady=(14,0))
        tk.Label(drop, text="📄", bg=CARD, fg=ACCENT,
                 font=("Segoe UI", 44)).pack()
        tk.Label(drop, text="Upload Your Resume",
                 bg=CARD, fg=TEXT, font=FL).pack(pady=(6,4))
        tk.Label(drop, text="Supported:  PDF  •  DOCX  •  TXT",
                 bg=CARD, fg=SUB, font=FSM).pack()
        tk.Button(drop, text="  📂   Browse File",
                  bg=ACCENT, fg=TEXT, font=FB, relief="flat",
                  cursor="hand2", activebackground=ACCENT2,
                  padx=26, pady=11,
                  command=self._browse).pack(pady=(18,0))

        self.file_lbl = tk.Label(t, text="No file selected",
                                  bg=BG, fg=SUB, font=FSM)
        self.file_lbl.pack(pady=(8,0))

        # Progress card
        prog = tk.Frame(t, bg=CARD, padx=20, pady=14,
                        highlightbackground=BORDER, highlightthickness=1)
        prog.pack(fill="x", padx=30, pady=(12,0))
        self.status_lbl = tk.Label(prog, text="Waiting for resume...",
                                    bg=CARD, fg=SUB, font=FSM)
        self.status_lbl.pack(anchor="w")
        self.prog_var = tk.DoubleVar(value=0)
        ttk.Progressbar(prog, variable=self.prog_var, maximum=100,
                        style="ATS.Horizontal.TProgressbar").pack(
                        fill="x", pady=(8,0))

        # Scan button
        self.scan_btn = tk.Button(
            t, text="  ⚡   SCAN & SCORE MY RESUME",
            bg=ACCENT, fg=TEXT, font=("Segoe UI", 13, "bold"),
            relief="flat", cursor="hand2",
            activebackground=ACCENT2, pady=15,
            state="disabled", command=self._scan)
        self.scan_btn.pack(fill="x", padx=30, pady=(12,8))

        # Dep status
        dep = tk.Frame(t, bg=CARD2, padx=18, pady=10,
                       highlightbackground=BORDER, highlightthickness=1)
        dep.pack(fill="x", padx=30, pady=(0,16))
        tk.Label(dep, text="Library Status", bg=CARD2, fg=ACCENT, font=FB).pack(anchor="w")
        tk.Label(dep,
                 text=("✅ pdfplumber ready" if PDF_OK else "⚠️ pip install pdfplumber  (for PDF support)"),
                 bg=CARD2, fg=(GREEN if PDF_OK else YELLOW), font=FSM).pack(anchor="w",pady=1)
        tk.Label(dep,
                 text=("✅ python-docx ready" if DOCX_OK else "⚠️ pip install python-docx  (for DOCX support)"),
                 bg=CARD2, fg=(GREEN if DOCX_OK else YELLOW), font=FSM).pack(anchor="w",pady=1)

    def _apply_api_key(self):
        global ANTHROPIC_API_KEY
        ANTHROPIC_API_KEY = self.api_entry.get().strip()
        if ANTHROPIC_API_KEY:
            self.ai_badge.config(text="🤖 AI MODE", bg="#4a42cc")
            messagebox.showinfo("API Key Set", "✅ AI Mode enabled!\nClaude will deeply analyze your resume.")
        else:
            self.ai_badge.config(text="🔧 LOCAL MODE", bg="#333")

    # ════════════════════════════════════════════════════════════════════════
    #  TAB 2 — SCORE BREAKDOWN
    # ════════════════════════════════════════════════════════════════════════
    def _tab_score(self):
        t = self.t_score

        # Big score header
        hdr = tk.Frame(t, bg=CARD, pady=22,
                       highlightbackground=BORDER, highlightthickness=1)
        hdr.pack(fill="x", padx=30, pady=(20,0))

        left = tk.Frame(hdr, bg=CARD)
        left.pack(side="left", padx=28)
        self.big_score = tk.Label(left, text="—", bg=CARD, fg=ACCENT, font=FXL)
        self.big_score.pack()
        tk.Label(left, text="/ 100", bg=CARD, fg=SUB, font=FSM).pack()

        right = tk.Frame(hdr, bg=CARD)
        right.pack(side="left", fill="x", expand=True)
        self.grade_lbl = tk.Label(right, text="Scan a resume to see results",
                                   bg=CARD, fg=TEXT, font=FL)
        self.grade_lbl.pack(anchor="w")
        self.grade_msg = tk.Label(right, text="",
                                   bg=CARD, fg=SUB, font=FSM,
                                   wraplength=700, justify="left")
        self.grade_msg.pack(anchor="w", pady=(6,0))

        # Strengths
        self.strength_frame = tk.Frame(hdr, bg=CARD)
        self.strength_frame.pack(side="right", padx=20, anchor="n")

        # Scrollable bars
        self.bars_inner = self._scrollable(t)
        self._bar_widgets = {}

        cats = [
            ("contact_info",   "📱 Contact Info",          15),
            ("sections",       "📂 Resume Sections",        20),
            ("action_verbs",   "🎯 Action Verbs",           15),
            ("tech_keywords",  "💻 Tech Keywords",          20),
            ("quantification", "📊 Quantified Results",     15),
            ("formatting",     "📐 Format & Length",        10),
            ("soft_skills",    "🤝 Soft Skills",             5),
        ]
        for key, label, mx in cats:
            self._make_bar(key, label, mx)

    def _make_bar(self, key, label, mx):
        f = tk.Frame(self.bars_inner, bg=CARD, padx=16, pady=12,
                     highlightbackground=BORDER, highlightthickness=1)
        f.pack(fill="x", pady=(0,7))
        f.columnconfigure(1, weight=1)

        tk.Label(f, text=label, bg=CARD, fg=TEXT, font=FB,
                 width=26, anchor="w").grid(row=0, column=0, sticky="w")

        bv = tk.DoubleVar(value=0)
        ttk.Progressbar(f, variable=bv, maximum=mx,
                        style="ATS.Horizontal.TProgressbar",
                        length=400).grid(row=0, column=1, sticky="ew", padx=10)

        sl = tk.Label(f, text=f"— / {mx}", bg=CARD, fg=SUB, font=FB, width=8)
        sl.grid(row=0, column=2)

        dl = tk.Label(f, text="", bg=CARD, fg=SUB, font=FSM,
                      anchor="w", wraplength=800, justify="left")
        dl.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4,0))

        self._bar_widgets[key] = (bv, sl, dl)

    # ════════════════════════════════════════════════════════════════════════
    #  TAB 3 — SUGGESTIONS
    # ════════════════════════════════════════════════════════════════════════
    def _tab_suggest(self):
        t = self.t_suggest
        hdr = tk.Frame(t, bg=CARD, padx=20, pady=12,
                       highlightbackground=BORDER, highlightthickness=1)
        hdr.pack(fill="x", padx=30, pady=(18,0))
        tk.Label(hdr, text="💡 Personalized Suggestions",
                 bg=CARD, fg=TEXT, font=FL).pack(anchor="w")
        tk.Label(hdr, text="Sorted by priority — fix Critical items first",
                 bg=CARD, fg=SUB, font=FSM).pack(anchor="w", pady=(3,0))

        # Quick wins box
        self.qw_frame = tk.Frame(t, bg="#0d1a0d", padx=20, pady=12,
                                  highlightbackground=GREEN, highlightthickness=1)
        self.qw_frame.pack(fill="x", padx=30, pady=(10,0))
        tk.Label(self.qw_frame, text="⚡ Quick Wins — Do These First",
                 bg="#0d1a0d", fg=GREEN, font=FB).pack(anchor="w")
        self.qw_lbl = tk.Label(self.qw_frame, text="Scan your resume to see quick wins",
                                bg="#0d1a0d", fg=SUB, font=FSM,
                                wraplength=900, justify="left")
        self.qw_lbl.pack(anchor="w", pady=(4,0))

        self.sugg_inner = self._scrollable(t)
        tk.Label(self.sugg_inner,
                 text="Upload and scan your resume to see suggestions here.",
                 bg=BG, fg=SUB, font=FM).pack(pady=30)

    # ════════════════════════════════════════════════════════════════════════
    #  TAB 4 — EXTRACTED TEXT
    # ════════════════════════════════════════════════════════════════════════
    def _tab_text(self):
        t = self.t_text
        hdr = tk.Frame(t, bg=CARD, padx=20, pady=10,
                       highlightbackground=BORDER, highlightthickness=1)
        hdr.pack(fill="x", padx=30, pady=(18,0))
        tk.Label(hdr, text="📝 What ATS Systems Read From Your Resume",
                 bg=CARD, fg=TEXT, font=FL).pack(side="left")

        self.text_box = scrolledtext.ScrolledText(
            t, bg=CARD, fg=TEXT, insertbackground=TEXT,
            font=("Consolas", 10), relief="flat",
            wrap="word", padx=16, pady=14)
        self.text_box.pack(fill="both", expand=True, padx=30, pady=(10,20))
        self.text_box.insert("1.0", "No resume loaded yet.\n\nUpload a file and click Scan to see extracted text.")
        self.text_box.config(state="disabled")

    # ════════════════════════════════════════════════════════════════════════
    #  ACTIONS
    # ════════════════════════════════════════════════════════════════════════
    def _browse(self):
        fp = filedialog.askopenfilename(
            title="Select Your Resume",
            filetypes=[
                ("All supported", "*.pdf *.docx *.txt"),
                ("PDF",  "*.pdf"),
                ("Word", "*.docx"),
                ("Text", "*.txt"),
            ]
        )
        if not fp:
            return
        self._filepath = fp
        size = os.path.getsize(fp)/1024
        self.file_lbl.config(
            text=f"✅  {os.path.basename(fp)}  ({size:.1f} KB)", fg=GREEN)
        self.scan_btn.config(state="normal")
        self.status_lbl.config(text=f"Ready: {os.path.basename(fp)}", fg=SUB)
        self.prog_var.set(0)

    def _scan(self):
        if not self._filepath:
            return
        self.scan_btn.config(state="disabled", text="⏳  Scanning...")
        self.prog_var.set(5)
        self.status_lbl.config(text="Starting scan...", fg=YELLOW)
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        try:
            # Step 1 extract
            self.after(0, lambda: self._upd_status("Extracting text from resume...", 15))
            text = extract_text(self._filepath)

            if text.startswith("ERROR:"):
                self.after(0, self._err, text)
                return

            if not text.strip() or len(text.strip()) < 50:
                self.after(0, self._err,
                    "Could not extract enough text from this file.\n\n"
                    "For PDF: install pdfplumber (pip install pdfplumber)\n"
                    "For DOCX: install python-docx (pip install python-docx)\n\n"
                    "Or convert your resume to .TXT format and try again.")
                return

            self._text = text
            self.after(0, lambda: self._upd_status("Text extracted successfully!", 35))

            # Step 2 score
            api_key = self.api_entry.get().strip()
            if api_key:
                self.after(0, lambda: self._upd_status("🤖 Claude AI is reading your resume...", 50))
                result = call_claude_api(text, api_key)
            else:
                self.after(0, lambda: self._upd_status("Running local ATS scoring engine...", 50))
                result = local_score(text)

            self._result = result
            self.after(0, lambda: self._upd_status("Generating suggestions...", 80))
            self.after(0, self._render, result, text)

        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="ignore")
            self.after(0, self._err, f"API Error {e.code}: {body[:300]}")
        except urllib.error.URLError as e:
            self.after(0, self._err, f"Network error: {e.reason}\nCheck your internet connection.")
        except json.JSONDecodeError as e:
            self.after(0, self._err, f"AI response parse error: {e}\nTry again or use local mode.")
        except Exception as e:
            self.after(0, self._err, str(e))

    def _upd_status(self, msg, pct):
        self.status_lbl.config(text=msg, fg=YELLOW)
        self.prog_var.set(pct)

    def _err(self, msg):
        self.scan_btn.config(state="normal", text="  ⚡   SCAN & SCORE MY RESUME")
        self.prog_var.set(0)
        self.status_lbl.config(text="❌ Error", fg=RED)
        if "NO_DOCX" in msg:
            msg = "python-docx not installed!\n\nFix:\n  pip install python-docx\n\nThen restart the app."
        elif "NO_PDFPLUMBER" in msg or "PDF_EXTRACT_FAILED" in msg:
            msg = ("Could not read PDF text.\n\n"
                   "Fix:\n  pip install pdfplumber\n\n"
                   "Or convert your PDF to .TXT:\n"
                   "• Open PDF → Select All → Copy → Paste into Notepad → Save as resume.txt\n"
                   "Then upload the .txt file.")
        messagebox.showerror("Scan Error", msg)

    def _render(self, r, text):
        self.prog_var.set(100)
        self.status_lbl.config(
            text=f"✅  Done!  ATS Score: {r['total_score']}/100  {r.get('grade_emoji','')}",
            fg=GREEN)
        self.scan_btn.config(state="normal", text="  ⚡   SCAN & SCORE MY RESUME")

        # ── Score tab ─────────────────────────────────────────────────────
        total = r["total_score"]
        clr   = GREEN if total>=80 else (YELLOW if total>=50 else RED)
        self.big_score.config(text=str(total), fg=clr)
        self.grade_lbl.config(
            text=f"{r.get('grade_emoji','')}  {r.get('grade','')}  —  {total}/100",
            fg=clr)
        self.grade_msg.config(text=r.get("grade_message",""))

        # Strengths
        for w in self.strength_frame.winfo_children():
            w.destroy()
        strengths = r.get("strengths", [])
        if strengths:
            tk.Label(self.strength_frame, text="✅ Strengths",
                     bg=CARD, fg=GREEN, font=FB).pack(anchor="w")
            for s in strengths[:4]:
                tk.Label(self.strength_frame, text=f"  • {s}",
                         bg=CARD, fg=SUB, font=FSM,
                         wraplength=280, justify="left").pack(anchor="w")

        # Category bars
        cats = r.get("categories", {})
        details_map = {
            "contact_info":   lambda c: c.get("detail",""),
            "sections":       lambda c: f"Found: {', '.join(c.get('found',[]))}  |  Missing: {', '.join(c.get('missing',[]))}",
            "action_verbs":   lambda c: f"{c.get('count',0)} verbs: {', '.join(c.get('found',[])[:7])}",
            "tech_keywords":  lambda c: f"{c.get('count',0)} skills: {', '.join(c.get('found',[])[:8])}",
            "quantification": lambda c: f"{c.get('count',0)} quantified achievements: {', '.join(str(x) for x in c.get('examples',[])[:4])}",
            "formatting":     lambda c: c.get("detail",""),
            "soft_skills":    lambda c: f"Found: {', '.join(c.get('found',[]))}",
        }
        for key, (bv, sl, dl) in self._bar_widgets.items():
            cat = cats.get(key, {})
            sc  = cat.get("score", 0)
            mx  = cat.get("max", 1)
            bv.set(sc)
            clr2= GREEN if sc>=mx*0.7 else (YELLOW if sc>=mx*0.4 else RED)
            sl.config(text=f"{sc} / {mx}", fg=clr2)
            try:
                dl.config(text=details_map[key](cat))
            except Exception:
                dl.config(text=cat.get("detail",""))

        # ── Suggestions tab ───────────────────────────────────────────────
        # Quick wins
        qw = r.get("quick_wins", [])
        if qw:
            self.qw_lbl.config(text="\n".join(f"  ✅  {w}" for w in qw), fg=GREEN)

        for w in self.sugg_inner.winfo_children():
            w.destroy()

        suggestions = r.get("suggestions", [])
        order = {"Critical":0, "Important":1, "Tip":2}
        suggestions.sort(key=lambda x: order.get(x.get("priority","Tip"),2))

        COLOR  = {"Critical": RED,    "Important": YELLOW, "Tip": GREEN}
        BG_MAP = {"Critical": "#1a0808","Important": "#1a1400","Tip": "#081a10"}
        ICON   = {"Critical": "🔴",    "Important": "🟡",    "Tip": "🟢"}

        for sug in suggestions:
            p    = sug.get("priority","Tip")
            bg_c = BG_MAP.get(p, CARD)
            clr3 = COLOR.get(p, SUB)
            icon = ICON.get(p,"•")

            card = tk.Frame(self.sugg_inner, bg=bg_c, padx=16, pady=11,
                            highlightbackground=clr3, highlightthickness=1)
            card.pack(fill="x", pady=(0,6))

            top = tk.Frame(card, bg=bg_c)
            top.pack(fill="x")
            tk.Label(top, text=f"{icon} {p.upper()}",
                     bg=bg_c, fg=clr3, font=FB).pack(side="left")
            tk.Label(top, text=sug.get("title",""),
                     bg=bg_c, fg=TEXT, font=FB).pack(side="left", padx=(10,0))
            tk.Label(card, text=sug.get("description",""),
                     bg=bg_c, fg=TEXT, font=FSM,
                     wraplength=950, justify="left", anchor="w").pack(
                     anchor="w", pady=(5,0))

        if not suggestions:
            tk.Label(self.sugg_inner, text="No suggestions — great resume!",
                     bg=BG, fg=GREEN, font=FM).pack(pady=20)

        # ── Text tab ──────────────────────────────────────────────────────
        self.text_box.config(state="normal")
        self.text_box.delete("1.0","end")
        self.text_box.insert("1.0", text.strip() or "No text could be extracted.")
        self.text_box.config(state="disabled")

        # Switch to score tab
        self.nb.select(1)


if __name__ == "__main__":
    app = ATSApp()
    app.mainloop()