# NIH Training Grant Tools

A modular toolkit to automate and streamline key parts of NIH training grant preparation, including trainee–publication matching, Table 5 generation, and biosketch verification.

---

## 📦 Available Modules

| Module | Description | Status |
|--------|-------------|--------|
| **Table 5A/5B Merger** | Upload single or multiple Excel files from mentors containing trainee data, plus corresponding publication CSV files, to generate a single NIH-formatted DOCX file with merged Table 5A and Table 5B | ✅ Working |
| **Publication Matcher** | Matches trainees to mentor publications and highlights their names in publication lists | ✅ Working |
| **Biosketch Checker** | LLM-powered validation tool for NIH biosketches (format, length, content alignment, missing sections) | 🛠️ In progress |

---

## ⚙️ How It Works

1. **Upload mentors’ trainee tables** (`.xlsx`) — may include Table 5A, Table 5B, or both.
2. **Upload mentors’ publication lists** (`.csv` exports from publication databases).
3. The app will:
   - Merge all Table 5A sheets together.
   - Merge all Table 5B sheets together.
   - Deduplicate trainees and merge mentor lists.
   - Sort trainees by **mentor surname** and then **trainee surname**.
4. **Output:**
   - One `.docx` file with two sections: Table 5A and Table 5B.
   - *(Optional)* Analytics summary: % of trainees with first-author papers, total publications, and publication year range — **not included in DOCX**.

---

## 🚀 How to Run

### 1. Local Installation
```bash
git clone https://github.com/your-username/nih-training-grant-tools.git
cd nih-training-grant-tools
pip install -r requirements.txt
streamlit run run_streamlit.py
```

### 2. Google Colab
Open the Colab notebook in the modules/ folder and follow the upload prompts.

### 3. Online (Streamlit Cloud)
https://table5ab.streamlit.app
⸻

# 📄 License
MIT License.

---

# 🤝 Contributors
- **Ugur Aygun** 
- **Justin Crest**

