# NIH Training Grant Tools

A modular toolkit developed to support NIH training grant applications, including trainee-publication matching, Table 5 generation, and biosketch verification. Created as part of the Research Development Certificate internship at Stanford.

---

## 🧩 Available Modules

| Module | Description | Status |
|--------|-------------|--------|
| **Publication Matcher** | Matches trainees to mentor publications (T32 support) | ✅ Working |
| **Table 5 Generator** | Generates NIH-format Table 5 from trainee input | 🛠️ In progress |
| **Biosketch Checker** | LLM-powered validation tool for NIH biosketches | 🧪 Planned |

---

## 📦 Module Details

### 🧑‍🎓 1. Publication Matcher

**Goal:** Identify and bold trainee names in a mentor's CV publication list.

**Inputs:**
- `.docx`: Publication list (mentor CV)
- `.xlsx`: Trainee list

**Output:**
- `.docx`: Table with trainee info and matched publications (names bolded)

→ [Usage Instructions](modules/publication_matcher/README.md)

---

### 📊 2. Table 5 Generator

**Goal:** Automatically generate Table 5 (T32) from structured trainee and mentor data.

**Inputs:**
- `.xlsx`: Extended trainee+mentor+outcomes list

**Output:**
- `.docx` or `.xlsx`: NIH-formatted Table 5

→ Coming soon...

---

### 📝 3. Biosketch Checker

**Goal:** Run LLM-based checks on NIH biosketches for:
- Format
- Length
- Content alignment
- Missing sections

**Inputs:**
- `.docx`: NIH biosketch draft

**Output:**
- Text-based feedback or annotated version

→ Coming soon...

---

## 🔧 How to Use

Each module runs independently via its own Python or Colab notebook. See the corresponding module folder for instructions and sample data.

---

## 📜 License

MIT License

---

## 🙋‍♂️ Author

Uğur Aygün  
Marie Curie Fellow · Postdoc @ Stanford · Research Assistant Professor @ Koç University  
