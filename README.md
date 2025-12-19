# 📊 Mplus2Excel – Competing Measurement Models & Factor Loadings (2025)

### **Automatically generate publication-ready tables from Mplus output files**

---

## 📌 What is this project?

Comparing competing measurement models in **Mplus** is essential, but extracting fit indices and standardized factor loadings manually is time-consuming and error-prone.

**Mplus2Excel** is a lightweight utility that scans a folder of **Mplus measurement model `.out` files**, automatically:

- Aggregates fit indices across all competing models
- Produces a clean **comparison table** of measurement models
- Extracts **STDYX standardized factor loadings**
- Outputs the **three best-fitting models** for inspection and reporting

All results are written directly to a structured, publication-ready **Excel workbook**.

---

## 🎯 What does it do exactly?

When run in a folder containing *only* Mplus **measurement model output files**, the tool will:

1. **Scan all `.out` files** in the folder
2. **Extract key fit indices** (e.g. χ², df, CFI, TLI, RMSEA, SRMR, AIC, BIC)
3. **Correlate all models into a single comparison table**
4. **Rank models by fit**
5. **Extract STDYX standardized factor loadings** for the **top 3 models**
6. Export everything to a single Excel workbook

This eliminates repetitive copy-paste work and reduces transcription errors.

---

## ⚠ Important usage constraint (read this)

> **The folder must contain only Mplus measurement model `.out` files.**

- Do **not** include structural models
- Do **not** include auxiliary analyses
- Do **not** mix different projects in the same folder

The script intentionally aggregates **all detected output files** into a single comparison table. Mixing analyses will contaminate results.

---

## 🤔 Why is this useful?

I built this tool after repeatedly wasting hours manually assembling model comparison tables and factor loading matrices for papers, reviews, and reports.

✅ Saves substantial time  
✅ Eliminates transcription errors  
✅ Produces consistent, reproducible tables  
✅ Ideal for exploratory and confirmatory measurement work  
✅ Designed for researchers who use Mplus heavily  

---

## 🛠 How to use (Windows – EXE)

1. Download `mplus2excel.exe` from **GitHub Releases**
2. Place the EXE in the folder containing your Mplus `.out` files
3. Double-click the EXE
4. An Excel workbook and a `run.log` file will be created in the same folder

If something goes wrong, the log file can be shared for troubleshooting.

---

## 🛠 How to use (Python)

```bash
pip install -r requirements.txt
python mplus2excel.py --input "." --output "Competing_Measurement_Models.xlsx" ```


Optional arguments:
--pattern "*.out" → restrict files
--top 3 → number of best-fitting models to extract

---

## 📂 What the Excel file contains

Instructions tab (factor labels and mappings)

Table 1: Competing measurement model fit indices

Table 2: STDYX standardized factor loadings (top 3 models)

Audit tab: Parsed values for verification and transparency

---

## 🔧 **Where to Get Help**  
💬 **[Open an Issue](https://github.com/llewellynvz/Mplus-to-Excel-Table-Generator/issues)** on GitHub  
✉️ **Contact me via my website** [psynalytics.com](https://www.psynalytics.com)  

---

## 👨‍💻 **Who Maintains This?**  
This project was created by **Prof. Llewellyn E. van Zyl** out of frustration with LinkedIn’s terrible feed algorithm.  

- 🌍 **Website:** [www.psynalytics.com](https://www.psynalytics.com)  
- 🔗 **GitHub:** [@llewellynvz](https://github.com/llewellynvz) 
- 

🔥 If you’d like to **contribute**, feel free to fork this repo, submit a pull request, or report bugs!  

---

## 📜 License

This software is proprietary. Use and local modification are permitted for
personal, academic, and internal research purposes only. Redistribution or
commercial use is not permitted.

See the LICENSE file for details.

---

## ⭐ If you found this useful

Star the repository ⭐

Share it with colleagues using Mplus

Contribute improvements or edge-case fixes