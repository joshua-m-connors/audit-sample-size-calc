# Audit Sample Size Calculator 🧮

For any friends out there in the audit field this is a modern Python GUI application for calculating **audit sampling sizes** under standard statistical and auditing assumptions (AU-C 530 / PCAOB AS 2315).

## ✨ Features

- **Audit-style planning** using a binomial approximation.
- **Rollforward testing support**:
  - Uses **observed interim deviation rate** (exceptions / interim sample) to re-plan the full-year sample.
  - Ensures **proportional coverage** for the rollforward population.
- **Professional output**:
  - Results Summary panel (base, replanned, rollforward, adjusted totals).
  - PDF and CSV export.
- **Modern UI**: Dark/light awareness in Windows, tooltips, and responsive layout.

## 📦 Requirements

- Python 3.8+
- `reportlab` for PDF export:
  ```bash
  pip install reportlab
  ```

## 🚀 Usage

1. Run the application:
   ```bash
   python audit_sample_size_calculator.py
   ```

2. Enter inputs:
   - **Population Size**: total number of items in scope.
   - **Confidence Level**: 90%, 95%, or 99%.
   - **Expected Deviation Rate**: default `0.00` for most control testing.
   - **Tolerable Deviation Rate**: default `0.05`.

3. (Optional) **Rollforward Testing**:
   - Check **Include Rollforward Testing**.
   - **Issues found in interim testing?** If *Yes*, enter **Interim Sample Size** and **Exceptions Count**.
   - Enter **Rollforward Population Size**.
   - Click **Calculate** to see:
     - **Base sample size (no rollforward)**
     - **Replanned full-year required sample**
     - **Additional rollforward sample required**
     - **Adjusted total sample (interim + rollforward)**

4. **Export** results to **PDF** or **CSV** via the **Export** button.

## 🧠 Methodology

- **Base sample size** is computed via a **binomial approximation**:
  \n
  \\[ n \\approx \\left\\lceil \\frac{\\ln(1-\\text{confidence})}{\\ln(1-\\text{tolerable deviation})} \\right\\rceil \\]\n
  (For small populations `< 1000`, a finite population correction is applied.)\n
- **Expected deviation** (entered by the user) is **raised** to at least the **observed interim deviation** (exceptions / interim sample) when rollforward is enabled.\n
- **Replanned full-year required sample** is recomputed using this **planned expected deviation**.\n
- **Additional rollforward sample** is the maximum of:\n
  - `replanned total – interim sample size`, and\n
  - `proportional minimum for the rollforward population`.\n
- If exceptions are found at interim, apply **professional judgment** for any additional procedures (AU-C 530 / PCAOB AS 2315).\n

## 🧪 Defaults (typical audit assumptions)

- Confidence: **90%**
- Expected deviation: **0%**
- Tolerable deviation: **5%**
- These yield a sample size of **≈ 44** (aligns with common audit sampling tables).

## 🛠 Troubleshooting

- **PDF export fails**: Ensure `reportlab` is installed (`pip install reportlab`).
- **Blurry UI on Windows**: High-DPI is enabled in the app; verify your display scaling.
- **Exceptions field not showing**: It appears only when *Issues found? = Yes*.

## GUI Image

<img width="622" height="934" alt="Screenshot From 2025-10-31 13-37-23" src="https://github.com/user-attachments/assets/184eb5cb-e9d6-40a4-979e-8e871b6a6e0c" />


## 📄 License

MIT — use freely in projects. Feedback and PRs welcome!
