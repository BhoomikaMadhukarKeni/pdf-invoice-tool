# 🧾 PDF Invoice Tool

A powerful desktop app built using **Python** and **Tkinter** that allows you to:
- Upload one or more **invoice PDFs**
- Automatically extract and process invoice data
- Generate:
  - 📊 An Excel (`.csv`) file listing all items purchased
  - 📈 A **bar chart** comparing item quantities
  - 🖼️ A PowerPoint (`.pptx`) presentation summarizing the analysis

---

## 📂 Sample Flow

1. Place a sample invoice like `original.pdf` (must be text-based, not scanned image) in any folder
2. Run the application
3. Select the invoice(s)
4. Click **“Process PDFs”**
5. The tool extracts data and saves:
   - `invoice_data.csv` – structured data of purchases
   - `chart.png` – visual chart of items vs quantities
   - `invoice_data.pptx` – presentation summary with insights and chart

---

## 📸 Output Example

- **CSV**:
    ```
    Item      | Quantity | Price
    ----------|----------|-------
    Part A    | 10       | $20
    Part B    | 5        | $10
    ```

- **Chart (chart.png)**:
    > A bar chart showing `Item vs Quantity` based on PDF content

- **PPT (invoice_data.pptx)**:
    > A 7-slide presentation covering:
    - Title & Purpose
    - Data analysis
    - Visual chart
    - Improvement suggestions
    - Conclusion

---

## 🛠 Technologies Used

- `pdfplumber` – for PDF text extraction  
- `Tkinter` – for the GUI  
- `matplotlib` – for chart generation  
- `python-pptx` – for creating PowerPoint slides  
- `pandas` – for handling tabular data  
- `Git` – for version control

---

## ▶️ How to Run

1. Clone this repository:
   ```bash
   git clone https://github.com/BhoomikaMadhukarKeni/pdf-invoice-tool.git
   cd pdf-invoice-tool
