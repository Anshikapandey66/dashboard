# 📊 Expense Tracker Excel Dashboard

A simple and interactive **Expense Tracker Dashboard** built using **Microsoft Excel**. This project helps users analyze their daily expenses, spending categories, and payment methods through a clean dashboard.

## 🚀 Features

- Track daily expenses
- Categorize expenses
- Calculate total expenses
- Calculate average expenses
- Find highest expense
- Analyze expenses by category
- Analyze expenses by payment method
- Interactive Excel dashboard
- Pie chart for category-wise expenses
- Bar chart for payment-method analysis

## 🛠️ Technologies Used

- Microsoft Excel
- Python
- Pandas
- OpenPyXL
- Excel Charts & Formulas

## 📁 Project Structure

```text
expense-tracker-excel-dashboard/
│
├── create_dashboard.py
├── Expense_Tracker.xlsx
├── requirements.txt
└── README.md
```

## 📊 Dashboard

The dashboard contains:

- **Total Expense**
- **Average Expense**
- **Highest Expense**
- **Total Records**
- **Expense by Category**
- **Expense by Payment Method**

## 📌 Excel Formulas Used

### Total Expense

```excel
=SUM(D2:D21)
```

### Average Expense

```excel
=AVERAGE(D2:D21)
```

### Highest Expense

```excel
=MAX(D2:D21)
```

### Category-wise Expense

```excel
=SUMIF(B:B,"Food",D:D)
```

## ▶️ How to Run

### 1. Clone the repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Python script

```bash
python create_dashboard.py
```

The script will automatically generate:

```text
Expense_Tracker.xlsx
```

## 🎯 Project Objective

The main objective of this project is to demonstrate practical skills in **Excel data analysis, dashboard creation, data visualization, and basic Python automation**.

## 👩‍💻 Author

**Anshika Pandey**

GitHub: **Anshikapandey66**

## ⭐ Future Improvements

- Add monthly expense tracking
- Add budget vs actual analysis
- Add interactive Excel slicers
- Add monthly trend charts
- Add automatic expense alerts

---

⭐ If you find this project useful, consider giving it a star!
