
# 💎 RapNet Diamond Exporter (Python Automation)

This is a Python automation tool that connects to the [RapNet](https://www.rapnet.com/) API, fetches diamond listing data based on filters (like shape, size, color, and clarity), and exports it to structured **Excel files** — saving hours of manual work.

---

## 🚀 Features

- Authenticates securely with API token
- Applies dynamic filters: size range, color, clarity
- Exports multi-sheet Excel files (one per filter combo)
- Includes diamond attributes and clean column mapping
- Designed for speed, clarity, and automation

---

## 🛠️ Tech Stack

- Python
- Pandas
- OpenPyXL
- jproperties

---

## 📂 Project Structure

```
rapnet-diamond-exporter/
├── rapnet_diamond_exporter.py       # Main script
├── filter.json                      # RapNet filter template
├── market_input_sample.txt          # Sample input config (safe to share)
├── .gitignore                       # Files to ignore
└── README.md                        # You're reading it
```

---

## ⚙️ How to Run

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/rapnet-diamond-exporter.git
   cd rapnet-diamond-exporter
   ```

2. **Install dependencies**
   ```bash
   pip install requests pandas openpyxl jproperties
   ```

3. **Add configuration**
   - Copy `market_input_sample.txt` ➜ rename to `market_input.txt`
   - Replace `YOUR_API_TOKEN_HERE` with your actual token

4. **Custom `filter.json`**
   - Replace `"your_account_id"` and `"your_contact_id"` with actual values

5. **Run the script**
   ```bash
   python rapnet_diamond_exporter.py
   ```

6. **Excel output** will be saved in the current directory.

---

## 🧪 Sample `market_input_sample.txt`

```ini
load_saved_search = EMERALD LG (GD)
size_range = 1.50:1.69
colors = D,E,F,G,H,I,J,K,L,M
clarities = IF,VVS1,VVS2,VS1,VS2,SI1,SI2
token = YOUR_API_TOKEN_HERE
```

---

## 🔐 Security Best Practices

- Do **not commit** your `market_input.txt` file
- Use `.gitignore` to protect sensitive or generated files:

```txt
market_input.txt
*.xlsx
__pycache__/
```

---

## 📦 Output

Each Excel file includes:
- Top rows: diamond search attributes (shape, color, clarity, etc.)
- Below: listing data (price, lab, measurements, etc.)
- One sheet per combination (e.g. `EMERALD_F_VS1`)

---

## 👤 About the Author

**Satya Sirisha Bolloju**   
🔗 [LinkedIn](www.linkedin.com/in/satya-sirisha-bolloju-031b33239)

---

## 📄 License

MIT License — free to use, modify, and contribute.

---

## ⭐ Feedback & Contributions

If you find this useful:
- Star the repo
- Fork the code
- Or open an Issue / Pull Request 

---


