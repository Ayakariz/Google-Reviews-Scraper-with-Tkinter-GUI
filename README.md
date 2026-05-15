# Google-Reviews-Scraper-with-Tkinter-GUI

A Python desktop application to scrape reviews from Google Maps with a user-friendly GUI interface. Supports automatic Google login, precise review data extraction with accurate timestamps, and export to CSV/Excel/JSON.

---

## ✨ Key Features

### 1. **Place ID Search**
- Search locations on Google Maps using Google Places API (New)
- View ratings, review counts, and addresses
- Support for multi-language search queries

### 2. **Web Scraping Reviews**
- Automatic Google login via Selenium
- Auto-scroll to load all reviews
- Extract complete review data:
  - Reviewer name
  - Rating (stars)
  - Review text
  - **Precise timestamp (DD-MM-YYYY HH:MM WIB)** via CDP Network interception
  - Outscraper format compatible (MM/DD/YYYY UTC)

### 3. **Modular Tkinter GUI**
- Tabs: **Place ID Finder**, **Scraper**, **Settings**
- Multi-language support (English/Indonesian)
- Real-time log viewer
- Progress indicator

### 4. **Flexible Export**
- CSV (.csv) — minimal columns (name, rating, text, date)
- Excel (.xlsx) — auto-formatted with freeze panes & styling
- JSON (.json) — raw data with full metadata

### 5. **Environment Setup**
- Save API key in `.env` (secure, not hardcoded)
- Support for webdriver-manager or manual ChromeDriver
- Automatic internet connectivity check

---

## 🛠️ Requirements

### Main Dependencies
```
requests>=2.31.0          # HTTP API calls
selenium>=4.0.0           # Web scraping
webdriver-manager>=4.0.0  # ChromeDriver auto-download
python-dotenv>=1.0.0      # .env file support
openpyxl>=3.0.0          # Excel export
python-dateutil>=2.8.2    # Relative date handling (optional but recommended)
```

### System Requirements
- **Python**: 3.8+
- **Chrome/Chromium**: Latest version (auto-download via webdriver-manager)
- **Internet**: Active connection (for API & scraping)

---

## 📦 Installation

### 1. Clone Repository
```bash
git clone https://github.com/yourusername/grs-gui.git
cd grs-gui
```

### 2. Create Virtual Environment
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

**File `requirements.txt`:**
```
requests>=2.31.0
selenium>=4.0.0
webdriver-manager>=4.0.0
python-dotenv>=1.0.0
openpyxl>=3.0.0
python-dateutil>=2.8.2
```

### 4. Setup Google API Key (Optional for Place ID Search)

**If you want to use the "Search Place ID" feature:**

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable API: **Places API (New)**
4. Create an API key in Credentials
5. Save the key in `.env` file in the same folder as the script:
   ```
   GOOGLE_PLACES_API_KEY=your_api_key_here
   ```

**Note:** The scraper works without an API key — you can directly input a known Place ID.

---

## 🚀 Usage

### Run the Script
```bash
python "GRS GUI ver.py"
```

The GUI will open with 3 tabs:

### Tab 1: **Place ID Finder** 🔍
1. Enter API key (or skip if already in `.env`)
2. Type location name (e.g., "Starbucks Jakarta")
3. Click **Search Place ID**
4. Select result, copy Place ID

### Tab 2: **Scraper** 🕷️
1. Paste Place ID from Tab 1 (or search manually on Google Maps)
2. Click **Start Scraper**
3. Browser opens — **login to your Google account**
4. Script auto-scrolls and scrapes all reviews
5. **Timestamp extracted via CDP Network — second-level precision**
6. Click **Export** to save as CSV/Excel/JSON

### Tab 3: **Settings** ⚙️
- Set language (English/Indonesian)
- View/update API key
- Dependency information
- Health check (internet, Chrome, dependencies)

---

## 📊 Output Data (CSV/Excel)

**Exported columns:**
```
username          | Reviewer name
rating            | Stars (1-5)
review_text       | Review content
review_date_wib   | DD-MM-YYYY HH:MM (Western Indonesian Time)
review_date_utc   | MM/DD/YYYY HH:MM:SS (UTC, Outscraper compatible)
review_timestamp  | UNIX seconds
review_id         | Unique review ID (from data-review-id)
```

**Example CSV:**
```csv
username,rating,review_text,review_date_wib,review_date_utc,review_timestamp,review_id
John Doe,5,"Great coffee and service",15-05-2026 14:30,05/15/2026 07:30:00,1747483800,Ci90tAMQACoDcHtyuF1oO:BEWEFva1JxcFU1YbZkQ3V...
Jane Smith,4,"Good but a bit crowded",14-05-2026 10:15,05/14/2026 03:15:00,1747397700,ChdDSIHM0ogKEICAgICZwIak...
```

---

## 🔐 Security & Privacy

- **API Key**: Stored in `.env` (git-ignored, not hardcoded)
- **Local Data**: All data saved locally in CSV/Excel, not sent to external servers
- **Selenium**: Uses CDP for network response capture instead of DOM parsing (more accurate)
- **User Agent**: Stealth mode enabled to avoid bot detection

---

## 🛡️ Error Handling

The application handles:
- ❌ No internet connection → clear message + solution suggestions
- ❌ ChromeDriver not found → fallback to local or auto-download
- ❌ Invalid API Key → error message + Google Cloud Console link
- ❌ Login timeout → manual wait + option to continue
- ❌ Network timeout → retry logic

---

## 📝 Tips & Tricks

### 1. If Login Times Out
Script waits 120 seconds for login. If timeout, a popup appears — click OK to continue.

### 2. If Reviews Not Scraped
- Ensure scroll to bottom completes (check logs)
- Some places have anti-bot protection — try disabling JavaScript in Selenium options
- If CDP timestamps empty, fallback to regex parsing

### 3. Manual ChromeDriver
If webdriver-manager fails:
1. Download from: https://chromedriver.chromium.org/downloads
2. Version must match your Chrome version (`chrome://version/`)
3. Copy `chromedriver.exe` to the same folder as the script
4. Script auto-detects & uses it

### 4. Export Excel vs CSV
- **CSV**: Lightweight, open in Excel/Google Sheets
- **Excel**: Formatted with freeze panes, color coding
- **JSON**: Raw data + metadata, for integration

---

## 🐛 Troubleshooting

| Problem | Solution |
|---------|----------|
| "ModuleNotFoundError: No module named 'requests'" | `pip install requests` |
| ChromeDriver version mismatch | Check `chrome://version/`, download matching version |
| API Key error 403 | Enable "Places API (New)" in Google Cloud Console |
| Timestamp empty | CDP may be disabled — check Chrome version, fallback regex |
| Login stuck | Check internet connection, retry, or login manually first |
| Memory leak on long scrape | Reduce scroll pause time, limit reviews per scrape |

---

## 🔄 Typical Workflow

```
1. [Place ID Finder] → Search location → Copy Place ID
                  ↓
2. [Scraper] → Paste ID → Click Start → Login → Auto-scrape
                  ↓
3. [Export] → Choose format (CSV/Excel/JSON) → Save
                  ↓
4. [Analyze] → Open in Excel/Python pandas → Visualize trends
```

---

## 📦 File Structure

```
grs-gui/
├── GRS GUI ver.py          # Main application
├── requirements.txt         # Dependencies
├── .env                     # API key (auto-created)
├── .gitignore              # Ignore .env, venv, exports/
├── chromedriver.exe        # Optional: manual ChromeDriver
├── README.md               # This file
└── exports/
    ├── reviews_2026-05-15.csv
    ├── reviews_2026-05-15.xlsx
    └── reviews_2026-05-15.json
```

---

## 🚦 Development & Contribution

### Roadmap
- [ ] Multi-language review translation
- [ ] Proxy rotation for large-scale scraping
- [ ] Database backend (SQLite/PostgreSQL)
- [ ] Scheduled scraping (cron jobs)
- [ ] Docker support
- [ ] REST API wrapper
- [ ] Sentiment analysis integration

### How to Contribute
1. Fork the repository
2. Create feature branch (`git checkout -b feature/feature-name`)
3. Commit changes (`git commit -m "Add: description"`)
4. Push branch (`git push origin feature/feature-name`)
5. Create Pull Request

---

## 📄 License

MIT License — Free for personal, commercial, and modification use.

---

## ⚠️ Disclaimer

- **Respect Google ToS**: This script is for educational purposes. Ensure you have permission before large-scale scraping.
- **Rate Limiting**: Use responsibly — don't DDoS Google Maps API.
- **Data Privacy**: If exporting review data, respect reviewer privacy.
- **Compliance**: Some regions may have legal restrictions on web scraping — check local laws.

---

## 🙏 Credits

- **Selenium**: Web automation framework
- **Google Places API**: Location search
- **Chrome DevTools Protocol (CDP)**: Network interception
- **tkinter**: GUI (Python built-in)
- **openpyxl**: Excel export
- **python-dotenv**: Environment variable management


Made with ❤️ for the community
