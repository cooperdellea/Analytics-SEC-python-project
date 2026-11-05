# SEC 10-K Python Code Viewer

GitHub Pages hosts the Python code as a static HTML viewer.

## Publish
1. Create a GitHub repository.
2. Upload `index.html` and `sec_financial_analysis.py` to the repository root.
3. Open Settings > Pages > Deploy from a branch > main > /(root) > Save.
4. Open the resulting GitHub Pages URL.

## Run the code locally
```powershell
py -m pip install requests beautifulsoup4 openpyxl yfinance
py sec_financial_analysis.py
```

SEC financial values and MD&A extraction boundaries require manual verification.
