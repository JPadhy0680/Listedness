
# LLT Extractor App

A simple Streamlit app to:
- upload a MedDRA master Excel file containing `LLT Code`, `LLT Term`, `PT Code`, `PT Term`
- upload a PT list Excel file containing `PT Code` and/or `PT Term`
- display all matching LLTs
- download the output as an Excel workbook

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Output sheets
- `Matched_LLTs`
- `Unmatched_PTs`
- `Summary`
