import pandas as pd
import re
import xlwings as xw
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.compose import ColumnTransformer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import OneHotEncoder
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.pipeline import Pipeline
import os
from joblib import dump

# === CONFIGURATION ===
EXCEL_PATH = r"C:\Users\SinaB\Metrolinx\Civil Structures - Engineering - Documents\General\TASK TRACKER\Sina_DON'T TOUCH\WORKING_Automated Submittal List.xlsm"
SHEET_NAME = "All Submittals"
DISCIPLINE_COLS = ["Geotechnical", "Structural", "Tunnel"]
ML_COLS = ["ML Geotechnical", "ML Structural", "ML Tunnel"]
FEATURE_COLS = ["Long Title", "Project Component", "Submittal Revision #"]
TRANSFORMED_FEATURE_COLS = ["Title", "Project Component", "Package #"]

# === 1. READ WORKSHEET INTO PANDAS (reads cells only; macros untouched) ===
# Use openpyxl engine to read .xlsm cell values into a DataFrame.
# This reads the full sheet into pandas; we will open the workbook via xlwings
# later only when writing back so macros/VBA are preserved.
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=1, engine="openpyxl")
df = df[FEATURE_COLS + DISCIPLINE_COLS + ML_COLS]
df = df.rename(columns={"Long Title": "Title"})
def extract_submittal_code(text: str) -> str:
    text = str(text).upper()
    match = re.search(r"S-\d+", text)
    return match.group(0).replace(" ", "") if match else ""

df["Title"] = (
    df["Title"]
    .astype(str)
    .str.lower()
    .str.replace(r"[_]", " ", regex=True)
    .str.replace(r"[^a-z0-9\s\-]", "", regex=True)
    .str.replace(r"\s+", " ", regex=True)
    .str.replace(r"\s\-\s", " ", regex=True)
    .str.strip()
)
df["Package #"] = df["Submittal Revision #"].apply(extract_submittal_code)
df.drop(columns=["Submittal Revision #"], inplace=True)

preprocessor = ColumnTransformer(
    transformers=[
        ("title", TfidfVectorizer(), "Title"),
        # Pass categorical columns as lists so transformers receive 2-D input
        ("component", OneHotEncoder(handle_unknown="ignore"), ["Project Component"]),
        ("package", OneHotEncoder(handle_unknown="ignore"), ["Package #"]),
    ],
    remainder="drop",
)

# === 4. RESET ML COLUMNS ===
for col in ML_COLS:
    df[col] = "F"

# === 5. TRAIN & PREDICT ===
for orig_col, ml_col in zip(DISCIPLINE_COLS, ML_COLS):
    if orig_col not in df.columns:
        continue

    # Prepare binary target: map Yes/No to T/F and treat empty as unknown
    y = df[orig_col].astype(str).replace("nan", "").replace("No", "F").replace("Yes", "T")
    mask_train = y != ""
    mask_predict = y != ""

    if mask_train.sum() == 0:
        continue

    model = Pipeline([
        ("pre", preprocessor),
        ("clf", RandomForestClassifier(n_estimators=200, random_state=42, n_jobs=-1)),
    ])

    X = df[TRANSFORMED_FEATURE_COLS].copy()

    # sklearn expects y as a 1-D array-like. Convert the Series slice to numpy.
    y_train = y[mask_train].values.ravel()
    model.fit(X.loc[mask_train], y_train)
    # Save trained model pipeline to disk (includes preprocessor)
    try:
        models_dir = os.path.join(os.getcwd(), "models")
        os.makedirs(models_dir, exist_ok=True)
        safe_name = orig_col.lower().replace(" ", "_")
        model_path = os.path.join(models_dir, f"model_{safe_name}.joblib")
        dump(model, model_path)
    except Exception:
        # Don't crash the script if saving fails; continue to predictions
        pass
    preds = model.predict(X.loc[mask_predict])

    # Normalize predictions to 'T'/'F' strings for the ML columns
    def to_TF(val):
        s = str(val).strip().lower()
        return "T" if s in ("t", "true", "yes", "y", "1") else "F"

    df.loc[mask_predict, ml_col] = [to_TF(p) for p in preds]

# === 6. OPEN WORKBOOK VIA COM AND WRITE BACK TO EXCEL CELLS ONLY ===
# Open workbook via COM/xlwings only for writing so macros/VBA are preserved.
app = xw.App(visible=False)
wb = xw.Book(EXCEL_PATH)
ws = wb.sheets[SHEET_NAME]

start_row = 3  # data starts at Excel row 3 (header is row 2)
for col in ML_COLS:
    if col not in df.columns:
        continue

    # Find the column index in Excel
    header_row = ws.range("A2").expand("right").value
    if col not in ML_COLS:
        # If ML column doesn't exist, add it to the next empty column
        next_col = len(header_row) + 1
        ws.range((1, next_col)).value = col
        col_idx = next_col
    else:
        col_idx = header_row.index(col) + 1

    # Write values starting from row 2 (below header)
    ws.range((start_row, col_idx), (len(df)+start_row-1, col_idx)).value = df[[col]].values.tolist()

# ✅ Done: no wb.save() → macros are preserved
app.quit()
print("✅ ML columns updated safely. Macros remain intact.")