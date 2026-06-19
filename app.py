
import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="LLT Extractor", layout="wide")

REQUIRED_MASTER_COLS = ["LLT Code", "LLT Term", "PT Code", "PT Term"]
POSSIBLE_PT_COLS = ["PT Code", "PT Term"]


def clean_header(col):
    col = str(col).strip()
    col = re.sub(r"\s+", " ", col)
    return col


def normalize_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.casefold()
    )


def read_excel_any(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = xls.sheet_names
        if len(sheet_names) == 1:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            return df, sheet_names[0]
        return None, sheet_names
    except Exception as e:
        raise ValueError(f"Unable to read the Excel file: {e}")


def validate_and_prepare_master(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_header(c) for c in df.columns]
    missing = [c for c in REQUIRED_MASTER_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Master file is missing required columns: {', '.join(missing)}")

    out = df[REQUIRED_MASTER_COLS].copy()
    for c in ["LLT Code", "PT Code"]:
        out[c] = out[c].astype(str).str.strip()
    for c in ["LLT Term", "PT Term"]:
        out[c] = out[c].fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    out["_pt_code_key"] = normalize_text(out["PT Code"])
    out["_pt_term_key"] = normalize_text(out["PT Term"])
    return out


def validate_and_prepare_pt_list(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_header(c) for c in df.columns]
    available = [c for c in POSSIBLE_PT_COLS if c in df.columns]
    if not available:
        raise ValueError("PT list file must contain at least one of these columns: PT Code, PT Term")

    keep_cols = [c for c in POSSIBLE_PT_COLS if c in df.columns]
    out = df[keep_cols].copy()
    if "PT Code" in out.columns:
        out["PT Code"] = out["PT Code"].astype(str).str.strip()
        out["_pt_code_key"] = normalize_text(out["PT Code"])
    else:
        out["PT Code"] = ""
        out["_pt_code_key"] = ""

    if "PT Term" in out.columns:
        out["PT Term"] = out["PT Term"].fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
        out["_pt_term_key"] = normalize_text(out["PT Term"])
    else:
        out["PT Term"] = ""
        out["_pt_term_key"] = ""

    out = out[(out["_pt_code_key"] != "") | (out["_pt_term_key"] != "")].copy()
    out = out.drop_duplicates(subset=["_pt_code_key", "_pt_term_key"]).reset_index(drop=True)
    return out


def match_llts(master: pd.DataFrame, pt_list: pd.DataFrame):
    result_frames = []
    unmatched_records = []

    for _, row in pt_list.iterrows():
        pt_code_key = row.get("_pt_code_key", "")
        pt_term_key = row.get("_pt_term_key", "")

        match = pd.DataFrame()
        if pt_code_key:
            match = master[master["_pt_code_key"] == pt_code_key].copy()

        if match.empty and pt_term_key:
            match = master[master["_pt_term_key"] == pt_term_key].copy()

        if not match.empty:
            match.insert(0, "Requested PT Code", row.get("PT Code", ""))
            match.insert(1, "Requested PT Term", row.get("PT Term", ""))
            result_frames.append(match)
        else:
            unmatched_records.append({
                "Requested PT Code": row.get("PT Code", ""),
                "Requested PT Term": row.get("PT Term", "")
            })

    if result_frames:
        matched = pd.concat(result_frames, ignore_index=True)
        matched = matched[
            [
                "Requested PT Code", "Requested PT Term",
                "LLT Code", "LLT Term", "PT Code", "PT Term"
            ]
        ].drop_duplicates().reset_index(drop=True)
    else:
        matched = pd.DataFrame(columns=[
            "Requested PT Code", "Requested PT Term",
            "LLT Code", "LLT Term", "PT Code", "PT Term"
        ])

    unmatched = pd.DataFrame(unmatched_records)

    if not matched.empty:
        summary = (
            matched.groupby(["PT Code", "PT Term"], dropna=False)
            .size()
            .reset_index(name="LLT Count")
            .sort_values(["PT Term", "PT Code"], kind="stable")
            .reset_index(drop=True)
        )
    else:
        summary = pd.DataFrame(columns=["PT Code", "PT Term", "LLT Count"])

    return matched, unmatched, summary


def build_output_excel(matched: pd.DataFrame, unmatched: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        matched.to_excel(writer, index=False, sheet_name="Matched_LLTs")
        unmatched.to_excel(writer, index=False, sheet_name="Unmatched_PTs")
        summary.to_excel(writer, index=False, sheet_name="Summary")
    buffer.seek(0)
    return buffer.getvalue()


def display_instructions():
    st.markdown(
        """
        ### How to use
        1. Upload the **MedDRA master file** with columns: `LLT Code`, `LLT Term`, `PT Code`, `PT Term`
        2. Upload the **PT list file** with at least one of: `PT Code`, `PT Term`
        3. Click **Extract LLTs**
        4. Review the results and click **Download results**
        """
    )


st.title("LLT Extractor")
st.caption("Upload a MedDRA master file and a PT list to fetch all matching LLTs.")

display_instructions()

with st.sidebar:
    st.header("Files")
    master_file = st.file_uploader("Upload MedDRA master Excel", type=["xlsx", "xls"])
    pt_file = st.file_uploader("Upload PT list Excel", type=["xlsx", "xls"])
    extract_btn = st.button("Extract LLTs", type="primary", use_container_width=True)

col1, col2 = st.columns(2)
with col1:
    st.subheader("Expected master format")
    st.dataframe(
        pd.DataFrame([
            {
                "LLT Code": "10000001",
                "LLT Term": "Ventilation pneumonitis",
                "PT Code": "10081988",
                "PT Term": "Hypersensitivity pneumonitis",
            }
        ]),
        use_container_width=True,
        hide_index=True,
    )
with col2:
    st.subheader("Expected PT list format")
    st.dataframe(
        pd.DataFrame([
            {"PT Code": "10081988", "PT Term": "Hypersensitivity pneumonitis"},
            {"PT Code": "", "PT Term": "Hypersensitivity pneumonitis"},
        ]),
        use_container_width=True,
        hide_index=True,
    )

if extract_btn:
    if master_file is None or pt_file is None:
        st.error("Please upload both the MedDRA master file and the PT list file.")
    else:
        try:
            master_df_raw = pd.read_excel(master_file, engine="openpyxl")
            pt_df_raw = pd.read_excel(pt_file, engine="openpyxl")

            master_df = validate_and_prepare_master(master_df_raw)
            pt_df = validate_and_prepare_pt_list(pt_df_raw)
            matched_df, unmatched_df, summary_df = match_llts(master_df, pt_df)

            st.success(f"Extraction completed. Matched rows: {len(matched_df)} | Unmatched PTs: {len(unmatched_df)}")

            tab1, tab2, tab3 = st.tabs(["Matched LLTs", "Unmatched PTs", "Summary"])
            with tab1:
                st.dataframe(matched_df, use_container_width=True, hide_index=True)
            with tab2:
                st.dataframe(unmatched_df, use_container_width=True, hide_index=True)
            with tab3:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            output_bytes = build_output_excel(matched_df, unmatched_df, summary_df)
            st.download_button(
                label="Download results (.xlsx)",
                data=output_bytes,
                file_name="llt_extractor_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(str(e))
else:
    st.info("Upload both files and click **Extract LLTs**.")
