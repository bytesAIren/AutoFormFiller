#!/usr/bin/env python3
"""
Quick MVP UI for Tender Form Filler.

Run:
    streamlit run app_streamlit.py
"""

import tempfile
from pathlib import Path

import streamlit as st

from tender_filler import (
    analyze_form_labels,
    fill_docx,
    fill_pdf,
    load_profile,
    validate_profile,
)


st.set_page_config(page_title="Tender Form Filler MVP", page_icon="🧪", layout="centered")
st.title("🧪 Tender Form Filler — MVP Test UI")
st.caption("Upload one form + one profile, run analysis/fill, and download output.")

form_file = st.file_uploader("1) Upload form (.docx or .pdf)", type=["docx", "pdf"])
profile_file = st.file_uploader("2) Upload profile (.json or .csv)", type=["json", "csv"])
run_analysis = st.checkbox("Run label coverage analysis before fill", value=True)

if form_file and profile_file:
    form_suffix = Path(form_file.name).suffix.lower()
    profile_suffix = Path(profile_file.name).suffix.lower()

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        form_path = tmpdir / form_file.name
        profile_path = tmpdir / profile_file.name
        out_path = tmpdir / f"{Path(form_file.name).stem}_COMPILATO{form_suffix}"

        form_path.write_bytes(form_file.getbuffer())
        profile_path.write_bytes(profile_file.getbuffer())

        try:
            profile = load_profile(profile_path)
        except Exception as exc:
            st.error(f"Profile parsing failed: {exc}")
            st.stop()

        missing = validate_profile(profile)
        if missing:
            st.warning("Missing recommended MVP keys:")
            for key in missing:
                st.write(f"- `{key}`")
        else:
            st.success("Profile has all recommended MVP keys.")

        if run_analysis:
            try:
                report = analyze_form_labels(form_path, profile)
                if report["supported"]:
                    st.info(f"Coverage: {report['matched']} / {report['total']} labels matched")
                    if report["unmatched_examples"]:
                        st.write("Unmatched examples:")
                        for item in report["unmatched_examples"]:
                            st.write(f"- {item}")
            except Exception as exc:
                st.warning(f"Analysis skipped due to error: {exc}")

        if st.button("3) Fill form now", type="primary"):
            try:
                if form_suffix == ".docx":
                    ok = fill_docx(str(form_path), profile, str(out_path))
                elif form_suffix == ".pdf":
                    ok = fill_pdf(str(form_path), profile, str(out_path))
                else:
                    st.error("Unsupported form format. Use .docx or .pdf")
                    st.stop()
            except Exception as exc:
                st.error(f"Form fill failed: {exc}")
                st.stop()

            if ok and out_path.exists():
                st.success("Form processed. Download output below.")
                st.download_button(
                    label="Download filled file",
                    data=out_path.read_bytes(),
                    file_name=out_path.name,
                    mime="application/octet-stream",
                )
            else:
                st.error("Processing ended with warnings. Check logs/terminal output.")
else:
    st.write("Upload both files to continue.")
