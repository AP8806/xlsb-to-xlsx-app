import io
import zipfile
import streamlit as st
from converter import convert_uploaded_files

# Page Configuration
st.set_page_config(
    page_title="XLSB to XLSX Pro Converter",
    page_icon="⚡",
    layout="centered"
)

# Custom CSS for Modern UI
st.markdown("""
<style>
    .hero-box {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.15);
    }
    .hero-box h1 {
        color: white !important;
        margin-bottom: 8px;
        font-size: 2.2rem;
    }
    .hero-box p {
        color: #e0e8f5;
        font-size: 1.05rem;
        margin: 0;
    }
    .mode-card {
        background-color: #f8fafc;
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        text-align: center;
        margin-bottom: 10px;
    }
    div.stButton > button {
        border-radius: 10px;
        font-weight: 600;
        padding: 0.6rem 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Hero Banner
st.markdown("""
<div class="hero-box">
    <h1>⚡ XLSB to XLSX Converter</h1>
    <p>Convert your Excel Binary (.xlsb) files into Standard Excel (.xlsx) format in one click</p>
</div>
""", unsafe_allow_html=True)

# Session State to track selected mode
if "selected_mode" not in st.session_state:
    st.session_state.selected_mode = None

st.markdown("### 👇 Step 1: Select Conversion Mode")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="mode-card">
        <h3>📄 Single File</h3>
        <p style="color:gray; font-size:0.9rem;">Upload a single .xlsb file and download the converted .xlsx file instantly</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📄 Select Single File", use_container_width=True, 
                 type="primary" if st.session_state.selected_mode == "single" else "secondary"):
        st.session_state.selected_mode = "single"
        st.rerun()

with col2:
    st.markdown("""
    <div class="mode-card">
        <h3>📚 Multiple Files</h3>
        <p style="color:gray; font-size:0.9rem;">Batch convert multiple .xlsb files at once and download as a ZIP archive</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📚 Select Multiple Files", use_container_width=True, 
                 type="primary" if st.session_state.selected_mode == "multiple" else "secondary"):
        st.session_state.selected_mode = "multiple"
        st.rerun()

st.divider()

# =========================================================
# OPTION 1: SINGLE FILE UPLOAD & CONVERT
# =========================================================
if st.session_state.selected_mode == "single":
    st.markdown("### 📤 Step 2: Upload a Single `.xlsb` File")
    single_file = st.file_uploader(
        "Drag and drop your .xlsb file here or click Browse files",
        type=["xlsb"],
        accept_multiple_files=False,
        key="single_uploader"
    )

    if single_file is not None:
        file_kb = round(single_file.size / 1024, 2)
        st.info(f"📎 **Selected File:** `{single_file.name}` ({file_kb} KB)")

        if st.button("🚀 Convert to XLSX", type="primary", use_container_width=True):
            with st.spinner("⏳ Converting your file, please wait..."):
                converted, failed = convert_uploaded_files([single_file])

            if converted:
                out_name, out_bytes = converted[0]
                st.balloons()
                st.success(f"✅ `{out_name}` has been converted successfully!")
                st.download_button(
                    label=f"📥 Download {out_name}",
                    data=out_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            if failed:
                st.error(f"❌ Conversion Failed ({failed[0][0]}): {failed[0][1]}")

# =========================================================
# OPTION 2: MULTIPLE FILES UPLOAD & CONVERT
# =========================================================
elif st.session_state.selected_mode == "multiple":
    st.markdown("### 📤 Step 2: Upload Multiple `.xlsb` Files")
    multi_files = st.file_uploader(
        "Select or drag and drop one or more .xlsb files",
        type=["xlsb"],
        accept_multiple_files=True,
        key="multi_uploader"
    )

    if multi_files:
        st.info(f"📂 Total **{len(multi_files)}** file(s) ready for conversion.")

        if st.button(f"⚡ Convert All ({len(multi_files)} Files)", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_box = st.empty()

            def update_ui(curr, tot, fname):
                progress_bar.progress(curr / tot)
                status_box.markdown(f"🔄 **Converting [{curr}/{tot}]:** `{fname}` ...")

            converted, failed = convert_uploaded_files(multi_files, progress_callback=update_ui)
            status_box.empty()

            # Summary Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("📂 Total Files", len(multi_files))
            m2.metric("✅ Converted", len(converted))
            m3.metric("❌ Failed", len(failed))

            if converted:
                st.balloons()
                st.success("🎉 All files converted successfully! Click below to download.")

                # Create ZIP Archive
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, fbytes in converted:
                        zf.writestr(fname, fbytes)

                st.download_button(
                    label="📦 Download All Files (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Converted_XLSX_Files.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

                # Individual File Downloads
                with st.expander("📄 Download Individual Files"):
                    for idx, (fname, fbytes) in enumerate(converted):
                        st.download_button(
                            label=f"📥 Download {fname}",
                            data=fbytes,
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{idx}"
                        )

            if failed:
                st.warning("⚠️ Some files could not be converted:")
                for fname, err in failed:
                    st.error(f"**{fname}**: {err}")

else:
    st.info("👆 Please click on **Select Single File** or **Select Multiple Files** above to get started.")
