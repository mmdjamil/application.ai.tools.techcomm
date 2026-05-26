"""
app.py
Streamlit web application — Inclusive Language & Broken Link Scanner
"""

import streamlit as st
import pandas as pd
from scanners import parse_file, scan_for_inclusive_language, check_all_links

# ──────────────────────────────────────────────
# Page Configuration
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Doc Scanner — Inclusive Language & Link Checker",
    page_icon="🔍",
    layout="wide",
)

st.title("📝 Document Scanner")
st.markdown("**Scan your technical documents for non-inclusive language and broken links.**")
st.markdown("Supported formats: `.docx` · `.pdf` · `.pptx` · `.xlsx`")

st.divider()

# ──────────────────────────────────────────────
# File Upload
# ──────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂 Upload your document",
    type=["docx", "pdf", "pptx", "xlsx"],
    help="Drag and drop or click to browse. Max 200MB."
)

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    filename = uploaded_file.name

    st.success(f"✅ Uploaded: **{filename}** ({len(file_bytes) / 1024:.1f} KB)")

    # ──────────────────────────────────────────
    # Scan Options
    # ──────────────────────────────────────────
    st.subheader("⚙️ Scan Options")
    col1, col2 = st.columns(2)
    with col1:
        run_inclusive = st.checkbox("🔤 Non-Inclusive Language Scan", value=True)
    with col2:
        run_links = st.checkbox("🔗 Broken Link Check", value=True)

    # ──────────────────────────────────────────
    # Run Scan
    # ──────────────────────────────────────────
    if st.button("🚀 Start Scan", type="primary", use_container_width=True):

        # Step 1: Parse the document
        with st.spinner("📄 Parsing document..."):
            try:
                parsed_lines = parse_file(file_bytes, filename)
                st.info(f"📄 Parsed **{len(parsed_lines)}** lines from the document.")
            except ValueError as e:
                st.error(str(e))
                st.stop()
            except Exception as e:
                st.error(f"❌ Error parsing file: {str(e)}")
                st.stop()

        # ──────────────────────────────────────
        # Non-Inclusive Language Results
        # ──────────────────────────────────────
        if run_inclusive:
            st.divider()
            st.subheader("🔤 Non-Inclusive Language Scan Results")

            with st.spinner("Scanning for non-inclusive terms..."):
                inclusive_results = scan_for_inclusive_language(parsed_lines)

            # Summary metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("🚨 Non-Inclusive Words Found", inclusive_results["total_non_inclusive_count"])
            m2.metric("📝 Total Word Count", f"{inclusive_results['total_word_count']:,}")
            m3.metric(
                "✅ Compliance Rate",
                f"{max(0, 100 - (inclusive_results['total_non_inclusive_count'] / max(inclusive_results['total_word_count'], 1) * 100)):.1f}%"
            )

            if inclusive_results["findings"]:
                df_inclusive = pd.DataFrame(inclusive_results["findings"])
                df_inclusive.columns = ["Page/Section", "Line", "Found Word", "Suggested Replacement", "Context"]
                df_inclusive.index = range(1, len(df_inclusive) + 1)
                st.dataframe(df_inclusive, use_container_width=True)

                csv_inclusive = df_inclusive.to_csv(index=False)
                st.download_button(
                    "📥 Download Inclusive Scan Results (CSV)",
                    csv_inclusive,
                    file_name="inclusive_scan_results.csv",
                    mime="text/csv"
                )
            else:
                st.success("🎉 No non-inclusive language found! Your document is clean.")

        # ──────────────────────────────────────
        # Broken Link Results
        # ──────────────────────────────────────
        if run_links:
            st.divider()
            st.subheader("🔗 Broken Link Check Results")

            progress_bar = st.progress(0, text="Checking links...")

            def update_progress(completed, total):
                progress_bar.progress(
                    completed / total,
                    text=f"Checking links... ({completed}/{total})"
                )

            link_results = check_all_links(parsed_lines, progress_callback=update_progress)
            progress_bar.empty()

            # Summary metrics
            l1, l2, l3 = st.columns(3)
            l1.metric("🔗 Total Links Checked", link_results["total_links_checked"])
            l2.metric("❌ Broken Links", link_results["broken_link_count"])
            l3.metric(
                "✅ Healthy Links",
                link_results["total_links_checked"] - link_results["broken_link_count"]
            )

            if link_results["total_links_checked"] > 0:
                if link_results["broken_links"]:
                    st.error(f"⚠️ {link_results['broken_link_count']} broken link(s) found!")
                    df_broken = pd.DataFrame(link_results["broken_links"])
                    df_broken = df_broken[["page", "line", "url", "status_code"]]
                    df_broken.columns = ["Page/Section", "Line", "URL", "Status"]
                    df_broken.index = range(1, len(df_broken) + 1)
                    st.dataframe(df_broken, use_container_width=True)
                else:
                    st.success("🎉 All links are healthy!")

                with st.expander("📋 View All Links Checked"):
                    df_all = pd.DataFrame(link_results["all_links"])
                    df_all = df_all[["page", "line", "url", "status_code", "is_broken"]]
                    df_all.columns = ["Page/Section", "Line", "URL", "Status", "Broken?"]
                    df_all["Broken?"] = df_all["Broken?"].map({True: "❌ Yes", False: "✅ No"})
                    df_all.index = range(1, len(df_all) + 1)
                    st.dataframe(df_all, use_container_width=True)

                df_download = pd.DataFrame(link_results["all_links"])
                csv_links = df_download.to_csv(index=False)
                st.download_button(
                    "📥 Download Link Check Results (CSV)",
                    csv_links,
                    file_name="link_check_results.csv",
                    mime="text/csv"
                )
            else:
                st.info("ℹ️ No links found in the document.")

# ──────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────
st.divider()
st.markdown("""
<div style='text-align: center; color: #888; font-size: 0.85em;'>
    📝 Doc Scanner v1.0 · Built for Technical Writers · Inclusive Language & Broken Link Checker
</div>
""", unsafe_allow_html=True)