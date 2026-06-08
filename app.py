"""
app.py
Streamlit web application — Inclusive Language & Broken Link Scanner
"""

import streamlit as st
import pandas as pd
from scanners import parse_file, scan_for_inclusive_language, check_all_links, rewrite_file

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

    # Reset cached scan results if a new file was uploaded
    if (
        "scan_data" in st.session_state
        and st.session_state["scan_data"].get("filename") != filename
    ):
        del st.session_state["scan_data"]

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
    # Link Authorization (optional)
    # ──────────────────────────────────────────
    with st.expander("🔐 Link authorization (optional)"):
        bearer_token = st.text_input(
            "🔑 ******",
            type="password",
            placeholder="ghp_xxxxxxxxxxxx",
            help="Sent as 'Authorization: ******'.",
        )
        auth_user = st.text_input("👤 Basic auth username", placeholder="username")
        auth_pass = st.text_input("🔒 Basic auth password", type="password", placeholder="password")
        custom_headers_raw = st.text_area(
            "📋 Custom headers (one per line, Key: Value)",
            placeholder="X-API-Key: abc123\nX-Custom-Header: value",
            help="Each line must be in 'Key: Value' format.",
        )
        allowed_hosts_raw = st.text_input(
            "🌐 Allowed hosts (comma-separated)",
            placeholder="api.github.com, raw.githubusercontent.com",
            help="If set, credentials are ONLY sent to these hosts.",
        )

    # Build auth_config from inputs
    auth_config: dict = {}
    if bearer_token.strip():
        auth_config["bearer_token"] = bearer_token.strip()
    if auth_user.strip() and auth_pass.strip():
        auth_config["basic_auth"] = (auth_user.strip(), auth_pass.strip())
    if custom_headers_raw.strip():
        custom_headers = {}
        for raw_line in custom_headers_raw.splitlines():
            if ":" in raw_line:
                key, _, value = raw_line.partition(":")
                key = key.strip()
                value = value.strip()
                if key:
                    custom_headers[key] = value
        if custom_headers:
            auth_config["headers"] = custom_headers
    if allowed_hosts_raw.strip():
        auth_config["allowed_hosts"] = [
            h.strip() for h in allowed_hosts_raw.split(",") if h.strip()
        ]
    auth_config = auth_config if auth_config else None

    # ──────────────────────────────────────────
    # Run Scan (only re-runs the scan when the button is clicked)
    # ──────────────────────────────────────────
    if st.button("🚀 Start Scan", type="primary", use_container_width=True):

        # Step 1: Parse the document
        with st.spinner("📄 Parsing document..."):
            try:
                parsed_lines = parse_file(file_bytes, filename)
            except ValueError as e:
                st.error(str(e))
                st.stop()
            except Exception as e:
                st.error(f"❌ Error parsing file: {str(e)}")
                st.stop()

        # Step 2: Scan for non-inclusive language
        inclusive_results = None
        if run_inclusive:
            with st.spinner("Scanning for non-inclusive terms..."):
                inclusive_results = scan_for_inclusive_language(parsed_lines)

        # Step 3: Check links
        link_results = None
        if run_links:
            progress_bar = st.progress(0, text="Checking links...")

            def update_progress(completed, total):
                progress_bar.progress(
                    completed / total,
                    text=f"Checking links... ({completed}/{total})"
                )

            link_results = check_all_links(
                parsed_lines,
                progress_callback=update_progress,
                auth_config=auth_config,
            )
            progress_bar.empty()

        # Persist for reruns triggered by review widgets
        st.session_state["scan_data"] = {
            "filename": filename,
            "parsed_lines": parsed_lines,
            "inclusive_results": inclusive_results,
            "link_results": link_results,
            "run_inclusive": run_inclusive,
            "run_links": run_links,
        }

    # ──────────────────────────────────────────
    # Render results from session state
    # (so checkboxes / generate-copy button survive reruns)
    # ──────────────────────────────────────────
    if (
        "scan_data" in st.session_state
        and st.session_state["scan_data"].get("filename") == filename
    ):
        scan_data = st.session_state["scan_data"]
        parsed_lines = scan_data["parsed_lines"]
        inclusive_results = scan_data["inclusive_results"]
        link_results = scan_data["link_results"]

        st.info(f"📄 Parsed **{len(parsed_lines)}** lines from the document.")

        # ──────────────────────────────────────
        # Non-Inclusive Language Results
        # ──────────────────────────────────────
        if scan_data["run_inclusive"] and inclusive_results is not None:
            st.divider()
            st.subheader("🔤 Non-Inclusive Language Scan Results")

            # Summary metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("🚨 Non-Inclusive Words Found", inclusive_results["total_non_inclusive_count"])
            m2.metric("📝 Total Word Count", f"{inclusive_results['total_word_count']:,}")
            m3.metric(
                "✅ Compliance Rate",
                f"{max(0, 100 - (inclusive_results['total_non_inclusive_count'] / max(inclusive_results['total_word_count'], 1) * 100)):.1f}%"
            )

            if inclusive_results["findings"]:
                # Build the review dataframe
                df_inclusive = pd.DataFrame(inclusive_results["findings"])
                df_inclusive["apply"] = True
                df_inclusive = df_inclusive.rename(columns={
                    "apply": "Apply?",
                    "page": "Page/Section",
                    "line": "Line",
                    "found_word": "Found Word",
                    "suggested_replacement": "Suggested Replacement",
                    "original_sentence": "Original Sentence",
                    "suggested_sentence": "Suggested Sentence",
                    "context": "Context",
                    "term": "_term",  # internal, hidden via column_order
                })
                df_inclusive.index = range(1, len(df_inclusive) + 1)

                st.markdown("### ✏️ Review suggested replacements")
                st.caption(
                    "Untick **Apply?** for any suggestion you don't want to keep, "
                    "then click **Generate corrected copy** to download a rewritten "
                    "version of your document."
                )

                edited_df = st.data_editor(
                    df_inclusive,
                    use_container_width=True,
                    column_order=[
                        "Apply?",
                        "Page/Section",
                        "Line",
                        "Found Word",
                        "Suggested Replacement",
                        "Original Sentence",
                        "Suggested Sentence",
                        "Context",
                    ],
                    column_config={
                        "Apply?": st.column_config.CheckboxColumn(
                            "Apply?",
                            help="Accept this replacement when generating the corrected copy.",
                            default=True,
                        ),
                        "Original Sentence": st.column_config.TextColumn(
                            "Original Sentence", width="medium"
                        ),
                        "Suggested Sentence": st.column_config.TextColumn(
                            "Suggested Sentence", width="medium"
                        ),
                    },
                    disabled=[
                        "Page/Section",
                        "Line",
                        "Found Word",
                        "Suggested Replacement",
                        "Original Sentence",
                        "Suggested Sentence",
                        "Context",
                    ],
                    key="inclusive_review_editor",
                )

                # Downloads / actions
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    csv_inclusive = (
                        edited_df.drop(columns=["_term"], errors="ignore")
                        .to_csv(index=False)
                    )
                    st.download_button(
                        "📥 Download scan results (CSV)",
                        csv_inclusive,
                        file_name="inclusive_scan_results.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

                with col_dl2:
                    generate_clicked = st.button(
                        "✅ Generate corrected copy",
                        type="primary",
                        use_container_width=True,
                        help="Apply the ticked replacements and download a rewritten copy of your document text.",
                    )

                if generate_clicked:
                    accepted_rows = edited_df[edited_df["Apply?"] == True]

                    if accepted_rows.empty:
                        st.warning(
                            "No suggestions selected. Tick **Apply?** for at least one row "
                            "to generate a corrected copy."
                        )
                    else:
                        # Group accepted (term, replacement) pairs by (page, line)
                        accepted_by_key: dict = {}
                        for _, row in accepted_rows.iterrows():
                            key = (row["Page/Section"], row["Line"])
                            term = row["_term"]
                            replacement = row["Suggested Replacement"]
                            accepted_by_key.setdefault(key, []).append((term, replacement))

                        accepted_findings = [
                            {
                                "page": row["Page/Section"],
                                "line": row["Line"],
                                "found_word": row["Found Word"],
                                "term": row["_term"],
                                "suggested_replacement": row["Suggested Replacement"],
                            }
                            for _, row in accepted_rows.iterrows()
                        ]

                        applied_count = sum(len(v) for v in accepted_by_key.values())
                        output_bytes, output_filename, mime = rewrite_file(
                            file_bytes,
                            filename,
                            accepted_by_key,
                            accepted_findings=accepted_findings,
                        )

                        st.success(
                            f"✅ Applied **{applied_count}** replacement(s) across "
                            f"**{len(accepted_by_key)}** line(s). Changed words are "
                            f"highlighted and (where supported) annotated in the "
                            f"corrected copy."
                        )

                        # Show a text preview only for the PDF plain-text fallback;
                        # previewing raw binary bytes is not useful.
                        if mime == "text/plain":
                            with st.expander("👀 Preview corrected document", expanded=True):
                                st.text_area(
                                    "Corrected document (text)",
                                    output_bytes.decode("utf-8", errors="replace"),
                                    height=400,
                                    label_visibility="collapsed",
                                )
                            st.caption(
                                "ℹ️ This PDF could not be annotated, so the corrected "
                                "copy was exported as plain text instead."
                            )
                        elif mime == "application/pdf":
                            st.caption(
                                "ℹ️ Open the corrected PDF in your reader of choice "
                                "to see highlights and sticky notes for each "
                                "suggested replacement."
                            )

                        st.download_button(
                            f"⬇️ Download corrected copy ({output_filename.rsplit('.', 1)[-1]})",
                            output_bytes,
                            file_name=output_filename,
                            mime=mime,
                            use_container_width=True,
                        )
            else:
                st.success("🎉 No non-inclusive language found! Your document is clean.")

        # ──────────────────────────────────────
        # Broken Link Results
        # ──────────────────────────────────────
        if scan_data["run_links"] and link_results is not None:
            st.divider()
            st.subheader("🔗 Broken Link Check Results")

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
