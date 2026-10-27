                    cell.font = Font(color="FFFFFF", bold=True)
                for cell in ws[2]:
                    cell.fill = PatternFill("solid", fgColor="DCECE8")
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                ws.freeze_panes = "G3"
            else:
                recheck["file"].to_excel(writer, sheet_name="REACH_RECHECK", index=False)
        for ws in writer.book.worksheets:
            if ws.title != "REACH_RECHECK":
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for cell in ws[1]:
                    cell.fill = PatternFill("solid", fgColor="146451")
                    cell.font = Font(color="FFFFFF", bold=True)
                for column in ws.columns:
                    letter = column[0].column_letter
                    ws.column_dimensions[letter].width = min(35, max(13, len(str(column[0].value or "")) + 3))
    return output.getvalue()


def count_comments(df):
    if "COMMENT" not in df:
        return 0
    comments = df["COMMENT"].fillna("").astype(str).str.strip()
    return int(comments.ne("").sum())


heading, photo = st.columns([3.2, 1], vertical_alignment="center")
with heading:
    st.markdown('<div class="hero-title">Malaria SSR1.6 — Data Validation</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Upload your trusted Python rules and malaria Excel records, run checks, and download results.</div>', unsafe_allow_html=True)
with photo:
    # A local copy works offline; otherwise Streamlit loads the Wikimedia Commons photo.
    try:
        st.image(str(LOCAL_PHOTO) if LOCAL_PHOTO.exists() else PHOTO_URL,
                 caption="Anopheles minimus • Photo: CDC / James Gathany (public domain)",
                 use_container_width=True)
    except Exception:
        st.caption("Mosquito photo not available offline. Add anopheles_mosquito.jpg beside this app.")

st.divider()
if TEMPLATE_FILE.is_file():
    st.download_button("Download blank Excel input template", data=TEMPLATE_FILE.read_bytes(),
                       file_name=TEMPLATE_FILE.name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.subheader("1. Upload malaria validation rules")
st.caption("Security: Python files run on the app server. Only upload rule scripts you trust. Administrator access is required.")
if not admin_authorized():
    st.stop()
rule_upload = st.file_uploader("Upload malaria rules (.py)", type=["py"], key="rule_file")
if rule_upload is not None:
    st.caption(f"Selected rules: {rule_upload.name}")

st.subheader("2. Upload Excel records")

mode = st.radio("Run checks", ["Positive + Aggregate", "Positive only", "Aggregate only"],
                horizontal=True, index=0)
need_pos = mode != "Aggregate only"
need_agg = mode != "Positive only"

left, right = st.columns(2)
with left:
    st.subheader("Positive cases")
    pos_upload = st.file_uploader("Upload positive records (.xlsx)", type=["xlsx"],
                                  key="positive", disabled=not need_pos)
    pos_sheet = select_sheet(pos_upload, "positive_sheet") if need_pos else None
with right:
    st.subheader("Aggregate / total records")
    agg_upload = st.file_uploader("Upload aggregate records (.xlsx)", type=["xlsx"],
                                  key="aggregate", disabled=not need_agg)
    agg_sheet = select_sheet(agg_upload, "aggregate_sheet") if need_agg else None

can_run = (rule_upload is not None and (not need_pos or pos_upload is not None) and (not need_agg or agg_upload is not None))
run = st.button("Run malaria validation", type="primary", disabled=not can_run,
                use_container_width=True)

if run:
    try:
        rules = load_uploaded_rules(rule_upload.getvalue())
        positive_source = read_sheet(pos_upload, pos_sheet) if need_pos else None
        aggregate_source = read_sheet(agg_upload, agg_sheet) if need_agg else None
        with st.spinner("Checking records against SSR1.6 rules..."):
            checked_pos = rules.malaria_positive(positive_source) if need_pos else None
            checked_agg = rules.malaria_aggregate(aggregate_source) if need_agg else None
            recheck = None
            if checked_pos is not None and checked_agg is not None:
                recheck = rules.malaria_recheck(
                    checked_pos, checked_agg,
                    original_input_columns=set(positive_source.columns),
                )
            result_file = to_excel(checked_pos, checked_agg, recheck)
        st.session_state["result"] = (checked_pos, checked_agg, recheck, result_file)
    except Exception as exc:
        st.session_state.pop("result", None)
        st.error(f"Validation failed: {exc}")

if "result" in st.session_state:
    checked_pos, checked_agg, recheck, result_file = st.session_state["result"]
    st.success("Validation finished. Review findings and download the checked workbook.")
    metrics = st.columns(4)
    metrics[0].metric("Positive records", len(checked_pos) if checked_pos is not None else "—")
    metrics[1].metric("Positive rows with comments", count_comments(checked_pos) if checked_pos is not None else "—")
    metrics[2].metric("Aggregate records", len(checked_agg) if checked_agg is not None else "—")
    metrics[3].metric("Aggregate rows with comments", count_comments(checked_agg) if checked_agg is not None else "—")

    st.download_button("Download all validation results (.xlsx)", data=result_file,
                       file_name="Malaria_SSR1_6_Validation_Results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")

    tabs = []
    if checked_pos is not None:
        tabs.append(("Positive results", checked_pos))
    if checked_agg is not None:
        tabs.append(("Aggregate results", checked_agg))
    if recheck is not None:
        tabs.append(("REACH_RECHECK preview", recheck.get("display", pd.DataFrame())))
    tab_widgets = st.tabs([x[0] for x in tabs])
    for tab, (_, df) in zip(tab_widgets, tabs):
        with tab:
            if "COMMENT" in df.columns:
                only_findings = st.checkbox("Show rows with comments only", key=f"filter_{tab_widgets.index(tab)}")
                if only_findings:
                    df = df[df["COMMENT"].fillna("").astype(str).str.strip().ne("")]
            st.dataframe(df, use_container_width=True, hide_index=True)
    if recheck is not None:
