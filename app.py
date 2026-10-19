import os, tempfile
import pandas as pd
import streamlit as st
from make_report import parse_pdf, build_workbook, hourly_table

st.set_page_config(page_title="Machine Time Report", page_icon="📊", layout="wide")
st.title("Machine Time Report")
st.caption("Upload the daily production report PDFs, preview the hourly sheet, then download it.")

def hms(sec):
    sec = int(sec); return f"{sec // 3600}:{sec % 3600 // 60:02d}:{sec % 60:02d}"

files = st.file_uploader("Report PDFs (select several at once)", type="pdf", accept_multiple_files=True)
c1, c2 = st.columns(2)
threshold = c1.number_input("Threshold (minutes)", min_value=1, max_value=60, value=10)
expect = c2.text_input("Machines that must appear (comma separated)", placeholder="B650 T, B1000 T")

if st.button("Generate", type="primary", disabled=not files):
    data, date = {}, None
    with tempfile.TemporaryDirectory() as tmp:
        for f in files:
            p = os.path.join(tmp, f.name)
            with open(p, "wb") as fh: fh.write(f.getbuffer())
            mach, d, rows = parse_pdf(p)
            if not mach:
                st.warning(f"Skipped {f.name}: machine name not found"); continue
            date = date or d
            seen = data.setdefault(mach, [])
            seen.extend(r for r in rows if r not in seen)
        if data:
            exp = [m.strip() for m in expect.split(",") if m.strip()]
            out = os.path.join(tmp, "report.xlsx")
            build_workbook(data, date, out, threshold, exp)
            st.session_state["res"] = dict(
                xlsx=open(out, "rb").read(), date=date, threshold=threshold,
                rows=hourly_table(data, sorted(set(data) | set(exp))),
                missing=[m for m in exp if m not in data], n=len(data))
        else:
            st.session_state.pop("res", None)

res = st.session_state.get("res")
if res:
    limit = res["threshold"] * 60
    df = pd.DataFrame(res["rows"])
    st.success(f"{res['n']} machine(s) processed" + (f" for {res['date']}" if res["date"] else ""))
    if res["missing"]: st.info("No data for: " + ", ".join(res["missing"]))

    tab1, tab2 = st.tabs(["Hourly", "Summary"])
    with tab1:
        machines = sorted(df["machine"].unique())
        pick = st.multiselect("Machines", machines, default=machines)
        only = st.checkbox(f"Only hours with missing time above {res['threshold']} min")
        v = df[df["machine"].isin(pick)]
        if only: v = v[v["missing"] > limit]
        show = pd.DataFrame({"Time Duration": v["hour"], "Shift": v["shift"], "Machine Name": v["machine"],
                             "On Time": v["on"].map(hms), "Idle Time": v["idle"].map(hms),
                             "Off Time": v["off"].map(hms), "Missing Time": v["missing"].map(hms),
                             "Status": v["missing"].map(lambda s: "Above threshold" if s > limit else "Within threshold")})
        def style(col):
            return ["background-color:#FFC7CE;color:#9C0006" if s == "Above threshold"
                    else "background-color:#C6EFCE;color:#006100" for s in show["Status"]]
        st.dataframe(show.style.apply(style, subset=["Missing Time", "Status"]),
                     width="stretch", hide_index=True, height=480)
    with tab2:
        g = df.groupby("machine")[["on", "idle", "off", "missing"]].sum()
        summ = g.map(hms).rename(columns={"on": "On Time", "idle": "Idle Time", "off": "Off Time", "missing": "Missing Time"})
        summ["Hours above threshold"] = df[df["missing"] > limit].groupby("machine").size().reindex(g.index).fillna(0).astype(int)
        summ.index.name = "Machine Name"
        st.dataframe(summ, width="stretch")

    st.subheader("Download")
    name = f"Hourly_Machine_Time_Report_{(res['date'] or 'report').replace(' ', '')}"
    d1, d2 = st.columns(2)
    d1.download_button("Download Excel (.xlsx)", res["xlsx"], file_name=name + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    csv = show.to_csv(index=False).encode("utf-8")
    d2.download_button("Download CSV (current view)", csv, file_name=name + ".csv", mime="text/csv")
