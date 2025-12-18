import os
import smtplib
import ssl
import unicodedata
from email.message import EmailMessage
from dotenv import load_dotenv

import streamlit as st
import pandas as pd
import plotly.express as px


st.set_page_config(page_title="Moodle Engagement Dashboard", layout="wide")


st.title("📊 Moodle Student Engagement Dashboard (from Log CSV)")

# load environment variables (optional)
load_dotenv()


# -------------------- Upload --------------------
file = st.file_uploader("Upload Moodle log CSV", type=["csv"])
if not file:
   st.info("Upload a CSV exported from Moodle logs to start.")
   st.stop()


df = pd.read_csv(file)


# -------------------- Normalize column names (fix hidden spaces / weird chars) --------------------
df.columns = (
   df.columns.astype(str)
   .str.replace("\ufeff", "", regex=False)
   .str.replace("\r", "", regex=False)
   .str.strip()
)


# -------------------- Validate columns --------------------
required_cols = ["Time", "User full name", "Event context", "Component", "Event name", "Origin"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
   st.error(f"Your CSV is missing columns: {missing}\n\nColumns found:\n{', '.join(df.columns)}")
   st.stop()


# -------------------- Parse Time --------------------
df["Time"] = pd.to_datetime(df["Time"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["Time"]).copy()
df["Date"] = df["Time"].dt.date


# -------------------- Sidebar filters --------------------
st.sidebar.header("Filters")


course_options = ["(All)"] + sorted(df["Event context"].dropna().unique().tolist())
course = st.sidebar.selectbox("Course / Context", course_options)


origin_options = ["(All)"] + sorted(df["Origin"].fillna("(blank)").unique().tolist())
default_origin_idx = origin_options.index("web") if "web" in origin_options else 0
origin = st.sidebar.selectbox("Origin", origin_options, index=default_origin_idx)


event_options = sorted(df["Event name"].dropna().unique().tolist())


default_events = [
   "Course viewed",
   "Section viewed",
   "Page viewed",
   "Resource viewed",
   "URL viewed",
   "File viewed",
   "Assignment viewed",
   "Assignment submitted",
   "Quiz attempted",
   "Quiz submission submitted",
   "Forum discussion viewed",
   "Post created",
   "Course module viewed",
]


default_selected = [e for e in default_events if e in event_options]
if not default_selected:
   default_selected = event_options[: min(8, len(event_options))]


selected_events = st.sidebar.multiselect(
   "Event names counted as engagement",
   options=event_options,
   default=default_selected
)


min_dt, max_dt = df["Time"].min(), df["Time"].max()
date_range = st.sidebar.date_input(
   "Date range",
   value=(min_dt.date(), max_dt.date()),
   min_value=min_dt.date(),
   max_value=max_dt.date()
)


lookback_days = st.sidebar.slider("Active days window (days)", 3, 30, 7)
risk_inactive_days = st.sidebar.slider("At-risk if inactive more than (days)", 7, 60, 14)


search_name = st.sidebar.text_input("Search student name (optional)")


# -------------------- Apply filters --------------------
f = df.copy()


if course != "(All)":
   f = f[f["Event context"] == course]


if origin != "(All)":
   f = f[f["Origin"].fillna("(blank)") == origin]


start_date, end_date = date_range
f = f[(f["Date"] >= start_date) & (f["Date"] <= end_date)]


if selected_events:
   f = f[f["Event name"].isin(selected_events)]


# drop cli if any remains
f = f[f["Origin"].fillna("") != "cli"]


if f.empty:
   st.warning("No data after filters. Try selecting more events or a wider date range.")
   st.stop()


# -------------------- Compute summary metrics --------------------
today = pd.Timestamp.now().normalize()
cutoff = (today - pd.Timedelta(days=lookback_days)).date()


content_view_events = {"Section viewed", "Page viewed", "Resource viewed", "URL viewed", "File viewed", "Course module viewed"}
submission_events = {"Assignment submitted", "Quiz attempted", "Quiz submission submitted", "Post created"}


summary = f.groupby("User full name").agg(
   last_access=("Time", "max"),
   total_events=("Event name", "count"),
   active_days=("Date", lambda x: pd.Series(x).loc[pd.Series(x) >= cutoff].nunique()),
   content_views=("Event name", lambda x: x.isin(content_view_events).sum()),
   course_views=("Event name", lambda x: (x == "Course viewed").sum()),
   submissions=("Event name", lambda x: x.isin(submission_events).sum()),
).reset_index()


summary["inactive_days"] = (today - summary["last_access"].dt.normalize()).dt.days


def status_row(row):
   if row["inactive_days"] > risk_inactive_days or row["active_days"] == 0:
       return "⚠️ At Risk"
   if row["active_days"] <= 2:
       return "🟡 Warning"
   return "✅ Active"


summary["status"] = summary.apply(status_row, axis=1)


# optional name search (after aggregation)
if search_name.strip():
   summary = summary[summary["User full name"].str.contains(search_name, case=False, na=False)]


# sort: at risk first, then by inactivity
order = {"⚠️ At Risk": 0, "🟡 Warning": 1, "✅ Active": 2}
summary["status_rank"] = summary["status"].map(order)
summary = summary.sort_values(["status_rank", "inactive_days"], ascending=[True, False]).drop(columns=["status_rank"])


# -------------------- KPI Row --------------------
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Students", int(summary["User full name"].nunique()))
c2.metric("Events (filtered)", int(len(f)))
c3.metric("At Risk", int((summary["status"] == "⚠️ At Risk").sum()))
c4.metric("Warning", int((summary["status"] == "🟡 Warning").sum()))
c5.metric("Active", int((summary["status"] == "✅ Active").sum()))


st.divider()


# -------------------- At-Risk Cards --------------------
st.subheader("🚨 Students needing attention (At Risk)")
at_risk = summary[summary["status"] == "⚠️ At Risk"].copy()


if at_risk.empty:
   st.success("No students are currently flagged as At Risk 🎉")
else:
   cols = st.columns(3)
   for i, (_, r) in enumerate(at_risk.iterrows()):
       with cols[i % 3]:
           st.error(
               f"**{r['User full name']}**\n\n"
               f"Inactive: **{int(r['inactive_days'])} days**\n\n"
               f"Active days (last {lookback_days}d): **{int(r['active_days'])}**\n\n"
               f"Content views: **{int(r['content_views'])}**\n\n"
               f"Submissions: **{int(r['submissions'])}**\n\n"
               f"Last access: **{r['last_access'].strftime('%Y-%m-%d %H:%M:%S')}**"
           )


st.divider()


# -------------------- E-mail notifier (FIXED SMTP) --------------------
st.header("📮 Notify coordinator about At-Risk students")

# coordinator e-mail + SMTP settings (fall back to env)
coord_mail = st.text_input("Coordinator e-mail"
                           # , value=os.getenv("COORD_EMAIL", "")
                           )
smtp_user = os.getenv("SMTP_USER")
smtp_pass = os.getenv("SMTP_PASS")
smtp_srv = os.getenv("SMTP_SERVER", "smtp.gmail.com")
smtp_port = int(os.getenv("SMTP_PORT", "587"))

if st.button("📩 Send alert email", type="primary"):
   if at_risk.empty:
      st.info("No at-risk students to notify.")
   else:
      if not coord_mail:
         st.error("Please enter a coordinator e-mail")
      else:
         body_lines = [
            f"Moodle early-alert – {pd.Timestamp.now():%Y-%m-%d %H:%M}",
            f"Students flagged as At Risk: {len(at_risk)}",
            "",
         ]
         for _, r in at_risk.iterrows():
            body_lines.append(f"{r['User full name']} — Inactive {int(r['inactive_days'])} days — Active days: {int(r['active_days'])}")
         body = "\n".join(body_lines)

         msg = EmailMessage()
         msg["Subject"] = f"Moodle early-alert – {len(at_risk)} students"
         msg["From"] = smtp_user or coord_mail
         msg["To"] = coord_mail
         msg.set_content(body)

         try:
            # sanitize SMTP credentials to ASCII (SMTP auth requires ASCII)
            def _to_ascii(s: str) -> str:
               if s is None:
                  return s
               # common smart quotes / dashes -> ascii equivalents
               s = s.replace(""", '"').replace(""", '"').replace("'", "'").replace("'", "'")
               s = s.replace("–", "-").replace("—", "-")
               s = unicodedata.normalize('NFKD', s)
               return s.encode('ascii', 'ignore').decode('ascii')

            login_user = _to_ascii(smtp_user) if smtp_user else None
            login_pass = _to_ascii(smtp_pass) if smtp_pass else None
            if (smtp_user and login_user != smtp_user) or (smtp_pass and login_pass != smtp_pass):
               st.warning("Non-ASCII characters were removed from SMTP credentials before sending.")

            # Create SSL context
            context = ssl.create_default_context()
            
            # Use SMTP with STARTTLS for port 587 (most common for Gmail)
            if smtp_port == 587:
                with smtplib.SMTP(smtp_srv, smtp_port, timeout=10) as server:
                    server.ehlo()  # Identify ourselves
                    server.starttls(context=context)  # Secure the connection
                    server.ehlo()  # Re-identify ourselves over secure connection
                    if login_user and login_pass:
                        server.login(login_user, login_pass)
                    server.send_message(msg)
            # Use SMTP_SSL for port 465 (SSL from the start)
            elif smtp_port == 465:
                with smtplib.SMTP_SSL(smtp_srv, smtp_port, context=context, timeout=10) as server:
                    if login_user and login_pass:
                        server.login(login_user, login_pass)
                    server.send_message(msg)
            else:
                st.error(f"Unsupported SMTP port: {smtp_port}. Use 587 (STARTTLS) or 465 (SSL).")
                st.stop()
                
            st.success(f"✅ Mail sent to {coord_mail}")
            if "notif_log" not in st.session_state:
               st.session_state.notif_log = []
            st.session_state.notif_log.append(f"{pd.Timestamp.now():%H:%M} – mail sent ({len(at_risk)} students)")
            
         except smtplib.SMTPAuthenticationError:
            st.error("❌ Authentication failed. Check your SMTP username and password.")
            st.info("💡 For Gmail: Use an App Password (requires 2FA enabled)")
         except smtplib.SMTPException as e:
            st.error(f"❌ SMTP error: {str(e)}")
         except ssl.SSLError as e:
            st.error(f"❌ SSL error: {str(e)}")
            st.info("💡 Try running: /Applications/Python\\ 3.13/Install\\ Certificates.command")
         except Exception as e:
            st.error("❌ Send failed")
            st.exception(e)

if st.session_state.get("notif_log"):
   with st.expander("Notification log"):
      st.text("\n".join(st.session_state.notif_log))


# -------------------- Tabs --------------------
tab1, tab2, tab3 = st.tabs(["📋 Engagement Table", "📈 Visualisations", "🔍 Student Detail"])


with tab1:
   st.subheader("Engagement Summary (color-coded)")


   def style_status(row):
       s = row["status"]
       if s == "⚠️ At Risk":
           return ["background-color: #ffe5e5"] * len(row)  # light red
       if s == "🟡 Warning":
           return ["background-color: #fff6d6"] * len(row)  # light yellow
       return ["background-color: #e9f7ef"] * len(row)      # light green


   st.dataframe(summary.style.apply(style_status, axis=1), use_container_width=True)


   st.download_button(
       "⬇️ Download summary CSV",
       data=summary.to_csv(index=False).encode("utf-8"),
       file_name="engagement_summary.csv",
       mime="text/csv"
   )


with tab2:
   st.subheader("Risk overview (names + clear status)")


   fig = px.bar(
       summary.sort_values("inactive_days", ascending=False).head(30),
       x="User full name",
       y="inactive_days",
       color="status",
       hover_data=["last_access", "active_days", "total_events", "content_views", "submissions"],
       title="Top 30 students by inactive days"
   )
   fig.update_layout(xaxis_title="Student", yaxis_title="Inactive Days", xaxis_tickangle=-45)
   st.plotly_chart(fig, use_container_width=True)


   st.subheader("Events over time")
   daily = f.groupby("Date").size().reset_index(name="events")
   st.line_chart(daily, x="Date", y="events")


   st.subheader("Top event types")
   evt = f["Event name"].value_counts().head(15).reset_index()
   evt.columns = ["event_name", "count"]
   fig2 = px.bar(evt, x="event_name", y="count", title="Top 15 event types (filtered)")
   fig2.update_layout(xaxis_title="Event name", yaxis_title="Count", xaxis_tickangle=-30)
   st.plotly_chart(fig2, use_container_width=True)


with tab3:
   st.subheader("Student activity timeline")


   if summary.empty:
       st.info("No students to display.")
   else:
       student = st.selectbox("Select student", sorted(summary["User full name"].unique().tolist()))
       sf = f[f["User full name"] == student].sort_values("Time", ascending=False)


       student_row = summary[summary["User full name"] == student].iloc[0]
       k1, k2, k3, k4 = st.columns(4)
       k1.metric("Status", student_row["status"])
       k2.metric("Inactive days", int(student_row["inactive_days"]))
       k3.metric(f"Active days (last {lookback_days}d)", int(student_row["active_days"]))
       k4.metric("Total events", int(student_row["total_events"]))


       st.write("Recent events (latest first)")
       show_cols = ["Time", "Event name", "Event context", "Component", "Origin"]
       if "IP address" in sf.columns:
           show_cols.append("IP address")
       st.dataframe(sf[show_cols].head(300), use_container_width=True)


       st.write("Event breakdown (this student)")
       breakdown = sf["Event name"].value_counts().reset_index()
       breakdown.columns = ["event_name", "count"]
       fig3 = px.bar(breakdown.head(20), x="event_name", y="count", title=f"Top events for {student}")
       fig3.update_layout(xaxis_tickangle=-30)
       st.plotly_chart(fig3, use_container_width=True)