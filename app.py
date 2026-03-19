from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


STABLE_SORT_COLUMNS = ["תאריך_הצטרפות", "client_id"]


def find_dataset_path() -> Path:
    """Locate dataset in the app folder or nearby parent folders."""
    app_dir = Path(__file__).resolve().parent
    search_dirs = [app_dir, app_dir.parent, app_dir.parent.parent]
    names = [
        "clients_data impact.xlsx",
        "clients_data.csv",
        "clients_data.xlsx",
    ]

    for base_dir in search_dirs:
        for name in names:
            candidate = base_dir / name
            if candidate.exists():
                return candidate

    raise FileNotFoundError(
        "Dataset file was not found. Expected one of: "
        "clients_data impact.xlsx, clients_data.xlsx, clients_data.csv"
    )


def load_data() -> pd.DataFrame:
    """Load the assignment dataset from Excel or CSV."""
    path = find_dataset_path()

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    df["תאריך_הצטרפות"] = pd.to_datetime(df["תאריך_הצטרפות"], errors="coerce")
    df["תאריך_נטישה"] = pd.to_datetime(df["תאריך_נטישה"], errors="coerce")

    return df


def build_corrected_ids(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy where client_id is reassigned by signup date order."""
    corrected = df.copy()
    corrected = corrected.sort_values(STABLE_SORT_COLUMNS, kind="mergesort").reset_index(drop=True)
    corrected["client_id"] = [f"C{1000 + i}" for i in range(len(corrected))]
    return corrected


def export_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="clients_corrected")
    return output.getvalue()


def add_filters(df: pd.DataFrame) -> pd.DataFrame:
    cities = sorted(df["עיר"].dropna().unique().tolist())
    services = sorted(df["סוג_שירות"].dropna().unique().tolist())
    statuses = sorted(df["סטטוס"].dropna().unique().tolist())

    with st.popover("🔍 סינון נתונים", use_container_width=False):
        selected_cities = st.multiselect("עיר", cities, default=cities)
        selected_services = st.multiselect("סוג שירות", services, default=services)
        selected_statuses = st.multiselect("סטטוס", statuses, default=statuses)

    filtered = df[
        df["עיר"].isin(selected_cities)
        & df["סוג_שירות"].isin(selected_services)
        & df["סטטוס"].isin(selected_statuses)
    ].copy()

    return filtered


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    """Convert hex color to rgba string."""
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    return f"rgba({r},{g},{b},{alpha})"


def _process_flowchart(is_current: bool, highlight_node: int | None = None) -> go.Figure:
    """Build a detailed process-map Sankey diagram with tooltips and highlighting."""
    if is_current:
        labels = [
            "פנייה נכנסת",        # 0
            "⚠ קליטה ידנית CRM",  # 1
            "בדיקת פרטים",        # 2
            "⚠ השלמת מידע",       # 3
            "⚠ העתקה ל-Drive",    # 4
            "⚠ שיוך ליועץ",       # 5
            "⚠ קשר ראשוני 48h",   # 6
            "⚠ תזכורות ידניות",    # 7
            "פגישת ייעוץ",        # 8
            "⚠ הצעה ידנית",       # 9
            "פולואפ",             # 10
            "✓ לקוח משלם",        # 11
            "✗ סגירת פנייה",      # 12
        ]
        node_descriptions = [
            "פנייה נכנסת<br>לקוח פוטנציאלי פונה דרך טלפון, מייל או אתר.<br>נקודת כניסה – פוטנציאל אובדן מיידי.",
            "⚠ קליטה ידנית CRM<br>מזכירה מזינה פרטים ידנית – 15-30 דק׳ לפנייה.<br>שגיאות הקלדה, שדות חסרים, עיכוב.",
            "בדיקת פרטים<br>יועץ בודק שלמות הנתונים ידנית.<br>תהליך איטי, לעיתים חסר מידע קריטי.",
            "⚠ השלמת מידע<br>חזרה ללקוח לבקש פרטים חסרים.<br>~30% מהמקרים – לקוח מתוסכל מהעיכוב.",
            "⚠ העתקה ל-Drive<br>העתקה ידנית של מסמכים CRM ← Drive.<br>כפילות מידע, טעויות, בזבוז זמן.",
            "⚠ שיוך ליועץ<br>חלוקה לפי עומס עבודה בלבד.<br>ללא התחשבות בהתמחות – שיוך לא אופטימלי.",
            "⚠ קשר ראשוני 48h<br>זמן תגובה 24-48 שעות!<br>מתחרים עונים תוך שעות, הלקוח כבר במקום אחר.",
            "⚠ תזכורות ידניות<br>3 ניסיונות התקשרות ידניים.<br>ללא תזמון אופטימלי – לקוחות רבים לא עונים.",
            "פגישת ייעוץ<br>המפגש המרכזי עם הלקוח.<br>רק ~45% מהפניות מגיעות לשלב זה.",
            "⚠ הצעה ידנית<br>יועץ מכין הצעה מאפס ב-Word/Excel.<br>2-4 שעות עבודה, סיכוי לטעויות.",
            "פולואפ<br>מעקב ידני אחר תשובת הלקוח.<br>ללא מערכת התראות – לעיתים נשכח.",
            "✓ לקוח משלם<br>הצלחה! הלקוח הפך ללקוח משלם.<br>רק ~20% מהפניות מגיעות לכאן.",
            "✗ סגירת פנייה<br>הפנייה נסגרת ללא המרה.<br>אובדן הזדמנות עסקית.",
        ]
        sources =  [0, 1, 2, 2, 3, 4, 5, 6, 6, 7, 7, 8, 9, 10, 10]
        targets =  [1, 2, 3, 4, 2, 5, 6, 7, 8, 8, 12, 9, 10, 11, 12]
        values  =  [100,100,30,70,30,70,70,25,45,15,10,45,45,20,25]
        node_colors = [
            "#339af0",  # פנייה - כחול
            "#e03131",  # קליטה ידנית - אדום כשל
            "#fab005",  # בדיקת פרטים - צהוב
            "#e03131",  # השלמת מידע - אדום כשל
            "#e03131",  # העתקה - אדום כשל
            "#e03131",  # שיוך - אדום כשל
            "#e03131",  # קשר 48h - אדום כשל
            "#e03131",  # תזכורות - אדום כשל
            "#51cf66",  # פגישה - ירוק
            "#e03131",  # הצעה ידנית - אדום כשל
            "#fab005",  # פולואפ - צהוב
            "#2b8a3e",  # לקוח משלם - ירוק כהה
            "#868e96",  # סגירה - אפור
        ]
        title = "תהליך נוכחי (As-Is) – צווארי בקבוק מסומנים באדום"
    else:
        labels = [
            "פנייה נכנסת",            # 0
            "✓ קליטה אוטומטית",      # 1
            "✓ AI השלמת נתונים",     # 2
            "הודעה אוטומטית",        # 3
            "✓ Lead Scoring",        # 4
            "👁 ביקורת ניקוד (HITL)", # 5
            "✓ ניתוב חכם",           # 6
            "✓ AI תקציר+מייל",       # 7
            "👁 יועץ מאשר (HITL)",    # 8
            "Escalation",            # 9
            "✓ פגישה Calendly",      # 10
            "✓ הצעה AI Draft",       # 11
            "👁 אישור הצעה (HITL)",   # 12
            "✓ לקוח + Onboarding",   # 13
            "Nurture אוטומטי",       # 14
        ]
        node_descriptions = [
            "פנייה נכנסת<br>לקוח פוטנציאלי פונה דרך כל ערוץ.<br>קליטה אוטומטית מיידית.",
            "✓ קליטה אוטומטית<br>כל הערוצים נקלטים ל-CRM אוטומטית.<br>API / Zapier – אפס מגע ידני.",
            "✓ AI השלמת נתונים<br>AI מזהה שדות חסרים ומשלים<br>מ-APIs חיצוניים אוטומטית.",
            "הודעה אוטומטית<br>מייל/SMS ללקוח תוך דקות.<br>\"קיבלנו את פנייתך, ניצור קשר בקרוב.\"",
            "✓ Lead Scoring<br>AI מדרג כל פנייה לפי פוטנציאל:<br>גודל תיק, סיכוי סגירה, שביעות רצון.",
            "👁 ביקורת ניקוד (HITL)<br>מנהל מכירות סוקר ומאשר ניקוד לידים גבוהים.<br>ביקורת אנושית למניעת שגיאות AI.",
            "✓ ניתוב חכם<br>שיוך אוטומטי ליועץ לפי התמחות,<br>עומס, וציון הליד.",
            "✓ AI תקציר+מייל<br>AI מכין תקציר לקוח וטיוטת מייל.<br>היועץ מקבל חבילה מוכנה.",
            "👁 יועץ מאשר (HITL)<br>יועץ סוקר תקציר AI ומאשר/עורך מייל<br>לפני שליחה. ביקורת אנושית חובה.",
            "Escalation<br>פנייה מורכבת מועברת למנהל.<br>טיפול ידני מלא עם תמיכת AI.",
            "✓ פגישה Calendly<br>קביעת פגישה אוטומטית.<br>הלקוח בוחר מועד נוח.",
            "✓ הצעה AI Draft<br>AI מכין טיוטת הצעה מותאמת אישית<br>על בסיס פרופיל הלקוח והפגישה.",
            "👁 אישור הצעה (HITL)<br>יועץ סוקר ומאשר הצעה סופית<br>לפני שליחה ללקוח. ביקורת אנושית חובה.",
            "✓ לקוח + Onboarding<br>לקוח חתם! Onboarding אוטומטי<br>כולל מיילים, הדרכות, ומעקב.",
            "Nurture אוטומטי<br>סדרת תוכן אוטומטית<br>לשימור קשר עם לקוחות שלא סגרו.",
        ]
        sources = [0, 1, 2, 2, 3, 4, 5, 6, 7, 8, 8, 9, 10, 11, 12, 12]
        targets = [1, 2, 3, 4, 2, 5, 6, 7, 8, 9, 10, 10, 11, 12, 13, 14]
        values  = [100, 100, 15, 85, 15, 85, 85, 85, 85, 20, 65, 20, 85, 85, 50, 35]
        node_colors = [
            "#339af0",  # פנייה - כחול
            "#2b8a3e",  # קליטה אוטו - ירוק כהה
            "#2b8a3e",  # AI נתונים - ירוק כהה
            "#339af0",  # הודעה - כחול
            "#2b8a3e",  # Lead Score - ירוק כהה
            "#f59f00",  # HITL ביקורת ניקוד - כתום
            "#2b8a3e",  # ניתוב - ירוק כהה
            "#2b8a3e",  # AI תקציר - ירוק כהה
            "#f59f00",  # HITL יועץ מאשר - כתום
            "#fab005",  # Escalation - צהוב
            "#2b8a3e",  # Calendly - ירוק כהה
            "#2b8a3e",  # AI Draft - ירוק כהה
            "#f59f00",  # HITL אישור הצעה - כתום
            "#51cf66",  # לקוח משלם - ירוק בהיר
            "#339af0",  # Nurture - כחול
        ]
        title = "תהליך יעד (To-Be) – AI בירוק, ביקורת אנושית (HITL) בכתום"

    # -- Dynamic coloring with optional highlighting --
    if highlight_node is not None:
        final_node_colors = ["rgba(100,100,100,0.15)"] * len(labels)
        final_node_colors[highlight_node] = node_colors[highlight_node]
        link_colors = []
        for s, t in zip(sources, targets):
            if s == highlight_node or t == highlight_node:
                link_colors.append(_hex_to_rgba(node_colors[highlight_node], 0.55))
                final_node_colors[s] = node_colors[s]
                final_node_colors[t] = node_colors[t]
            else:
                link_colors.append("rgba(100,100,100,0.04)")
    else:
        final_node_colors = node_colors
        link_colors = [_hex_to_rgba(node_colors[s], 0.4) for s, t in zip(sources, targets)]

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node={
            "pad": 30, "thickness": 22,
            "line": {"color": "rgba(255,255,255,0.3)", "width": 1},
            "label": labels,
            "color": final_node_colors,
            "customdata": node_descriptions,
            "hovertemplate": "%{customdata}<extra></extra>",
        },
        link={"source": sources, "target": targets, "value": values, "color": link_colors},
    )])
    fig.update_layout(
        title_text=title, font_size=14,
        margin={"l": 20, "r": 20, "t": 55, "b": 15},
        height=550,
    )
    return fig


def show_kpis(filtered: pd.DataFrame) -> None:
    total = len(filtered)
    active = int((filtered["סטטוס"] == "פעיל").sum())
    churned = total - active
    churn_rate = (churned / total * 100) if total else 0
    avg_response = filtered["זמן_תגובה_ממוצע_שעות"].mean() if total else 0
    avg_sat = filtered["שביעות_רצון"].mean() if total else 0
    avg_income = filtered["הכנסה_חודשית"].mean() if total else 0
    avg_portfolio = filtered["סכום_תיק"].mean() if total else 0

    # ── Executive KPIs ──
    st.markdown(
        '<div class="colored-header" style="background:linear-gradient(90deg,#1971c2,#0ca678);color:white;">'
        'סיכום מנהלים – מבט-על</div>',
        unsafe_allow_html=True,
    )
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("סה״כ לקוחות", f"{total:,}")
    k2.metric("פעילים", f"{active:,}")
    k3.metric("נטשו", f"{churned:,}")
    k4.metric("שיעור נטישה", f"{churn_rate:.1f}%")
    k5.metric("שביעות רצון ⌀", f"{avg_sat:.1f}")
    k6.metric("זמן תגובה ⌀", f"{avg_response:.0f}h")

    k7, k8, k9 = st.columns(3)
    k7.metric("סכום תיק ממוצע", f"₪{avg_portfolio:,.0f}")
    k8.metric("הכנסה חודשית ⌀", f"₪{avg_income:,.0f}")
    k9.metric("מספר ערים ייחודיות", f"{filtered['עיר'].nunique()}")

    # ── CEO Quick Insights ──
    st.markdown(
        '<div class="colored-header" style="background:#e67700;color:white; font-size:1.3rem;">'
        'תובנות מהירות למנכ״ל</div>',
        unsafe_allow_html=True,
    )
    insights = _generate_ceo_insights(filtered)
    cols = st.columns(len(insights))
    _insight_colors = ["#e03131", "#1971c2", "#0ca678", "#7048e8", "#e67700"]
    for i, (icon, text) in enumerate(insights):
        color = _insight_colors[i % len(_insight_colors)]
        cols[i].markdown(
            f'<div style="background:{color}12; border-right:4px solid {color}; '
            f'padding:12px 14px; border-radius:8px; min-height:80px;">'
            f'<span style="font-size:1rem; font-weight:600; direction:rtl; display:block;">{text}</span></div>',
            unsafe_allow_html=True,
        )


def _generate_ceo_insights(df: pd.DataFrame) -> list[tuple[str, str]]:
    """Auto-generate key insights from the filtered data."""
    insights: list[tuple[str, str]] = []

    # Worst city by churn
    city_churn = df.groupby("עיר")["סטטוס"].apply(lambda s: (s == "לא פעיל").mean())
    if not city_churn.empty:
        worst_city = city_churn.idxmax()
        worst_pct = city_churn.max() * 100
        insights.append(("", f"העיר עם הנטישה הגבוהה ביותר: <b>{worst_city}</b> ({worst_pct:.0f}%)"))

    # Worst service
    svc_churn = df.groupby("סוג_שירות")["סטטוס"].apply(lambda s: (s == "לא פעיל").mean())
    if not svc_churn.empty:
        worst_svc = svc_churn.idxmax()
        insights.append(("", f"שירות בסיכון: <b>{worst_svc}</b> – נטישה {svc_churn.max()*100:.0f}%"))

    # Response time gap
    active_resp = df.loc[df["סטטוס"] == "פעיל", "זמן_תגובה_ממוצע_שעות"].mean()
    churn_resp = df.loc[df["סטטוס"] == "לא פעיל", "זמן_תגובה_ממוצע_שעות"].mean()
    if pd.notna(active_resp) and pd.notna(churn_resp):
        insights.append(("", f"זמן תגובה: פעילים {active_resp:.0f}h לעומת נוטשים {churn_resp:.0f}h"))

    # High-value at risk
    high_val_churned = df[(df["סכום_תיק"] > df["סכום_תיק"].quantile(0.75)) & (df["סטטוס"] == "לא פעיל")]
    if len(high_val_churned):
        total_lost = high_val_churned["סכום_תיק"].sum()
        insights.append(("", f"{len(high_val_churned)} לקוחות גדולים נטשו – שווי תיקים: ₪{total_lost:,.0f}"))

    # Satisfaction
    low_sat = df[df["שביעות_רצון"] <= 3]
    if len(low_sat):
        pct_low = len(low_sat) / len(df) * 100 if len(df) else 0
        insights.append(("", f"{pct_low:.0f}% מהלקוחות עם שביעות רצון 3 ומטה — דורשים התערבות מיידית"))

    return insights[:5]


def show_visuals(filtered: pd.DataFrame) -> None:
    # ── Row 1: Churn Trend + Status Donut ──
    st.markdown(
        '<div class="colored-header" style="background:#1971c2;color:white;">'
        'מגמות נטישה והתפלגות סטטוס</div>',
        unsafe_allow_html=True,
    )
    col_trend, col_donut = st.columns([3, 2])

    with col_trend:
        churn_over_time = (
            filtered.dropna(subset=["תאריך_נטישה"])
            .groupby(filtered["תאריך_נטישה"].dt.to_period("M"))
            .size()
            .reset_index(name="כמות_נטישה")
        )
        if not churn_over_time.empty:
            churn_over_time["תאריך_נטישה"] = churn_over_time["תאריך_נטישה"].astype(str)
            fig_churn = px.area(
                churn_over_time,
                x="תאריך_נטישה",
                y="כמות_נטישה",
                markers=True,
                title="נטישה לאורך זמן (חודשי)",
            )
            fig_churn.update_traces(fill="tozeroy", line_color="#e03131")
            fig_churn.update_layout(
                xaxis_title="חודש-שנה",
                yaxis_title="מספר נוטשים",
                xaxis=dict(tickangle=-45, nticks=20),
            )
            st.plotly_chart(fig_churn, use_container_width=True)
        else:
            st.info("אין נתוני נטישה בתצוגה הנוכחית.")

    with col_donut:
        status_counts = filtered["סטטוס"].value_counts().reset_index()
        status_counts.columns = ["סטטוס", "כמות"]
        fig_donut = px.pie(
            status_counts, names="סטטוס", values="כמות",
            hole=0.5, title="התפלגות סטטוס לקוחות",
            color="סטטוס",
            color_discrete_map={"פעיל": "#2b8a3e", "לא פעיל": "#e03131"},
        )
        fig_donut.update_traces(textinfo="label+percent", textposition="outside")
        st.plotly_chart(fig_donut, use_container_width=True)

    # ── Row 2: Satisfaction Histogram + Satisfaction by Service ──
    st.markdown(
        '<div class="colored-header" style="background:#0ca678;color:white;">'
        'ניתוח שביעות רצון</div>',
        unsafe_allow_html=True,
    )
    col_hist, col_box = st.columns(2)

    with col_hist:
        # Average satisfaction: active vs churned – clear grouped bar
        _sat_by_svc = (
            filtered.groupby(["סוג_שירות", "סטטוס"], as_index=False)["שביעות_רצון"].mean()
        )
        fig_sat = px.bar(
            _sat_by_svc, x="סוג_שירות", y="שביעות_רצון",
            color="סטטוס", barmode="group",
            title="שביעות רצון ממוצעת – פעילים מול נוטשים",
            color_discrete_map={"פעיל": "#2b8a3e", "לא פעיל": "#e03131"},
            text="שביעות_רצון",
        )
        fig_sat.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig_sat.update_layout(
            xaxis_title="סוג שירות", yaxis_title="שביעות רצון ממוצעת (1-10)",
            yaxis_range=[0, 10], legend_title_text="סטטוס",
        )
        st.plotly_chart(fig_sat, use_container_width=True)
        st.caption("עמודה ירוקה = פעילים, אדומה = נוטשים. שימו לב: הפערים בין העמודות קטנים מאוד (~0.25 נקודה) – שביעות רצון לבדה לא מסבירה נטישה בדאטה הזה.")

    with col_box:
        # Churn rate by satisfaction bucket – intuitive bar
        _sat_df = filtered.copy()
        _sat_df["קבוצת_שביעות_רצון"] = pd.cut(
            _sat_df["שביעות_רצון"],
            bins=[0, 3, 5, 7, 10],
            labels=["נמוכה (1-3)", "בינונית (4-5)", "טובה (6-7)", "גבוהה (8-10)"],
        )
        _churn_by_sat = (
            _sat_df.groupby("קבוצת_שביעות_רצון", observed=True)
            .agg(שיעור_נטישה=("סטטוס", lambda s: (s == "לא פעיל").mean()),
                 לקוחות=("סטטוס", "count"))
            .reset_index()
        )
        fig_box = px.bar(
            _churn_by_sat, x="קבוצת_שביעות_רצון", y="שיעור_נטישה",
            title="שיעור נטישה לפי רמת שביעות רצון – הפער זניח",
            color="שיעור_נטישה", color_continuous_scale="RdYlGn_r",
            text="לקוחות",
        )
        fig_box.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_box.update_layout(
            xaxis_title="רמת שביעות רצון", yaxis_title="שיעור נטישה",
            yaxis_tickformat=".0%", coloraxis_showscale=False,
        )
        st.plotly_chart(fig_box, use_container_width=True)
        st.caption("ממצא מפתיע: שיעור הנטישה כמעט זהה בכל הקבוצות (34%-38%). שביעות רצון לבדה אינה מנבאת נטישה – הסיבות האמיתיות נמצאות בפילוח גיאוגרפי וסוג שירות.")

    # ── Row 3: Income by Service + Churn by City ──
    st.markdown(
        '<div class="colored-header" style="background:#7048e8;color:white;">'
        'הכנסות ונטישה לפי פילוח</div>',
        unsafe_allow_html=True,
    )
    col_svc, col_city = st.columns(2)

    with col_svc:
        svc_income = filtered.groupby("סוג_שירות", as_index=False).agg(
            הכנסה_ממוצעת=("הכנסה_חודשית", "mean"),
            לקוחות=("client_id", "count"),
        )
        fig_income = px.bar(
            svc_income, x="סוג_שירות", y="הכנסה_ממוצעת",
            text="לקוחות",
            title="הכנסה חודשית ממוצעת לפי סוג שירות",
            color="הכנסה_ממוצעת",
            color_continuous_scale="Teal",
        )
        fig_income.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_income.update_layout(xaxis_title="סוג שירות", yaxis_title="₪ הכנסה ממוצעת",
                                  coloraxis_showscale=False)
        st.plotly_chart(fig_income, use_container_width=True)

    with col_city:
        city_churn = (
            filtered.groupby("עיר", as_index=False)
            .agg(נטישה=("סטטוס", lambda s: (s == "לא פעיל").mean()), לקוחות=("client_id", "count"))
        )
        city_churn = city_churn[city_churn["לקוחות"] >= 5].sort_values("נטישה", ascending=True).tail(10)
        fig_city = px.bar(
            city_churn, y="עיר", x="נטישה", orientation="h",
            title="10 ערים עם נטישה גבוהה (לפחות 5 לקוחות)",
            text="לקוחות",
            color="נטישה", color_continuous_scale="Reds",
        )
        fig_city.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="inside", insidetextanchor="end")
        fig_city.update_layout(xaxis_tickformat=".0%", xaxis_title="שיעור נטישה",
                                yaxis_title="", coloraxis_showscale=False)
        st.plotly_chart(fig_city, use_container_width=True)

    # ── Row 4: Age Distribution + Response Time ──
    st.markdown(
        '<div class="colored-header" style="background:#e67700;color:white;">'
        'דמוגרפיה וזמני תגובה</div>',
        unsafe_allow_html=True,
    )
    col_age, col_resp = st.columns(2)

    with col_age:
        fig_age = px.histogram(
            filtered, x="גיל", color="סטטוס", nbins=20, barmode="group",
            title="התפלגות גילאים לפי סטטוס",
            color_discrete_map={"פעיל": "#2b8a3e", "לא פעיל": "#e03131"},
            category_orders={"סטטוס": ["פעיל", "לא פעיל"]},
        )
        fig_age.update_traces(opacity=0.85)
        fig_age.update_layout(xaxis_title="גיל", yaxis_title="מספר לקוחות")
        st.plotly_chart(fig_age, use_container_width=True)

    with col_resp:
        fig_resp = px.histogram(
            filtered, x="זמן_תגובה_ממוצע_שעות", color="סטטוס", nbins=25, barmode="group",
            title="התפלגות זמן תגובה (שעות) לפי סטטוס",
            color_discrete_map={"פעיל": "#2b8a3e", "לא פעיל": "#e03131"},
            category_orders={"סטטוס": ["פעיל", "לא פעיל"]},
        )
        fig_resp.update_traces(opacity=0.85)
        fig_resp.update_layout(xaxis_title="שעות", yaxis_title="מספר לקוחות")
        st.plotly_chart(fig_resp, use_container_width=True)

    # ── Row 5: What actually drives churn? ──
    st.markdown(
        '<div class="colored-header" style="background:#e03131;color:white;">'
        'מה באמת גורם לנטישה? ניתוח אובייקטיבי</div>',
        unsafe_allow_html=True,
    )

    col_combo, col_profile = st.columns(2)

    with col_combo:
        # Top city×service combos by churn rate
        _combo_df = filtered.copy()
        _combo_df["churn"] = (_combo_df["סטטוס"] == "לא פעיל").astype(int)
        _combo = (
            _combo_df.groupby(["עיר", "סוג_שירות"], as_index=False)
            .agg(n=("churn", "count"), rate=("churn", "mean"))
        )
        _combo = _combo[_combo["n"] >= 10].sort_values("rate", ascending=True).tail(10)
        _combo["תיאור"] = _combo["עיר"] + " + " + _combo["סוג_שירות"]
        fig_combo = px.bar(
            _combo, y="תיאור", x="rate", orientation="h",
            title="10 שילובי עיר ושירות עם נטישה גבוהה (10 לקוחות ומעלה)",
            text="n",
            color="rate", color_continuous_scale="Reds",
        )
        fig_combo.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="inside", insidetextanchor="end")
        fig_combo.update_layout(
            xaxis_tickformat=".0%", xaxis_title="שיעור נטישה",
            yaxis_title="", coloraxis_showscale=False,
            height=450,
        )
        st.plotly_chart(fig_combo, use_container_width=True)
        st.caption("רמת גן + ביטוח בריאות: כ-79% נטישה — החריג הבולט ביותר. שילוב עיר ושירות חושף דפוסים שמשתנה בודד מחמיץ.")

    with col_profile:
        # Risk profile comparison – multi-variable
        _risk_df = filtered.copy()
        _risk_df["churn"] = (_risk_df["סטטוס"] == "לא פעיל").astype(int)
        _r_risk = _risk_df[
            (_risk_df["שביעות_רצון"] <= 4)
            & (_risk_df["זמן_תגובה_ממוצע_שעות"] >= 60)
            & (_risk_df["מספר_פניות_שנה_אחרונה"] <= 4)
        ]
        _r_safe = _risk_df[
            (_risk_df["שביעות_רצון"] >= 7)
            & (_risk_df["זמן_תגובה_ממוצע_שעות"] <= 30)
            & (_risk_df["מספר_פניות_שנה_אחרונה"] >= 8)
        ]
        _r_all = _risk_df["churn"].mean()
        _profiles_dash = pd.DataFrame([
            {"פרופיל": "פרופיל מסוכן", "שיעור_נטישה": _r_risk["churn"].mean() if len(_r_risk) else 0, "n": len(_r_risk), "color": "#e03131"},
            {"פרופיל": "ממוצע כללי", "שיעור_נטישה": _r_all, "n": len(_risk_df), "color": "#e67700"},
            {"פרופיל": "פרופיל בטוח", "שיעור_נטישה": _r_safe["churn"].mean() if len(_r_safe) else 0, "n": len(_r_safe), "color": "#2b8a3e"},
        ])
        fig_prof = px.bar(
            _profiles_dash, x="פרופיל", y="שיעור_נטישה",
            title="פרופילי סיכון: שילוב 3 משתנים חושף פערי נטישה",
            color="שיעור_נטישה", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_prof.update_traces(texttemplate="מתוך <b>%{text}</b> לקוחות", textposition="outside")
        fig_prof.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="", coloraxis_showscale=False,
            height=450, yaxis_range=[0, 0.55],
        )
        st.plotly_chart(fig_prof, use_container_width=True)
        st.caption(
            "🔴 פרופיל מסוכן: שביעות רצון נמוכה (עד 4), זמן תגובה ארוך (מעל 60 שעות), מעט פניות (עד 4). "
            "🟢 פרופיל בטוח: שביעות רצון גבוהה (7 ומעלה), תגובה מהירה (עד 30 שעות), פניות תכופות (8 ומעלה). "
            "הפער בין הפרופילים: כ-16 נקודות אחוז – שילוב המשתנים חזק יותר מכל משתנה בנפרד."
        )


def show_part_a() -> None:
    """Full Part A – Process Analysis section with all 5 deliverables."""

    st.markdown(
        '<div style="background: linear-gradient(90deg, #1971c2, #0ca678); padding: 1rem 1.5rem; '
        'border-radius: 10px; color: white; font-size: 1.6rem; font-weight: 700; text-align: center;">'
        'חלק א׳ – אפיון תהליך עבודה</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "התהליך במוקד האפיון: \"טיפול בפנייה חדשה של לקוח פוטנציאלי\" – "
        "מהרגע שהלקוח פונה ועד שהוא הופך ללקוח משלם או שהפנייה נסגרת."
    )

    # ── 1. תרשים תהליך נוכחי ──────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#e03131; color:white;">'
        '1. מפת תהליך נוכחי (As-Is) – צווארי בקבוק</div>',
        unsafe_allow_html=True,
    )
    st.markdown("""
**בעלי תפקיד:** מזכירה / נציג קליטה · יועץ פיננסי · מנהל מכירות  
**מערכות:** Monday.com CRM · Google Drive · Mailchimp · Excel · Slack/Teams
""")

    # Visual timeline roadmap for As-Is
    _as_is_steps = [
        ("1", "קליטת פנייה", "מזכירה", False),
        ("2", "הזנה ידנית ל-CRM", "מזכירה", True),
        ("3", "בדיקת שלמות פרטים", "יועץ", False),
        ("4", "חזרה ללקוח להשלמת מידע", "יועץ", True),
        ("5", "העתקת מסמכים ל-Drive", "יועץ", True),
        ("6", "שיוך ליועץ לפי עומס", "מנהל מכירות", True),
        ("7", "יצירת קשר ראשוני (24-48h)", "יועץ", True),
        ("8", "תזכורות ידניות (3 ניסיונות)", "יועץ", True),
        ("9", "פגישת ייעוץ", "יועץ", False),
        ("10", "הכנת הצעה ידנית", "יועץ", True),
        ("11", "שליחת הצעה + פולואפ", "יועץ", False),
        ("12", "לקוח משלם / סגירה", "—", False),
    ]
    _roadmap_html = '<div style="display:flex; flex-wrap:wrap; gap:8px; direction:rtl; margin:0.8rem 0;">'
    for step_num, step_name, owner, is_bottleneck in _as_is_steps:
        bg = "#e03131" if is_bottleneck else "#339af0"
        icon = "⚠" if is_bottleneck else "←"
        _roadmap_html += (
            f'<div style="background:{bg}; color:white; padding:10px 14px; border-radius:8px; '
            f'min-width:140px; text-align:center; flex:1;">'
            f'<div style="font-size:0.75rem; opacity:0.85;">שלב {step_num} · {owner}</div>'
            f'<div style="font-size:1rem; font-weight:600;">{icon} {step_name}</div></div>'
        )
    _roadmap_html += '</div>'
    _roadmap_html += '<p style="direction:rtl; font-size:1rem; font-weight:600; color:#ccc;"><span style="color:#e03131; font-size:1.1rem;">&#x26A0;</span> = צוואר בקבוק / פעולה ידנית מיותרת</p>'
    st.markdown(_roadmap_html, unsafe_allow_html=True)

    _asis_steps_list = [
        "פנייה נכנסת", "קליטה ידנית CRM", "בדיקת פרטים", "השלמת מידע",
        "העתקה ל-Drive", "שיוך ליועץ", "קשר ראשוני 48h",
        "תזכורות ידניות", "פגישת ייעוץ", "הצעה ידנית", "פולואפ",
        "לקוח משלם", "סגירת פנייה",
    ]
    _asis_options = ["הכל"] + _asis_steps_list
    _asis_hl_val = st.session_state.get("asis_hl", "הכל")
    _asis_hl = None if _asis_hl_val == "הכל" else _asis_options.index(_asis_hl_val) - 1
    st.plotly_chart(_process_flowchart(is_current=True, highlight_node=_asis_hl), use_container_width=True)
    st.radio("הדגש שלב בגרף:", _asis_options, horizontal=True, key="asis_hl")

    # ── 2. תיאור כתוב: צווארי בקבוק, AI, סיכונים ─────────────────
    st.markdown(
        '<div class="colored-header" style="background:#fab005; color:#1a1a1a;">'
        '2. צווארי בקבוק, הזדמנויות AI ואוטומציה וסיכונים</div>',
        unsafe_allow_html=True,
    )

    bot_col, ai_col, risk_col = st.columns(3)

    with bot_col:
        st.markdown(
            '<div style="background:#e0313120; border-right:4px solid #e03131; padding:0.8rem; border-radius:6px; direction:rtl; text-align:right;">'
            '<h4 style="color:#e03131; margin:0 0 0.5rem 0;">🔴 צווארי בקבוק</h4>'
            '<ul style="padding-right:1.2em; padding-left:0; margin:0.3rem 0 0 0;">'
            '<li><b>קליטה ידנית:</b> כל פנייה מוזנת ידנית ל-CRM – טעויות, שדות חסרים, עיכוב.</li>'
            '<li><b>זמן תגובה 48 שעות:</b> מתחרים עונים באותו יום, הלקוח כבר פנה למקום אחר.</li>'
            '<li><b>שיוך לא חכם:</b> פניות מחולקות ללא התחשבות בהתמחות היועץ.</li>'
            '<li><b>העתקות ידניות:</b> CRM ← Drive ← Excel ← Mailchimp – בזבוז ושגיאות.</li>'
            '<li><b>אין תעדוף:</b> תיק של ₪2M מטופל כמו תיק של ₪50K.</li>'
            '</ul></div>',
            unsafe_allow_html=True,
        )

    with ai_col:
        st.markdown(
            '<div style="background:#2b8a3e20; border-right:4px solid #2b8a3e; padding:0.8rem; border-radius:6px; direction:rtl; text-align:right;">'
            '<h4 style="color:#2b8a3e; margin:0 0 0.5rem 0;">🟢 הזדמנויות AI ואוטומציה</h4>'
            '<ul style="padding-right:1.2em; padding-left:0; margin:0.3rem 0 0 0;">'
            '<li><b>קליטה אוטומטית:</b> כל הערוצים ל-CRM בלי מגע ידני (API / Zapier).</li>'
            '<li><b>AI Lead Scoring:</b> דירוג לפי פוטנציאל, שביעות רצון, סיכוי סגירה.</li>'
            '<li><b>תגובה מיידית:</b> מייל/SMS אוטומטי תוך דקות.</li>'
            '<li><b>ניתוב חכם:</b> שיוך ליועץ לפי התמחות ועומס.</li>'
            '<li><b>טיוטת הצעה AI:</b> לאישור יועץ בלחיצה.</li>'
            '<li><b>מעקב אוטומטי:</b> תזכורות ודוחות בזמן אמת.</li>'
            '</ul></div>',
            unsafe_allow_html=True,
        )

    with risk_col:
        st.markdown(
            '<div style="background:#fab00520; border-right:4px solid #fab005; padding:0.8rem; border-radius:6px; direction:rtl; text-align:right;">'
            '<h4 style="color:#fab005; margin:0 0 0.5rem 0;">🟡 סיכונים</h4>'
            '<ul style="padding-right:1.2em; padding-left:0; margin:0.3rem 0 0 0;">'
            '<li><b>דיוק AI:</b> ציון שגוי → הזנחת לקוח חשוב.</li>'
            '<li><b>פרטיות ורגולציה:</b> מידע פיננסי חייב הצפנה ובקרות.</li>'
            '<li><b>אובדן מגע אישי:</b> אוטומציה מוגזמת מרתיעה.</li>'
            '<li><b>התנגדות עובדים:</b> יועצים ותיקים מתנגדים לשינוי.</li>'
            '<li><b>Human-in-the-loop:</b> AI מציע, אדם מחליט.</li>'
            '</ul></div>',
            unsafe_allow_html=True,
        )

    # ── 3. תרשים תהליך יעד ────────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#2b8a3e; color:white;">'
        '3. מפת תהליך יעד (To-Be) – שלבים אוטומטיים</div>',
        unsafe_allow_html=True,
    )

    _to_be_steps = [
        ("1", "קליטה אוטומטית", "API", "ai"),
        ("2", "AI השלמת נתונים", "AI", "ai"),
        ("3", "Lead Scoring", "AI", "ai"),
        ("4", "ביקורת ניקוד", "HITL", "hitl"),
        ("5", "ניתוב חכם ליועץ", "AI", "ai"),
        ("6", "AI תקציר + מייל", "AI", "ai"),
        ("7", "יועץ מאשר (<4h)", "HITL", "hitl"),
        ("8", "פגישה Calendly", "אוטומטי", "ai"),
        ("9", "הצעה AI Draft", "AI", "ai"),
        ("10", "אישור הצעה סופי", "HITL", "hitl"),
        ("11", "לקוח משלם / Nurture", "—", "other"),
    ]
    _roadmap2 = '<div style="display:flex; flex-wrap:wrap; gap:8px; direction:rtl; margin:0.8rem 0;">'
    for step_num, step_name, owner, step_type in _to_be_steps:
        if step_type == "hitl":
            bg = "#f59f00"
            icon = "👁"
        elif step_type == "ai":
            bg = "#2b8a3e"
            icon = "✓"
        else:
            bg = "#339af0"
            icon = "←"
        _roadmap2 += (
            f'<div style="background:{bg}; color:white; padding:10px 14px; border-radius:8px; '
            f'min-width:140px; text-align:center; flex:1;">'
            f'<div style="font-size:0.75rem; opacity:0.85;">שלב {step_num} · {owner}</div>'
            f'<div style="font-size:1rem; font-weight:600;">{icon} {step_name}</div></div>'
        )
    _roadmap2 += '</div>'
    _roadmap2 += ('<p style="direction:rtl; font-size:1rem; font-weight:600; color:#ccc;">'
        '<span style="color:#51cf66; font-size:1.1rem;">✓</span> = אוטומטי / AI'
        '&nbsp;&nbsp;|&nbsp;&nbsp;'
        '<span style="background:#f59f00; color:#1a1a1a; border-radius:4px; padding:1px 7px; font-size:1.15rem; font-weight:900;">👁</span>'
        ' = ביקורת אנושית (Human-in-the-Loop)'
        '</p>')
    st.markdown(_roadmap2, unsafe_allow_html=True)

    _tobe_steps_list = [
        "פנייה נכנסת", "קליטה אוטומטית", "AI השלמת נתונים", "הודעה אוטומטית",
        "Lead Scoring", "ביקורת ניקוד (HITL)", "ניתוב חכם",
        "AI תקציר+מייל", "יועץ מאשר (HITL)", "Escalation",
        "פגישה Calendly", "הצעה AI Draft", "אישור הצעה (HITL)",
        "לקוח + Onboarding", "Nurture אוטומטי",
    ]
    _tobe_options = ["הכל"] + _tobe_steps_list
    _tobe_hl_val = st.session_state.get("tobe_hl", "הכל")
    _tobe_hl = None if _tobe_hl_val == "הכל" else _tobe_options.index(_tobe_hl_val) - 1
    st.plotly_chart(_process_flowchart(is_current=False, highlight_node=_tobe_hl), use_container_width=True)
    st.radio("הדגש שלב בגרף:", _tobe_options, horizontal=True, key="tobe_hl")

    # ── 4. שלוש שאלות למנכ"ל ──────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#1971c2; color:white;">'
        '4. שלוש שאלות למנכ"ל בפגישה הראשונה</div>',
        unsafe_allow_html=True,
    )
    q1, q2, q3 = st.columns(3)
    with q1:
        st.info(
            '**שאלה 1 – מדד הצלחה:**\n\n'
            '"מה ה-KPI המרכזי שלך ל-90 הימים הקרובים – '
            'קיצור זמן תגובה, שיפור שיעור ההמרה מפנייה ללקוח, או צמצום נטישה? '
            'התשובה תקבע את סדר העדיפויות שלנו."'
        )
    with q2:
        st.info(
            '**שאלה 2 – נקודות כשל:**\n\n'
            '"באיזה שלב בתהליך הכי הרבה פניות נופלות בין הכיסאות היום – '
            'בקליטה, בהמתנה לתגובה, או אחרי שליחת ההצעה? '
            'ומה לפי הערכתך העלות העסקית החודשית של זה?"'
        )
    with q3:
        st.info(
            '**שאלה 3 – מגבלות טכנולוגיות ורגולטוריות:**\n\n'
            '"האם יש מערכות שאסור לחבר באינטגרציה מלאה מסיבות רגולטוריות, '
            'אבטחת מידע או מגבלות IT? ומה מדיניות החברה לגבי שליחת תקשורת אוטומטית ללקוחות?"'
        )

    # ── 5. בעלי תפקידים לפגישות ────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#7048e8; color:white;">'
        '5. דמויות מפתח בחברה לפגישות אפיון</div>',
        unsafe_allow_html=True,
    )
    _people = [
        ("מנהל/ת מכירות", "מכיר/ה את תהליך הפניות מקצה לקצה, יודע/ת היכן נתקעות עסקאות ומה ה-SLA בפועל.", "#e03131"),
        ("יועץ/ת פיננסי (2-3)", 'משתמשי קצה – לאמת פער בין "איך אמור" ל"איך באמת עובד".', "#1971c2"),
        ("מנהל/ת תפעול", "ממפה העברות ידניות בין מערכות, שגיאות חוזרות, זמני טיפול.", "#e67700"),
        ("מנהל/ת IT", "מגדיר/ה מה אפשרי טכנולוגית – אינטגרציות, API, תשתית.", "#0ca678"),
        ("אחראי/ת Compliance", "תקנות פיננסיות, פרטיות מידע – חייבים הסכמה לפני אוטומציה.", "#7048e8"),
    ]
    _people_html = '<div style="display:flex; flex-wrap:wrap; gap:10px; direction:rtl; margin:0.8rem 0;">'
    for name, reason, color in _people:
        _people_html += (
            f'<div style="background:{color}18; border-right:4px solid {color}; padding:12px 16px; '
            f'border-radius:8px; flex:1; min-width:200px;">'
            f'<div style="font-weight:700; color:{color}; font-size:1rem;">{name}</div>'
            f'<div style="font-size:0.85rem; margin-top:4px;">{reason}</div></div>'
        )
    _people_html += '</div>'
    st.markdown(_people_html, unsafe_allow_html=True)



def show_cross_tab(filtered: pd.DataFrame) -> None:
    """Interactive cross-tabulation explorer."""
    st.markdown(
        '<div class="colored-header" style="background:linear-gradient(90deg,#7048e8,#1971c2);color:white;">'
        'מצליב הנתונים – הצלבה אינטראקטיבית</div>',
        unsafe_allow_html=True,
    )
    st.caption("בחרו ציר פילוח, מדד לחישוב, ופילוח משני (צבע) – הגרף מתעדכן מיידית.")

    _cross_df = filtered.copy()
    _cross_df["churn"] = (_cross_df["סטטוס"] == "לא פעיל").astype(int)
    _cross_df["קבוצת_גיל"] = pd.cut(
        _cross_df["גיל"], bins=[20, 30, 40, 50, 60, 70],
        labels=["20-30", "31-40", "41-50", "51-60", "61-70"],
    )
    _cross_df["רמת_שביעות_רצון"] = pd.cut(
        _cross_df["שביעות_רצון"], bins=[0, 3, 5, 7, 10],
        labels=["נמוכה (1-3)", "בינונית (4-5)", "טובה (6-7)", "גבוהה (8-10)"],
    )

    _axis_options = {
        "עיר": "עיר",
        "סוג שירות": "סוג_שירות",
        "סטטוס": "סטטוס",
        "קבוצת גיל": "קבוצת_גיל",
        "רמת שביעות רצון": "רמת_שביעות_רצון",
    }
    _metric_options = {
        "שיעור נטישה (%)": ("churn", "mean", ".0%"),
        "שביעות רצון ממוצעת": ("שביעות_רצון", "mean", ".1f"),
        "זמן תגובה ממוצע (שעות)": ("זמן_תגובה_ממוצע_שעות", "mean", ".0f"),
        "הכנסה חודשית ממוצעת (₪)": ("הכנסה_חודשית", "mean", ",.0f"),
        "סכום תיק ממוצע (₪)": ("סכום_תיק", "mean", ",.0f"),
        "כמות לקוחות": ("client_id", "count", ","),
    }

    sel1, sel2, sel3 = st.columns(3)
    with sel1:
        x_label = st.selectbox("ציר פילוח (X):", list(_axis_options.keys()), index=0, key="cross_x")
    with sel2:
        metric_label = st.selectbox("מדד לחישוב:", list(_metric_options.keys()), index=0, key="cross_metric")
    with sel3:
        color_choices = ["ללא"] + [k for k in _axis_options if k != x_label]
        color_label = st.selectbox("פילוח משני (צבע):", color_choices, index=0, key="cross_color")

    x_col = _axis_options[x_label]
    metric_col, agg_func, fmt = _metric_options[metric_label]
    use_color = color_label != "ללא"

    if use_color:
        color_col = _axis_options[color_label]
        _grp = _cross_df.groupby([x_col, color_col], as_index=False, observed=True).agg(
            value=(metric_col, agg_func), n=("client_id", "count")
        )
        fig_cross = px.bar(
            _grp, x=x_col, y="value", color=color_col, barmode="group",
            text="value",
            title=f"{metric_label} לפי {x_label}, מפולח ב-{color_label}",
        )
    else:
        _grp = _cross_df.groupby(x_col, as_index=False, observed=True).agg(
            value=(metric_col, agg_func), n=("client_id", "count")
        )
        fig_cross = px.bar(
            _grp, x=x_col, y="value",
            text="value",
            color="value", color_continuous_scale="Teal",
            title=f"{metric_label} לפי {x_label}",
        )
        fig_cross.update_layout(coloraxis_showscale=False)

    fig_cross.update_traces(texttemplate=f"%{{text:{fmt}}}", textposition="outside")
    fig_cross.update_layout(
        xaxis_title=x_label, yaxis_title=metric_label,
        height=480,
    )
    if "%" in fmt:
        fig_cross.update_layout(yaxis_tickformat=".0%")
    st.plotly_chart(fig_cross, use_container_width=True)

    # Show data table below
    if use_color:
        _pivot = _grp.pivot_table(index=x_col, columns=color_col, values="value", observed=True)
        st.dataframe(_pivot.style.format(fmt.replace(",", "")), use_container_width=True, height=250)
    else:
        _show = _grp.rename(columns={x_col: x_label, "value": metric_label, "n": "כמות לקוחות"})
        st.dataframe(_show, use_container_width=True, hide_index=True, height=250)


def show_explore_table(filtered: pd.DataFrame) -> None:
    st.markdown(
        '<div class="colored-header" style="background:#1a1a1a;color:white;">'
        'חקירת נתונים – טבלאה אינטראקטיבית</div>',
        unsafe_allow_html=True,
    )

    exp_col1, exp_col2 = st.columns([2, 3])
    with exp_col1:
        search_text = st.text_input("🔍 חיפוש חופשי (שם, עיר, ID...)", "")
    with exp_col2:
        cols_to_show = st.multiselect(
            "עמודות לתצוגה",
            filtered.columns.tolist(),
            default=filtered.columns.tolist(),
        )

    display_df = filtered[cols_to_show] if cols_to_show else filtered
    if search_text:
        mask = display_df.astype(str).apply(lambda row: row.str.contains(search_text, case=False, na=False)).any(axis=1)
        display_df = display_df[mask]

    st.caption(f"מציג {len(display_df):,} שורות מתוך {len(filtered):,}")
    st.dataframe(display_df, use_container_width=True, height=420)


def show_export_section(df: pd.DataFrame) -> None:
    # ── Item 7 – תיאור מימוש (Part B deliverable) ──
    st.markdown(
        '<div class="colored-header" style="background:#1a1a1a;color:white;">'
        'פריט 7 – תיאור הדשבורד, הכלים ומימוש התיקון</div>',
        unsafe_allow_html=True,
    )
    d7_l, d7_r = st.columns([3, 2])
    with d7_l:
        st.markdown("""
**כלים שנבחרו:** Python עם Streamlit לממשק האינטראקטיבי, Pandas לניתוח ומניפולציה של הנתונים, Plotly Express לגרפים דינמיים, ו-Openpyxl לייצוא Excel. הבחירה ב-Streamlit מאפשרת פריסה מהירה ללא תשתית Front-End נפרדת, עם תמיכה מלאה ב-RTL לעברית.

**מה הדשבורד מציג:** תמונת מצב שלמה בזמן אמת של 2,000 לקוחות – 9 KPIs מנהלים, תובנות מהירות אוטומטיות (עיר/שירות בסיכון, לקוחות גדולים שנטשו), גרפי נטישה לאורך זמן, ניתוח שביעות רצון, פילוח גיאוגרפי ודמוגרפי, וחיפוש חופשי בטבלה. ניתן לסנן לפי עיר, סוג שירות וסטטוס.

**מימוש תיקון ה-client_id:** מיון הרשומות לפי `תאריך_הצטרפות` (Mergesort יציב) ← הקצאת מספר חדש C1000, C1001... לפי סדר כרונולוגי ← ייצוא כקובץ Excel עם Openpyxl. כל שאר עמודות הנתונים נשמרות ללא שינוי; רק `client_id` מתעדכן.
""")
    with d7_r:
        st.markdown(
            '<div style="background:#1971c215; border-right:4px solid #1971c2; padding:1rem; border-radius:8px;">'
            '<div style="font-weight:700; color:#1971c2; margin-bottom:0.6rem;">Stack טכני</div>'
            '<table style="width:100%; font-size:0.88rem;">'
            '<tr><td><b>Streamlit</b></td><td>UI + אינטראקטיביות</td></tr>'
            '<tr><td><b>Plotly</b></td><td>10 גרפים דינמיים</td></tr>'
            '<tr><td><b>Pandas</b></td><td>ניתוח וטרנספורמציה</td></tr>'
            '<tr><td><b>Openpyxl</b></td><td>ייצוא Excel</td></tr>'
            '<tr><td><b>Python 3.14</b></td><td>שפת פיתוח</td></tr>'
            '</table></div>',
            unsafe_allow_html=True,
        )

    st.divider()

    # ── Item 8 Header ──
    st.markdown(
        '<div class="colored-header" style="background:linear-gradient(90deg,#1971c2,#364fc7);color:white;">'
        'פריט 8 – תיקון סדר מספרי לקוח וייצוא Excel</div>',
        unsafe_allow_html=True,
    )

    # ── Problem explanation ──
    prob_col, arrow_col, sol_col = st.columns([5, 1, 5])
    with prob_col:
        st.markdown(
            '<div style="background:#e0313115; border-right:4px solid #e03131; padding:1rem 1.2rem; border-radius:8px;">'
            '<div style="font-weight:700; color:#e03131; font-size:1rem; margin-bottom:0.5rem;">הבעיה הקיימת</div>'
            'מספרי לקוח (<code>client_id</code>) אינם תואמים את סדר הצטרפות בפועל. לקוח עם מספר נמוך לא בהכרח הצטרף לפני לקוח עם מספר גבוה. המערכת אינה מאפשרת תיקון פנימי כרגע.'
            '</div>',
            unsafe_allow_html=True,
        )
    with arrow_col:
        st.markdown('<div style="text-align:center; font-size:2rem; padding-top:1rem;">→</div>', unsafe_allow_html=True)
    with sol_col:
        st.markdown(
            '<div style="background:#2b8a3e15; border-right:4px solid #2b8a3e; padding:1rem 1.2rem; border-radius:8px;">'
            '<div style="font-weight:700; color:#2b8a3e; font-size:1rem; margin-bottom:0.5rem;">הפתרון</div>'
            'מיון לפי <code>תאריך_הצטרפות</code> והקצאת <code>client_id</code> חדש לכל לקוח: הראשון מקבל C1000, השני C1001 וכן הלאה. כל שאר הנתונים נשמרים בדיוק.'
            '</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Before / After comparison ──
    corrected_df = build_corrected_ids(df)

    # pick 8 rows that show the ordering fix most dramatically
    sample_before = (
        df.sort_values("client_id")
        .head(8)[["client_id", "שם", "תאריך_הצטרפות", "עיר", "סוג_שירות"]]
        .rename(columns={"client_id": "מספר לקוח (לפני תיקון)"})
    )
    sample_after = (
        corrected_df.head(8)[["client_id", "שם", "תאריך_הצטרפות", "עיר", "סוג_שירות"]]
        .rename(columns={"client_id": "מספר לקוח (אחרי תיקון)"})
    )

    left_tbl, right_tbl = st.columns(2)
    with left_tbl:
        st.markdown(
            '<div style="text-align:center; font-weight:700; color:#e03131; padding:0.4rem; '
            'background:#e0313112; border-radius:6px; margin-bottom:0.4rem;">לפני תיקון – סדר שגוי</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(sample_before, use_container_width=True, hide_index=True)

    with right_tbl:
        st.markdown(
            '<div style="text-align:center; font-weight:700; color:#2b8a3e; padding:0.4rem; '
            'background:#2b8a3e12; border-radius:6px; margin-bottom:0.4rem;">אחרי תיקון – סדר נכון</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(sample_after, use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Stats strip ──
    s1, s2, s3 = st.columns(3)
    s1.metric("סה\u05f4כ רשומות", f"{len(corrected_df):,}")
    s2.metric("טווח חדש (מין)", f"C{corrected_df['client_id'].str[1:].astype(int).min():,}")
    s3.metric("טווח חדש (מקס)", f"C{corrected_df['client_id'].str[1:].astype(int).max():,}")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Download button (big and centered) ──
    excel_bytes = export_to_excel_bytes(corrected_df)
    st.markdown(
        '<div style="text-align:center;">',
        unsafe_allow_html=True,
    )
    st.download_button(
        label="⬇ הורדת קובץ Excel מתוקן",
        data=excel_bytes,
        file_name="clients_data_corrected.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.markdown('</div>', unsafe_allow_html=True)
    st.caption("הקובץ כולל את כל 2,000 הלקוחות. כל שאר הנתונים נשמרים בדיוק – רק העמודה client_id שונתה.")


def show_insights_section(filtered: pd.DataFrame, full_df: pd.DataFrame) -> None:
    """Part C – Data Analysis, Insights & Prediction Model (items 9-15)."""

    analysis_df = full_df.copy()
    analysis_df["churn"] = (analysis_df["סטטוס"] == "לא פעיל").astype(int)

    # ── 9. ממצאים מרכזיים ──────────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:linear-gradient(90deg,#1971c2,#0ca678);color:white;">'
        'סעיף 9 – ממצאים מרכזיים מהנתונים</div>',
        unsafe_allow_html=True,
    )

    churn_rate = analysis_df["churn"].mean() * 100
    sat_active = analysis_df.loc[analysis_df["סטטוס"] == "פעיל", "שביעות_רצון"].mean()
    sat_churn = analysis_df.loc[analysis_df["סטטוס"] == "לא פעיל", "שביעות_רצון"].mean()

    # Risk profile calculation
    _at_risk = analysis_df[
        (analysis_df["שביעות_רצון"] <= 4)
        & (analysis_df["זמן_תגובה_ממוצע_שעות"] >= 60)
        & (analysis_df["מספר_פניות_שנה_אחרונה"] <= 4)
    ]
    risk_rate = _at_risk["churn"].mean() * 100 if len(_at_risk) else 0
    _safe = analysis_df[
        (analysis_df["שביעות_רצון"] >= 7)
        & (analysis_df["זמן_תגובה_ממוצע_שעות"] <= 30)
        & (analysis_df["מספר_פניות_שנה_אחרונה"] >= 8)
    ]
    safe_rate = _safe["churn"].mean() * 100 if len(_safe) else 0

    # Days to churn
    _churned = analysis_df[analysis_df["churn"] == 1].copy()
    _churned["days_to_churn"] = (
        pd.to_datetime(_churned["תאריך_נטישה"], errors="coerce")
        - pd.to_datetime(_churned["תאריך_הצטרפות"], errors="coerce")
    ).dt.days

    # Best/worst city×service
    _combo = analysis_df.groupby(["עיר", "סוג_שירות"], as_index=False).agg(
        n=("churn", "count"), rate=("churn", "mean")
    )
    _combo_top = _combo[_combo["n"] >= 10].sort_values("rate", ascending=False)
    worst_combo = _combo_top.iloc[0] if len(_combo_top) else None

    # ── KPI row 1: Core numbers ──
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("שיעור נטישה כולל", f"{churn_rate:.1f}%", help="715 מתוך 2,000 לקוחות")
    k2.metric("עיר הכי בסיכון", "רמת גן", "46.7% נטישה")
    k3.metric("שירות הכי בסיכון", "ביטוח רכב", "39.6% נטישה")
    k4.metric("זמן חציוני לנטישה", f"{_churned['days_to_churn'].median():.0f} יום", help="מרגע ההצטרפות")
    k5.metric("25% נוטשים תוך", f"{_churned['days_to_churn'].quantile(0.25):.0f} יום", help="חלון התערבות מוקדם")

    # ── KPI row 2: Risk profiles ──
    k6, k7, k8, k9 = st.columns(4)
    k6.metric("פרופיל מסוכן", f"{risk_rate:.1f}%", f"+{risk_rate - churn_rate:.1f} נק׳ מהממוצע", help=f"שביעות עד 4, תגובה מעל 60 שעות, פניות עד 4 ({len(_at_risk)} לקוחות)")
    k7.metric("פרופיל בטוח", f"{safe_rate:.1f}%", f"{safe_rate - churn_rate:+.1f} נק׳ מהממוצע", help=f"שביעות 7 ומעלה, תגובה עד 30 שעות, פניות 8 ומעלה ({len(_safe)} לקוחות)")
    if worst_combo is not None:
        k8.metric("שילוב הכי מסוכן", f"{worst_combo['עיר']} + {worst_combo['סוג_שירות']}", f"{worst_combo['rate']:.0%} נטישה")
    else:
        k8.metric("שילוב הכי מסוכן", "—", "")
    k9.metric("פער שביעות רצון", f"{sat_active - sat_churn:.2f}", "כיווני אך חלש לבדו")

    # ── Charts row 1: Service churn + City churn ──
    ch_left, ch_right = st.columns(2)
    with ch_left:
        svc_churn = (
            analysis_df.groupby("סוג_שירות", as_index=False)["churn"]
            .mean()
            .sort_values("churn", ascending=False)
        )
        fig_svc = px.bar(
            svc_churn, x="סוג_שירות", y="churn",
            title="שיעור נטישה לפי סוג שירות",
            color="churn", color_continuous_scale="Reds",
            text="churn",
        )
        fig_svc.update_traces(texttemplate="%{text:.1%}", textposition="outside")
        fig_svc.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="סוג שירות", coloraxis_showscale=False,
        )
        st.plotly_chart(fig_svc, use_container_width=True)

    with ch_right:
        city_churn = (
            analysis_df.groupby("עיר", as_index=False)
            .agg(n=("churn", "count"), rate=("churn", "mean"))
        )
        city_churn = city_churn[city_churn["n"] >= 20].sort_values("rate", ascending=True).tail(10)
        fig_city = px.bar(
            city_churn, y="עיר", x="rate", orientation="h",
            title="ערים עם נטישה גבוהה (20 לקוחות ומעלה)",
            color="rate", color_continuous_scale="Reds", text="n",
        )
        fig_city.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="inside", insidetextanchor="end")
        fig_city.update_layout(
            xaxis_tickformat=".0%", xaxis_title="שיעור נטישה",
            yaxis_title="", coloraxis_showscale=False,
        )
        st.plotly_chart(fig_city, use_container_width=True)

    # ── Charts row 2: Risk profile comparison + Days to churn ──
    st.markdown(
        '<div class="colored-header" style="background:#e67700;color:white;">'
        'פרופילי סיכון ממשיים וחלון התערבות</div>',
        unsafe_allow_html=True,
    )
    ch2_left, ch2_right = st.columns(2)

    with ch2_left:
        _profiles = pd.DataFrame([
            {"פרופיל": "פרופיל מסוכן", "שיעור_נטישה": risk_rate / 100, "n": len(_at_risk)},
            {"פרופיל": "ממוצע כללי", "שיעור_נטישה": churn_rate / 100, "n": len(analysis_df)},
            {"פרופיל": "פרופיל בטוח", "שיעור_נטישה": safe_rate / 100, "n": len(_safe)},
        ])
        fig_profiles = px.bar(
            _profiles, x="פרופיל", y="שיעור_נטישה",
            title="השוואת פרופילי סיכון – שילוב 3 משתנים",
            color="שיעור_נטישה", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_profiles.update_traces(texttemplate="מתוך <b>%{text}</b> לקוחות", textposition="outside")
        fig_profiles.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="", coloraxis_showscale=False,
            yaxis_range=[0, 0.55],
        )
        st.plotly_chart(fig_profiles, use_container_width=True)
        st.caption(
            "🔴 פרופיל מסוכן: שביעות רצון נמוכה (עד 4), זמן תגובה ארוך (מעל 60 שעות), מעט פניות (עד 4)."
        )
        st.caption(
            "🟢 פרופיל בטוח: שביעות רצון גבוהה (7 ומעלה), תגובה מהירה (עד 30 שעות), פניות תכופות (8 ומעלה)."
        )
        st.caption(
            "📊 הפער: כ-16 נקודות אחוז – משתנה בודד לא מנבא, אך השילוב חושף הבדל משמעותי."
        )

    with ch2_right:
        if not _churned["days_to_churn"].dropna().empty:
            _days_data = _churned["days_to_churn"].dropna()
            _months = (_days_data / 30).round(0).astype(int)
            _month_counts = _months.value_counts().sort_index().reset_index()
            _month_counts.columns = ["חודשים", "לקוחות"]
            fig_days = px.bar(
                _month_counts, x="חודשים", y="לקוחות",
                title="התפלגות זמן עד נטישה (בחודשים)",
                color_discrete_sequence=["#e03131"],
            )
            _median_months = _days_data.median() / 30
            _q25_months = _days_data.quantile(0.25) / 30
            fig_days.add_vline(x=_median_months, line_dash="dash", line_color="#1971c2",
                               annotation_text=f"חציון: {_median_months:.0f} חודשים", annotation_position="top right")
            fig_days.add_vline(x=_q25_months, line_dash="dot", line_color="#e67700",
                               annotation_text=f"25%: {_q25_months:.0f} חודשים", annotation_position="top left")
            fig_days.update_layout(
                xaxis_title="חודשים מהצטרפות עד נטישה",
                yaxis_title="מספר לקוחות",
                xaxis_dtick=3,
                bargap=0.15,
            )
            st.plotly_chart(fig_days, use_container_width=True)
            st.caption("25% מהנוטשים עוזבים תוך 7 חודשים – זהו חלון ההתערבות הקריטי. החציון: כשנה.")

    # ── Charts row 3: Directional trends – quintile analysis ──
    st.markdown(
        '<div class="colored-header" style="background:#364fc7;color:white;">'
        'מגמות כיווניות – ניתוח חמישונים</div>',
        unsafe_allow_html=True,
    )
    ch3_left, ch3_right = st.columns(2)

    with ch3_left:
        _sat_q = analysis_df.copy()
        _sat_q["חמישון_שביעות_רצון"] = pd.qcut(
            _sat_q["שביעות_רצון"], q=5, duplicates="drop"
        )
        _sat_q_agg = (
            _sat_q.groupby("חמישון_שביעות_רצון", observed=True)
            .agg(rate=("churn", "mean"), n=("churn", "count"))
            .reset_index()
        )
        _labels = ["נמוך ביותר", "נמוך", "בינוני", "גבוה", "גבוה ביותר"]
        _sat_q_agg = _sat_q_agg.sort_values("חמישון_שביעות_רצון").reset_index(drop=True)
        _sat_q_agg["label"] = [_labels[i] if i < len(_labels) else str(i+1) for i in range(len(_sat_q_agg))]
        fig_sat_q = px.bar(
            _sat_q_agg, x="label", y="rate",
            title="שיעור נטישה לפי חמישון שביעות רצון",
            color="rate", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_sat_q.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_sat_q.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="חמישון שביעות רצון (נמוך ← גבוה)",
            coloraxis_showscale=False, yaxis_range=[0, 0.45],
        )
        st.plotly_chart(fig_sat_q, use_container_width=True)

    with ch3_right:
        _port_q = analysis_df.copy()
        _port_q["חמישון_סכום_תיק"] = pd.qcut(
            _port_q["סכום_תיק"], q=5, duplicates="drop"
        )
        _port_q_agg = (
            _port_q.groupby("חמישון_סכום_תיק", observed=True)
            .agg(rate=("churn", "mean"), n=("churn", "count"))
            .reset_index()
        )
        _labels_p = ["נמוך ביותר", "נמוך", "בינוני", "גבוה", "גבוה ביותר"]
        _port_q_agg = _port_q_agg.sort_values("חמישון_סכום_תיק").reset_index(drop=True)
        _port_q_agg["label"] = [_labels_p[i] if i < len(_labels_p) else str(i+1) for i in range(len(_port_q_agg))]
        fig_port_q = px.bar(
            _port_q_agg, x="label", y="rate",
            title="שיעור נטישה לפי חמישון סכום תיק",
            color="rate", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_port_q.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_port_q.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="חמישון סכום תיק (נמוך ← גבוה)",
            coloraxis_showscale=False, yaxis_range=[0, 0.45],
        )
        st.plotly_chart(fig_port_q, use_container_width=True)

    st.caption("מגמות כיווניות עקביות: שביעות רצון גבוהה וסכום תיק גדול קשורים לנטישה מעט נמוכה יותר (כ-6 עד 7 נקודות אחוז). לא חזק לבדו, אך תורם במודל משולב.")

    # ── Charts row 4: Age + Contacts vs churn ──
    ch4_left, ch4_right = st.columns(2)

    with ch4_left:
        _age_df = analysis_df.copy()
        _age_df["קבוצת_גיל"] = pd.cut(
            _age_df["גיל"], bins=[20, 30, 40, 50, 60, 70],
            labels=["20-30", "31-40", "41-50", "51-60", "61-70"]
        )
        _age_agg = (
            _age_df.groupby("קבוצת_גיל", observed=True)
            .agg(rate=("churn", "mean"), n=("churn", "count"))
            .reset_index()
        )
        fig_age_churn = px.bar(
            _age_agg, x="קבוצת_גיל", y="rate",
            title="שיעור נטישה לפי קבוצת גיל",
            color="rate", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_age_churn.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_age_churn.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="קבוצת גיל", coloraxis_showscale=False,
            yaxis_range=[0, 0.45],
        )
        st.plotly_chart(fig_age_churn, use_container_width=True)
        st.caption("צעירים (20-30) נוטשים ב-39.6%, לעומת 32.5% בגיל 41-50. "
                   "הסבר: לקוחות צעירים רגילים להשוות ולהחליף ספק באינטרנט, "
                   "בעוד לקוחות מבוגרים נוטים להישאר נאמנים לספק קיים. נקודת אופטימום: גיל 41-50 (32.5%).")

    with ch4_right:
        _cnt_df = analysis_df.copy()
        _cnt_df["קבוצת_פניות"] = pd.cut(
            _cnt_df["מספר_פניות_שנה_אחרונה"],
            bins=[0, 3, 5, 8, 12, 20],
            labels=["1-3", "4-5", "6-8", "9-12", "13+"]
        )
        _cnt_agg = (
            _cnt_df.groupby("קבוצת_פניות", observed=True)
            .agg(rate=("churn", "mean"), n=("churn", "count"))
            .reset_index()
        )
        fig_cnt = px.bar(
            _cnt_agg, x="קבוצת_פניות", y="rate",
            title="שיעור נטישה לפי מספר פניות שנתי",
            color="rate", color_continuous_scale="RdYlGn_r",
            text="n",
        )
        fig_cnt.update_traces(texttemplate="סה״כ %{text} לקוחות", textposition="outside")
        fig_cnt.update_layout(
            yaxis_tickformat=".0%", yaxis_title="שיעור נטישה",
            xaxis_title="מספר פניות בשנה", coloraxis_showscale=False,
            yaxis_range=[0, 0.45],
        )
        st.plotly_chart(fig_cnt, use_container_width=True)
        st.caption("לקוחות עם 13+ פניות נוטשים פחות (30.6%) – כנראה מעורבים יותר. לקוחות שקטים = סיכון.")

    # ── Written findings – comprehensive ──
    st.markdown(
        '<div style="background:#1971c215; border-right:4px solid #1971c2; padding:1.2rem 1.4rem; border-radius:8px;">'
        '<div style="font-weight:700; color:#1971c2; font-size:1.1rem; margin-bottom:0.6rem;">סיכום ממצאים מרכזיים</div>'
        #
        '<div style="font-weight:700; color:#e03131; margin-bottom:0.3rem;">ממצאים חזקים (חד-משתניים):</div>'
        '<ul style="margin:0 0 0.6rem 0;">'
        '<li><b style="color:#e03131;">רמת גן:</b> 46.7% נטישה — חריגה ב-15 נקודות אחוז מעל מודיעין (31.7%). הפער העירוני הגדול ביותר.</li>'
        '<li><b style="color:#e03131;">ביטוח רכב:</b> 39.6% נטישה — הגבוה ביותר בין סוגי השירות (פער של 8.7 נקודות אחוז מחיסכון ארוך טווח).</li>'
        '<li><b style="color:#e03131;">רמת גן + ביטוח בריאות:</b> 78.6% נטישה (14 לקוחות) — שילוב קיצוני שדורש תשומת לב מיוחדת.</li>'
        '</ul>'
        #
        '<div style="font-weight:700; color:#e67700; margin-bottom:0.3rem;">ממצאים כיווניים (שילוב משתנים):</div>'
        '<ul style="margin:0 0 0.6rem 0;">'
        '<li><b>פרופיל "מסוכן"</b> (שביעות עד 4, תגובה מעל 60 שעות, פניות עד 4): <b>45.5%</b> נטישה — פער של 10 נקודות אחוז מהממוצע.</li>'
        '<li><b>פרופיל "בטוח"</b> (שביעות 7 ומעלה, תגובה עד 30 שעות, פניות 8 ומעלה): <b>29.3%</b> נטישה — פי 1.55 נמוך מהמסוכן.</li>'
        '<li><b>צעירים (20-30):</b> 39.6% נטישה — גבוה ב-7 נקודות אחוז מגיל 41-50. לקוחות צעירים רגילים להשוות ולהחליף ספק באינטרנט, בעוד מבוגרים נאמנים יותר לספק קיים.</li>'
        '<li><b>לקוחות עם 13 פניות ומעלה:</b> 30.6% נטישה — הנמוך ביותר. לקוח שקט = לקוח בסיכון.</li>'
        '<li><b>סכום תיק עליון:</b> 32.5% לעומת 39.5% לתחתון — פער של כ-7 נקודות אחוז.</li>'
        '</ul>'
        #
        '<div style="font-weight:700; color:#0ca678; margin-bottom:0.3rem;">ממצאים תפעוליים:</div>'
        '<ul style="margin:0 0 0.6rem 0;">'
        '<li><b>חלון התערבות:</b> 25% מהנוטשים עוזבים תוך 220 יום (7 חודשים). החציון: 382 יום (~שנה).</li>'
        '<li><b>טווח נטישה:</b> 53 עד 699 יום – רוב הנטישה מתרכזת בשנה הראשונה.</li>'
        '<li><b>2025 נטישה נמוכה (28.6%):</b> כנראה אפקט צנזור – לקוחות שהצטרפו ב-2025 טרם הספיקו לנטוש.</li>'
        '</ul>'
        #
        '<div style="margin-top:0.7rem; padding:0.7rem; background:#e6770015; border-radius:6px; font-weight:600; color:#e67700;">'
        'מסקנת-על: אף משתנה נומרי בודד לא מנבא נטישה (כל |r| &lt; 0.06). '
        'אך <b>שילוב משתנים</b> חושף פערי נטישה של 10 עד 16 נקודות אחוז, ו<b>אינטראקציות קטגוריאליות</b> (עיר ושירות) מגיעות ל-78.6%. '
        'לכן נדרש <b>מודל רב-משתני כמו XGBoost</b> שיזהה דפוסים מורכבים שמשתנה בודד לא חושף.'
        '</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 10. הצעה מבוססת נתונים ─────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#0ca678;color:white;">'
        'סעיף 10 – הצעה מבוססת נתונים לצמצום נטישה</div>',
        unsafe_allow_html=True,
    )

    prop_l, prop_r = st.columns([3, 2])
    with prop_l:
        st.markdown("""
**הצעה: "לקוח על הגדר" – זיהוי מוקדם של לקוחות בסיכון נטישה**

הניתוח חשף שני דפוסים מרכזיים שמאפשרים לפעול לפני שהלקוח עוזב:
- **חלון הזדמנות של 7 חודשים:** רבע מהנוטשים עוזבים בתוך 220 יום מרגע ההצטרפות. זה פרק הזמן שבו ניתן עדיין להציל אותם.
- **דפוס סיכון ברור:** לקוח לא מרוצה (שביעות רצון עד 4), שמחכה הרבה זמן למענה (מעל 60 שעות), ולא פונה הרבה לחברה (עד 4 פניות) – מגיע ל-45.5% סיכון נטישה, לעומת 35.8% בממוצע.

איך זה יעבוד בפועל?

1. **ניקוד שבועי לכל לקוח** – המערכת תחשב ציון סיכון על בסיס שביעות רצון, תדירות פניות, זמן תגובה, סוג שירות ועיר מגורים (למשל, לקוח מרמת גן מקבל +10 נקודות סיכון).
2. **התראה ב-7 החודשים הראשונים** – לקוח שעובר את סף הסיכון מחובר ליועץ ייעודי תוך 4 שעות.
3. **טיפול מותאם אישית** – לקוח מרמת גן עם ביטוח בריאות (78.6% נטישה!) מטופל מיידית. לקוח צעיר (20-30) עם מעט פניות מקבל יצירת קשר יזומה.
4. **מדידת אפקטיביות** – השוואת עלות שימור מול גיוס (יחס 1:5), מעקב חודשי אחר שינוי בפרופילי הסיכון.

**אימפקט צפוי:** ירידה של 20-30% בנטישה → שימור כ-150 לקוחות בשנה → 
שווי הכנסה שנתית שנשמרת: כ-6 מיליון ₪.
""")

    with prop_r:
        # Funnel visualization
        fig_funnel = go.Figure(go.Funnel(
            y=["כל הלקוחות (2,000)", "בסיכון (ציון 60 ומעלה)", "קיבלו התראה", "שימור מוצלח"],
            x=[2000, 600, 400, 300],
            textinfo="value+percent initial",
            marker_color=["#1971c2", "#e67700", "#e03131", "#2b8a3e"],
        ))
        fig_funnel.update_layout(
            title="משפך שימור צפוי",
            margin={"l": 10, "r": 10, "t": 50, "b": 10},
            height=350,
        )
        st.plotly_chart(fig_funnel, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 11. בחירת מודל חיזוי ───────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#7048e8;color:white;">'
        'סעיף 11 – מודל חיזוי נטישה (לא מבוסס רגרסיה)</div>',
        unsafe_allow_html=True,
    )

    m1, m2, m3 = st.columns(3)
    _model_cards = [
        ("XGBoost / LightGBM", "#2b8a3e", "מומלץ ראשי",
         "Gradient Boosted Trees – מוביל בתחרויות Kaggle על דאטה טבלאי. "
         "תומך בקשרים לא-ליניאריים, מטפל בחסרים באופן מובנה, "
         "ומאפשר הסבר דרך SHAP ו-Feature Importance. "
         "מתאים מצוין ל-2,000 שורות עם 10 פיצ'רים."),
        ("KNN (K-Nearest Neighbors)", "#1971c2", "אלטרנטיבה פשוטה",
         "מסווג לפי דמיון ללקוחות הקרובים ביותר. "
         "קל להסבר למנכ\"ל: \"הלקוח דומה ל-5 לקוחות שנטשו\". "
         "דורש נרמול פיצ'רים ורגיש לממדים גבוהים, "
         "אבל אפקטיבי על דאטה סט קטן-בינוני כמו שלנו."),
        ("Random Forest", "#e67700", "מודל Ensemble חזק",
         "יער עצי החלטה עצמאיים שמצביעים ברוב. "
         "יציב, עמיד לרעש ולא דורש כיוון רב. "
         "אפשר לחלץ Feature Importance ישירות. "
         "פחות חד מ-XGBoost אבל פחות Overfitting."),
    ]
    for col, (name, color, badge, desc) in zip([m1, m2, m3], _model_cards):
        col.markdown(
            f'<div style="background:{color}12; border-right:4px solid {color}; '
            f'padding:1rem; border-radius:8px; min-height:220px;">'
            f'<span style="background:{color}; color:white; padding:2px 8px; border-radius:4px; '
            f'font-size:0.75rem;">{badge}</span>'
            f'<div style="font-weight:700; color:{color}; font-size:1.05rem; margin:0.5rem 0;">{name}</div>'
            f'<div style="font-size:0.88rem;">{desc}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("""
> **נימוק הבחירה ב-XGBoost, מבוסס על ניתוח הנתונים:**
> 
> 1. **אף משתנה בודד לא מספיק** – כל הקורלציות הנומריות קרובות לאפס (|r| < 0.06).
> 2. **שילוב משתנים עובד** – פרופיל מסוכן (3 משתנים) מגיע ל-45.5% נטישה, פי 1.55 מהפרופיל הבטוח.
> 3. **אינטראקציות קטגוריאליות קריטיות** – רמת גן + ביטוח בריאות = 78.6% נטישה.
> 4. **XGBoost** מזהה אינטראקציות אלו **אוטומטית** דרך עצי החלטה רצפיים, ו-**SHAP** מסביר למנכ״ל *למה* כל לקוח סומן כמועמד לנטישה.
> 5. **חשוב:** לקוח "שקט" (מעט פניות) + צעיר + רמת גן = פרופיל שמודל ליניארי יחמיץ, אך XGBoost ילכוד.
""")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 12. נתוני מאקרו ────────────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#e67700;color:white;">'
        'סעיף 12 – נתוני מאקרו שיש לבדוק</div>',
        unsafe_allow_html=True,
    )

    macro_data = [
        ("ריבית בנק ישראל", "#e03131",
         "שינוי ריבית משפיע ישירות על אטרקטיביות מוצרי השקעה ופנסיה. "
         "עלייה בריבית → לקוחות בוחנים אלטרנטיבות → סיכון נטישה עולה."),
        ("מדד המחירים לצרכן (CPI)", "#1971c2",
         "אינפלציה גבוהה מכרסמת בהכנסה הפנויה → לקוחות מקטינים תיק או עוזבים שירותים \"לא הכרחיים\"."),
        ("שיעור אבטלה מקומי", "#0ca678",
         "אבטלה לפי אזור גיאוגרפי (רמת גן!) – קורלציה ישירה ליכולת תשלום ולהישארות."),
        ("מדד S&P 500 / ת\"א 125", "#7048e8",
         "ביצועי שוק ההון משפיעים על שביעות רצון מייעוץ השקעות ועל גודל התיק."),
        ("עונתיות + אירועים", "#e67700",
         "חודשים לפני פסח/ספטמבר – הוצאות גדולות. רגולציה חדשה (IFRS, SOX) יכולה לשנות התנהגות."),
    ]

    for title, color, text in macro_data:
        st.markdown(
            f'<div style="background:{color}10; border-right:4px solid {color}; '
            f'padding:0.7rem 1rem; border-radius:6px; margin-bottom:0.5rem;">'
            f'<span style="font-weight:700; color:{color};">{title}</span> – {text}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("""
> **ההיגיון:** מודל שרואה רק נתונים פנימיים (שביעות רצון, זמן תגובה) יחמיץ 
> שינויים מאקרו-כלכליים שמשפיעים על כל הלקוחות בבת אחת. שילוב נתוני מאקרו 
> כפיצ'רים חיצוניים (**ARIMAX-style** או כ-features ב-XGBoost) מעלה את הדיוק 
> ומאפשר **חיזוי פרואקטיבי** – לפני שהלקוח מרגיש את השינוי.
""")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 13. חלוקת אימון/בדיקה ──────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#e03131;color:white;">'
        'סעיף 13 – חלוקת נתוני אימון ובדיקה</div>',
        unsafe_allow_html=True,
    )

    split_l, split_r = st.columns([3, 2])
    with split_l:
        st.markdown("""
**אסטרטגיית חלוקה: Stratified K-Fold + Holdout**

| פרמטר | ערך | נימוק |
|-------|-----|------|
| **חלוקה ראשית (Holdout)** | 80% אימון / 20% בדיקה | מספיק דאטה לאימון (1,600) ובדיקה (400) |
| **שכבות (Stratification)** | לפי עמודת `סטטוס` | שומר על יחס 35.8% נטישה בכל חלוקה |
| **אימות צולב (Cross-Validation)** | 5 קפלים שכבתיים | הערכה יציבה ומניעת התאמת-יתר |
| **מצב אקראי (Random State)** | קבוע (42) | שחזור תוצאות |

**למה Stratified?** שיעור הנטישה (35.8%) הוא לא מאוזן. חלוקה רנדומית עלולה 
לייצר Fold עם 30% או 42% – ולהטות את המודל. Stratification מבטיח שכל Fold 
מייצג את האוכלוסייה.
""")

    with split_r:
        fig_split = px.pie(
            pd.DataFrame({"קבוצה": ["אימון (80%)", "בדיקה (20%)"], "n": [1600, 400]}),
            names="קבוצה", values="n", hole=0.5,
            color="קבוצה",
            color_discrete_map={"אימון (80%)": "#1971c2", "בדיקה (20%)": "#e03131"},
        )
        fig_split.update_traces(textinfo="label+value", textposition="outside")
        fig_split.update_layout(
            title="חלוקת אימון / בדיקה",
            margin={"l": 10, "r": 10, "t": 50, "b": 10},
            height=300,
        )
        st.plotly_chart(fig_split, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 14. השפעת קורלציה הצטרפות-תיק ──────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#364fc7;color:white;">'
        'סעיף 14 – קשר בין תאריך הצטרפות לסכום תיק</div>',
        unsafe_allow_html=True,
    )

    corr_l, corr_r = st.columns([3, 2])
    with corr_l:
        st.markdown("""
**הנתון:** ככל שתאריך ההצטרפות מאוחר יותר – סכום התיק גדול יותר (קשר חזק).

**ההשפעה על סעיפים 11-13:**

| נושא | השפעה | פעולה |
|------|-------|-------|
| **מודל (11)** | `תאריך_הצטרפות` ו-`סכום_תיק` עברו **Multicollinearity** – המודל עלול לספור את אותו אות פעמיים | **פתרון:** הסרת אחד מהפיצ'רים, או חישוב VIF ובחירת הטוב מביניהם |
| **מאקרו (12)** | הקורלציה עלולה לשקף **אינפלציה / עלייה בשווי נכסים** ולא באמת שלקוחות חדשים עשירים יותר | **פתרון:** נרמול סכום תיק למדד תאריך (Real vs. Nominal) |
| **חלוקת נתונים (13)** | חלוקה רנדומית תערבב לקוחות ותיקים (תיק קטן) עם חדשים (תיק גדול) → **Data Leakage טמפורלי** | **פתרון:** חלוקה **כרונולוגית** – אימון על 80% הראשונים, בדיקה על 20% האחרונים |

> **המסקנה המרכזית:** הקורלציה הזו מחייבת מעבר מ-Random Split ל-**Time-Based Split**, 
> ובדיקת **Collinearity** לפני הכנסת שני הפיצ'רים יחד למודל.
""")

    with corr_r:
        # Scatter plot showing the correlation
        _scatter_df = analysis_df.dropna(subset=["תאריך_הצטרפות", "סכום_תיק"]).copy()
        _scatter_df["רבעון_הצטרפות"] = _scatter_df["תאריך_הצטרפות"].dt.to_period("Q").astype(str)
        _agg = _scatter_df.groupby("רבעון_הצטרפות", as_index=False)["סכום_תיק"].mean()
        _agg = _agg.sort_values("רבעון_הצטרפות").reset_index(drop=True)
        fig_corr = px.bar(
            _agg, x="רבעון_הצטרפות", y="סכום_תיק",
            title="סכום תיק ממוצע לפי רבעון הצטרפות",
            color="סכום_תיק", color_continuous_scale="Blues",
        )
        fig_corr.update_layout(
            xaxis_title="רבעון הצטרפות", yaxis_title="סכום תיק ממוצע (₪)",
            xaxis_tickangle=-45, height=350, coloraxis_showscale=False,
        )
        st.plotly_chart(fig_corr, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 15. כלים ────────────────────────────────────────────────────
    st.markdown(
        '<div class="colored-header" style="background:#1a1a1a;color:white;">'
        'סעיף 15 – כלים בהם נעשה שימוש</div>',
        unsafe_allow_html=True,
    )

    _tools = [
        ("Python 3.14", "שפת הפיתוח הראשית", "#339af0"),
        ("Pandas", "ניתוח וטרנספורמציית נתונים", "#1971c2"),
        ("Plotly Express", "ויזואליזציה אינטראקטיבית", "#7048e8"),
        ("Streamlit", "בניית דשבורד אינטראקטיבי", "#e03131"),
        ("Openpyxl", "קריאה וכתיבת קבצי Excel", "#e67700"),
        ("GitHub Copilot", "סיוע AI בפיתוח ובכתיבה", "#0ca678"),
        ("Claude Opus 4.6", "מודל שפה לניתוח וכתיבה", "#d97706"),
        ("VS Code", "סביבת הפיתוח", "#0078d4"),
    ]
    _tools_html = '<div style="display:flex; flex-wrap:wrap; gap:10px; direction:rtl; margin:0.5rem 0;">'
    for name, desc, color in _tools:
        _tools_html += (
            f'<div style="background:{color}15; border-right:3px solid {color}; padding:10px 14px; '
            f'border-radius:8px; flex:1; min-width:150px;">'
            f'<div style="font-weight:700; color:{color};">{name}</div>'
            f'<div style="font-size:0.82rem;">{desc}</div></div>'
        )
    _tools_html += '</div>'
    st.markdown(_tools_html, unsafe_allow_html=True)


def _inject_rtl_css() -> None:
    """Force RTL layout for Hebrew and hide Streamlit deploy button."""
    st.markdown("""
    <style>
    /* Customizing Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 60px;
        white-space: pre-wrap;
        background-color: var(--secondary-background-color);
        border-radius: 8px;
        padding-top: 10px;
        padding-bottom: 10px;
        padding-left: 20px;
        padding-right: 20px;
        font-size: 1.15rem !important;
        font-weight: 700 !important;
        border: 1px solid rgba(128, 128, 128, 0.2);
        direction: rtl;
        text-align: center;
        color: var(--text-color);
    }
    .stTabs [aria-selected="true"] {
        background-color: rgba(255, 75, 75, 0.1) !important;
        border: 2px solid #ff4b4b !important;
        color: #ff4b4b !important;
    }
    /* RTL for all text */
    .stApp, .stMarkdown, .stMarkdown p, .stMarkdown li,
    .stMarkdown td, .stMarkdown th, .stMarkdown h1, .stMarkdown h2,
    .stMarkdown h3, .stMarkdown h4, .stAlert p,
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"],
    .stCaption {
        direction: rtl;
        text-align: right;
    }
    /* RTL for tables */
    .stMarkdown table { direction: rtl; }
    .stMarkdown th, .stMarkdown td { text-align: right !important; }
    /* RTL for lists */
    .stMarkdown ul, .stMarkdown ol {
        direction: rtl;
        text-align: right;
        padding-right: 1.5em;
        padding-left: 0;
    }
    /* Hide deploy button */
    [data-testid="stToolbar"] { display: none !important; }
    .stDeployButton { display: none !important; }
    header[data-testid="stHeader"] button[kind="header"] { display: none !important; }
    /* Hide sidebar completely */
    [data-testid="stSidebar"],
    [data-testid="collapsedControl"],
    button[kind="headerNoPadding"] {
        display: none !important;
        width: 0 !important;
        min-width: 0 !important;
        max-width: 0 !important;
        overflow: hidden !important;
        border: none !important;
        box-shadow: none !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        display: none !important;
    }
    section[data-testid="stSidebar"] {
        display: none !important;
    }
    /* Remove vertical column dividers */
    [data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
        border: none !important;
    }
    [data-testid="column"] { border: none !important; }
    div[data-testid="stHorizontalBlock"] > div::after,
    div[data-testid="stHorizontalBlock"] > div::before { display: none !important; }
    /* Colored subheaders */
    .colored-header { padding: 0.55rem 1rem; border-radius: 6px; margin: 1rem 0 0.5rem 0; text-align: center; font-weight: 700; font-size: 1.05rem; letter-spacing: 0.02em; }
    </style>
    """, unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="מערכת ניהול AI - מגדל שירותים פיננסיים", layout="wide")
    _inject_rtl_css()
    st.markdown('<h1 style="text-align:right; direction:rtl;">מערכת ויזואלית משולבת - מגדל שירותים פיננסיים</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:right; direction:rtl; font-size:1.1rem; color:#666;">מקום אחד שמרכז תהליך, דשבורד ניהולי, תובנות נטישה וייצוא נתונים מתוקנים.</p>', unsafe_allow_html=True)

    try:
        df = load_data()
    except Exception as exc:
        st.error(f"שגיאה בטעינת הנתונים: {exc}")
        st.stop()

    filtered = add_filters(df)

    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 חלק א' - אפיון תהליך",
        "📊 חלק ב' - דשבורד מנהלים ונתונים",
        "💡 חלק ג' - תובנות וחיזוי מתקדם",
        "💾 ייצוא ותיקון נתונים (סיום חלק ב')",
    ])

    with tab1:
        show_part_a()

    with tab2:
        show_kpis(filtered)
        # compact export button at top of dashboard
        _corrected_quick = build_corrected_ids(df)
        _excel_quick = export_to_excel_bytes(_corrected_quick)
        st.download_button(
            label="⬇ הורדת Excel עם מספרי לקוח מתוקנים",
            data=_excel_quick,
            file_name="clients_data_corrected.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        show_visuals(filtered)
        show_cross_tab(filtered)
        show_explore_table(filtered)

    with tab3:
        show_insights_section(filtered, df)

    with tab4:
        show_export_section(df)


if __name__ == "__main__":
    main()
