import os
import pandas as pd # type: ignore
import dash # type: ignore
from dash import dcc, html, Input, Output, State, dash_table, clientside_callback# type: ignore
import plotly.graph_objects as go# type: ignore
import plotly.express as px# type: ignore
import io
import base64
import subprocess
import platform
from threading import Timer
from datetime import datetime, timedelta
from groq import Groq# type: ignore
from dotenv import load_dotenv# type: ignore

load_dotenv()

groq_api = os.getenv("GROQ_API")
groq_client = Groq(api_key=groq_api)

# ============================================================
# Data Extraction
# ============================================================

def safe_float(val):
    try:
        v = float(val)
        if pd.isna(v):
            return 0.0
        return v
    except (ValueError, TypeError):
        return 0.0


def safe_str(val):
    s = str(val).strip()
    if s.lower() in ["nan", "nat", "none", ""]:
        return ""
    return s


def hex_to_rgba(hex_color, alpha=1.0):
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def extract_projects_from_sheet(df):
    col_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if "project name" in cl:
            col_map["project_name"] = c
        elif "current stage" in cl:
            col_map["current_stage"] = c
        elif "opening" in cl and "date" in cl:
            col_map["opening_date"] = c
        elif "proposed budget" in cl:
            col_map["proposed_budget"] = c
        elif "client budget" in cl:
            col_map["client_budget"] = c
        elif "orders placed" in cl and "value" in cl:
            col_map["orders_placed"] = c
        elif "orders in progress" in cl:
            col_map["orders_in_progress"] = c
        elif "currency" in cl:
            col_map["currency"] = c
        elif "proc" in cl and "process" in cl and "started" in cl:
            col_map["proc_started"] = c
        elif "total no" in cl and "package" in cl:
            col_map["total_packages"] = c
        elif "ordering completed" in cl:
            col_map["packages_completed"] = c
        elif "ordering in progress" in cl:
            col_map["packages_in_progress"] = c
        elif "delivery process" in cl and "started" in cl:
            col_map["delivery_started"] = c
        elif "total no" in cl and "po" in cl and "raised" in cl:
            col_map["total_pos"] = c
        elif "total delivered" in cl:
            col_map["delivered_pos"] = c
        elif "concern" in cl:
            col_map["concerns"] = c
        elif "overall" in cl and "procurement" in cl:
            col_map["overall_proc"] = c

    project_indices = []
    for i in range(len(df)):
        pn = safe_str(df.iloc[i].get(col_map.get("project_name", ""), ""))
        if pn:
            project_indices.append(i)

    projects = []
    for idx, pi in enumerate(project_indices):
        row = df.iloc[pi]
        next_pi = project_indices[idx + 1] if idx + 1 < len(project_indices) else len(df)

        data = {}
        data["project_name"] = safe_str(row.get(col_map.get("project_name", ""), ""))
        data["current_stage"] = safe_str(row.get(col_map.get("current_stage", ""), ""))

        od = safe_str(row.get(col_map.get("opening_date", ""), ""))
        if od:
            if od.lower() == "opened":
                data["opening_date"] = "Opened"
                data["opening_date_parsed"] = None
            else:
                try:
                    parsed = pd.to_datetime(od, dayfirst=True)
                    data["opening_date"] = parsed.strftime("%Y-%m-%d")
                    data["opening_date_parsed"] = parsed
                except Exception:
                    data["opening_date"] = od
                    data["opening_date_parsed"] = None
        else:
            data["opening_date"] = "TBD"
            data["opening_date_parsed"] = None

        data["proposed_budget"] = safe_float(row.get(col_map.get("proposed_budget", ""), 0))
        data["client_budget"] = safe_float(row.get(col_map.get("client_budget", ""), 0))
        data["orders_placed"] = safe_float(row.get(col_map.get("orders_placed", ""), 0))
        data["orders_in_progress"] = safe_float(row.get(col_map.get("orders_in_progress", ""), 0))
        data["currency"] = safe_str(row.get(col_map.get("currency", ""), "USD")) or "USD"
        data["proc_started"] = safe_str(row.get(col_map.get("proc_started", ""), ""))
        data["total_packages"] = safe_float(row.get(col_map.get("total_packages", ""), 0))
        data["packages_completed"] = safe_float(row.get(col_map.get("packages_completed", ""), 0))
        data["packages_in_progress"] = safe_float(row.get(col_map.get("packages_in_progress", ""), 0))
        data["packages_to_start"] = max(
            0,
            data["total_packages"] - data["packages_completed"] - data["packages_in_progress"],
        )
        data["delivery_started"] = safe_str(row.get(col_map.get("delivery_started", ""), ""))
        data["total_pos"] = safe_float(row.get(col_map.get("total_pos", ""), 0))
        data["delivered_pos"] = safe_float(row.get(col_map.get("delivered_pos", ""), 0))
        data["delivery_in_progress"] = max(0, data["total_pos"] - data["delivered_pos"])

        overall_proc_val = safe_float(row.get(col_map.get("overall_proc", ""), 0))
        data["overall_proc_from_file"] = (
            overall_proc_val * 100 if overall_proc_val <= 1 else overall_proc_val
        )

        budget = data["client_budget"] if data["client_budget"] > 0 else data["proposed_budget"]
        data["effective_budget"] = budget
        data["orders_placed_pct"] = (data["orders_placed"] / budget * 100) if budget > 0 else 0

        pkg_completion = (
            (data["packages_completed"] / data["total_packages"] * 100)
            if data["total_packages"] > 0
            else 0
        )
        delivery_completion = (
            (data["delivered_pos"] / data["total_pos"] * 100) if data["total_pos"] > 0 else 0
        )

        data["overall_completion"] = data["overall_proc_from_file"]
        data["pkg_completion_pct"] = pkg_completion
        data["delivery_completion_pct"] = delivery_completion

        if budget > 0 and data["orders_placed"] > 0:
            data["savings_overrun"] = budget - data["orders_placed"] - data["orders_in_progress"]
        else:
            data["savings_overrun"] = 0

        # Budget remaining (not yet committed)
        data["budget_uncommitted"] = max(
            0, budget - data["orders_placed"] - data["orders_in_progress"]
        )

        # Risk score (0-100, higher = more risk)
        risk = 0
        if data["opening_date_parsed"]:
            days_left = (data["opening_date_parsed"] - pd.Timestamp.now()).days
            if days_left < 0:
                risk += 10
            elif days_left < 90:
                risk += 40
            elif days_left < 180:
                risk += 20
        if data["overall_completion"] < 30 and data["opening_date"] != "TBD":
            risk += 30
        if data["savings_overrun"] < 0:
            risk += 20
        if data["total_packages"] > 0 and data["packages_to_start"] / data["total_packages"] > 0.3:
            risk += 10
        data["risk_score"] = min(100, risk)

        concerns = []
        concern_col = col_map.get("concerns", "")
        if concern_col:
            c = safe_str(row.get(concern_col, ""))
            if c:
                concerns.append(c)
            for j in range(pi + 1, next_pi):
                c = safe_str(df.iloc[j].get(concern_col, ""))
                if c:
                    concerns.append(c)
        data["concerns"] = concerns
        projects.append(data)

    return projects


def parse_file(contents, filename):
    content_type, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    buf = io.BytesIO(decoded)
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""

    all_projects = []
    if ext == "csv":
        df = pd.read_csv(buf)
        all_projects.extend(extract_projects_from_sheet(df))
    elif ext in ["xlsx", "xlsm", "xls"]:
        engine = "openpyxl" if ext in ["xlsx", "xlsm"] else "xlrd"
        xls = pd.ExcelFile(buf, engine=engine)
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.dropna(how="all").empty:
                all_projects.extend(extract_projects_from_sheet(df))
    return all_projects


# ============================================================
# Browser
# ============================================================

def open_browser(url):
    try:
        system = platform.system().lower()
        if system == "linux":
            for b in ["xdg-open", "google-chrome", "firefox", "chromium-browser"]:
                try:
                    subprocess.Popen(
                        [b, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                    )
                    return
                except FileNotFoundError:
                    continue
        elif system == "darwin":
            subprocess.Popen(["open", url])
        elif system == "windows":
            subprocess.Popen(["start", url], shell=True)
    except Exception:
        pass


# ============================================================
# Theme
# ============================================================

THEME = {
    "bg": "#f0f4f8",
    "card_bg": "#ffffff",
    "dark": "#0f172a",
    "dark2": "#1e293b",
    "text": "#1e293b",
    "text_light": "#64748b",
    "text_muted": "#94a3b8",
    "border": "#e2e8f0",
    "success": "#10b981",
    "warning": "#f59e0b",
    "danger": "#ef4444",
    "info": "#3b82f6",
    "purple": "#8b5cf6",
}

PROJECT_COLORS = [
    {
        "main": "#6366f1",
        "light": "#eef2ff",
        "gradient": "linear-gradient(135deg, #6366f1, #8b5cf6)",
    },
    {
        "main": "#0ea5e9",
        "light": "#f0f9ff",
        "gradient": "linear-gradient(135deg, #0ea5e9, #06b6d4)",
    },
    {
        "main": "#f97316",
        "light": "#fff7ed",
        "gradient": "linear-gradient(135deg, #f97316, #fb923c)",
    },
    {
        "main": "#ec4899",
        "light": "#fdf2f8",
        "gradient": "linear-gradient(135deg, #ec4899, #f472b6)",
    },
    {
        "main": "#14b8a6",
        "light": "#f0fdfa",
        "gradient": "linear-gradient(135deg, #14b8a6, #2dd4bf)",
    },
]

F = {
    "xs": "11px",
    "sm": "13px",
    "md": "15px",
    "lg": "18px",
    "xl": "22px",
    "xxl": "28px",
    "xxxl": "36px",
    "label": "11px",
    "value": "15px",
    "section": "13px",
    "chart_text": "11px",
}


def fmt_num(val, currency=""):
    prefix = f"{currency} " if currency else ""
    if abs(val) >= 1_000_000:
        return f"{prefix}{val / 1_000_000:,.2f}M"
    elif abs(val) >= 1_000:
        return f"{prefix}{val / 1_000:,.1f}K"
    else:
        return f"{prefix}{val:,.0f}"


def generate_ai_summary(project):
    try:
        data_text = f"""
        Project Name: {project.get("project_name")}
        Stage: {project.get("current_stage")}
        Opening: {project.get("opening_date")}
        Budget: {project.get("client_budget") or project.get("proposed_budget")}
        Orders Placed: {project.get("orders_placed")}
        Orders In Progress: {project.get("orders_in_progress")}
        Packages Completed: {project.get("packages_completed")} / {project.get("total_packages")}
        Deliveries: {project.get("delivered_pos")} / {project.get("total_pos")}
        Overall Completion: {project.get("overall_completion")}%
        Concerns: {'; '.join(project.get("concerns", [])) or 'None'}
        """

        completion = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "user",
                    "content": f"""Give a very short professional summary (max 3 sentences)
                    about this project's procurement progress, risks, and next steps.
                    {data_text}""",
                }
            ],
            temperature=0.3,
        )

        return completion.choices[0].message.content.strip()

    except Exception:
        return "AI summary unavailable"


def status_color(pct):
    if pct >= 75:
        return THEME["success"]
    elif pct >= 40:
        return THEME["warning"]
    return THEME["danger"]


def risk_color(score):
    if score >= 60:
        return THEME["danger"]
    elif score >= 30:
        return THEME["warning"]
    return THEME["success"]


def risk_label(score):
    if score >= 60:
        return "High Risk"
    elif score >= 30:
        return "Medium Risk"
    return "Low Risk"


def glass_style(extra=None):
    base = {
        "backgroundColor": "rgba(255, 255, 255, 0.82)",
        "backdropFilter": "blur(20px)",
        "WebkitBackdropFilter": "blur(20px)",
        "borderRadius": "18px",
        "border": "1px solid rgba(255, 255, 255, 0.35)",
        "boxShadow": "0 8px 32px rgba(0, 0, 0, 0.08)",
    }
    if extra:
        base.update(extra)
    return base


# ============================================================
# Portfolio Summary Header (Enhanced)
# ============================================================

def make_portfolio_header(all_data):
    n = len(all_data)
    active = sum(1 for d in all_data if d.get("overall_completion", 0) > 0)
    planning = n - active

    # Aggregate KPIs
    total_budget = sum(d.get("effective_budget", 0) for d in all_data)
    total_ordered = sum(d.get("orders_placed", 0) for d in all_data)
    total_packages = sum(d.get("total_packages", 0) for d in all_data)
    total_pkg_done = sum(d.get("packages_completed", 0) for d in all_data)
    total_pos = sum(d.get("total_pos", 0) for d in all_data)
    total_delivered = sum(d.get("delivered_pos", 0) for d in all_data)
    avg_completion = (
        sum(d.get("overall_completion", 0) for d in all_data) / n if n > 0 else 0
    )

    def header_kpi(label, value, icon):
        return html.Div(
            [
                html.Span(icon, style={"fontSize": "24px"}),
                html.Div(
                    [
                        html.P(
                            value,
                            style={
                                "fontSize": F["xl"],
                                "fontWeight": "800",
                                "color": "#fff",
                                "margin": "0",
                            },
                        ),
                        html.P(
                            label,
                            style={
                                "fontSize": F["xs"],
                                "color": "rgba(255,255,255,0.5)",
                                "margin": "2px 0 0 0",
                                "textTransform": "uppercase",
                                "letterSpacing": "1px",
                            },
                        ),
                    ]
                ),
            ],
            style={
                "display": "flex",
                "alignItems": "center",
                "gap": "12px",
                "backgroundColor": "rgba(255,255,255,0.07)",
                "padding": "16px 22px",
                "borderRadius": "14px",
                "border": "1px solid rgba(255,255,255,0.08)",
                "flex": "1",
                "minWidth": "180px",
            },
        )

    project_pills = []
    for i, d in enumerate(all_data):
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        overall = d.get("overall_completion", 0)
        sc = status_color(overall)
        rc = risk_color(d.get("risk_score", 0))

        project_pills.append(
            html.Div(
                [
                    html.Div(
                        style={
                            "width": "10px",
                            "height": "10px",
                            "borderRadius": "50%",
                            "backgroundColor": pc["main"],
                            "flexShrink": "0",
                            "boxShadow": f"0 0 8px {hex_to_rgba(pc['main'], 0.5)}",
                        }
                    ),
                    html.Div(
                        [
                            html.Span(
                                d.get("project_name", ""),
                                style={
                                    "fontSize": F["sm"],
                                    "fontWeight": "600",
                                    "color": "#e2e8f0",
                                },
                            ),
                            html.Div(
                                [
                                    html.Span(
                                        f"{overall:.0f}%",
                                        style={
                                            "fontSize": F["xs"],
                                            "fontWeight": "800",
                                            "color": sc,
                                            "backgroundColor": hex_to_rgba(sc, 0.18),
                                            "padding": "2px 8px",
                                            "borderRadius": "8px",
                                        },
                                    ),
                                    html.Span(
                                        risk_label(d.get("risk_score", 0)),
                                        style={
                                            "fontSize": "10px",
                                            "fontWeight": "700",
                                            "color": rc,
                                            "backgroundColor": hex_to_rgba(rc, 0.18),
                                            "padding": "2px 8px",
                                            "borderRadius": "8px",
                                        },
                                    ),
                                ],
                                style={"display": "flex", "gap": "6px", "marginTop": "4px"},
                            ),
                        ],
                        style={"flex": "1"},
                    ),
                ],
                style={
                    "display": "flex",
                    "alignItems": "center",
                    "gap": "10px",
                    "backgroundColor": "rgba(255,255,255,0.07)",
                    "padding": "12px 18px",
                    "borderRadius": "12px",
                    "border": "1px solid rgba(255,255,255,0.08)",
                    "backdropFilter": "blur(10px)",
                    "minWidth": "220px",
                },
            )
        )

    return html.Div(
        [
            # Title row
            html.Div(
                [
                    html.Div(
                        [
                            html.H2(
                                "Portfolio Overview",
                                style={
                                    "margin": "0",
                                    "color": "#fff",
                                    "fontSize": F["xxl"],
                                    "fontWeight": "800",
                                    "letterSpacing": "-0.5px",
                                },
                            ),
                            html.P(
                                f"{n} Projects  •  {active} Active  •  {planning} Planning",
                                style={
                                    "margin": "6px 0 0 0",
                                    "color": "rgba(255,255,255,0.5)",
                                    "fontSize": F["sm"],
                                    "fontWeight": "500",
                                },
                            ),
                        ]
                    ),
                    html.Div(
                        project_pills,
                        style={
                            "display": "flex",
                            "gap": "12px",
                            "flexWrap": "wrap",
                            "flex": "1",
                            "justifyContent": "flex-end",
                        },
                    ),
                ],
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "gap": "30px",
                    "flexWrap": "wrap",
                    "marginBottom": "20px",
                },
            ),
            # KPI Row
            html.Div(
                [
                    header_kpi("Avg Completion", f"{avg_completion:.0f}%", "📈"),
                    header_kpi(
                        "Total Packages",
                        f"{int(total_pkg_done)}/{int(total_packages)}",
                        "📦",
                    ),
                    header_kpi(
                        "Total Deliveries",
                        f"{int(total_delivered)}/{int(total_pos)}",
                        "🚚",
                    ),
                    header_kpi("Active Projects", str(active), "🔥"),
                ],
                style={
                    "display": "flex",
                    "gap": "12px",
                    "flexWrap": "wrap",
                },
            ),
        ],
        style={
            "background": f"linear-gradient(135deg, {THEME['dark']} 0%, {THEME['dark2']} 50%, rgba(99,102,241,0.15) 100%)",
            "padding": "28px 36px",
            "borderRadius": "20px",
            "margin": "0 28px 24px 28px",
            "boxShadow": "0 8px 40px rgba(0,0,0,0.18)",
            "border": "1px solid rgba(255,255,255,0.05)",
        },
    )


# ============================================================
# Timeline Chart
# ============================================================

def make_timeline_chart(all_data):
    today = pd.Timestamp.now()
    fig = go.Figure()

    has_dates = False
    for i, d in enumerate(all_data):
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        name = d.get("project_name", "?")
        parsed_date = d.get("opening_date_parsed")
        opening_str = d.get("opening_date", "TBD")

        if opening_str == "Opened":
            fig.add_trace(
                go.Scatter(
                    x=[today],
                    y=[name],
                    mode="markers+text",
                    marker=dict(
                        size=18,
                        color=THEME["success"],
                        symbol="star",
                        line=dict(width=2, color="#fff"),
                    ),
                    text=["OPENED"],
                    textposition="middle right",
                    textfont=dict(
                        size=11, color=THEME["success"], family="Inter", weight=700
                    ),
                    showlegend=False,
                    hovertemplate=f"<b>{name}</b><br>Status: Already Opened<extra></extra>",
                )
            )
            has_dates = True
        elif parsed_date is not None:
            days_left = (parsed_date - today).days

            fig.add_trace(
                go.Scatter(
                    x=[today, parsed_date],
                    y=[name, name],
                    mode="lines",
                    line=dict(color=pc["main"], width=8, dash=None),
                    showlegend=False,
                    hoverinfo="skip",
                )
            )

            fig.add_trace(
                go.Scatter(
                    x=[parsed_date],
                    y=[name],
                    mode="markers+text",
                    marker=dict(
                        size=16,
                        color=pc["main"],
                        symbol="diamond",
                        line=dict(width=2, color="#fff"),
                    ),
                    text=[f"{days_left}d left" if days_left > 0 else f"{abs(days_left)}d overdue"],
                    textposition="middle right",
                    textfont=dict(
                        size=11,
                        color=THEME["danger"] if days_left < 0 else pc["main"],
                        family="Inter",
                        weight=700
                    ),
                    showlegend=False,
                    hovertemplate=(
                        f"<b>{name}</b><br>"
                        f"Opening: {parsed_date.strftime('%d %b %Y')}<br>"
                        f"Days remaining: {days_left}<extra></extra>"
                    ),
                )
            )
            has_dates = True
        else:
            fig.add_trace(
                go.Scatter(
                    x=[today],
                    y=[name],
                    mode="markers+text",
                    marker=dict(size=14, color=THEME["text_muted"], symbol="circle"),
                    text=["TBD"],
                    textposition="middle right",
                    textfont=dict(
                        size=11, color=THEME["text_muted"], family="Inter"
                    ),
                    showlegend=False,
                )
            )
            has_dates = True

    # Today line
    fig.add_vline(
        x=today.timestamp() * 1000,
        line=dict(color=THEME["danger"], width=2, dash="dot"),
        annotation_text="Today",
        annotation_position="top",
    )

    fig.update_layout(
        height=max(180, len(all_data) * 70 + 60),
        margin=dict(t=40, b=40, l=20, r=40),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        xaxis=dict(
            gridcolor="rgba(0,0,0,0.04)",
            tickfont={"size": 11, "color": THEME["text_light"]},
        ),
        yaxis=dict(
            tickfont={"size": 12, "color": THEME["text"], "family": "Inter"},
            gridcolor="rgba(0,0,0,0.02)",
        ),
    )

    return fig


# ============================================================
# Budget Waterfall per Project
# ============================================================

def make_budget_waterfall(data, pc_idx):
    pc = PROJECT_COLORS[pc_idx % len(PROJECT_COLORS)]
    currency = data.get("currency", "USD")
    budget = data.get("effective_budget", 0)
    ordered = data.get("orders_placed", 0)
    in_progress = data.get("orders_in_progress", 0)
    remaining = max(0, budget - ordered - in_progress)

    fig = go.Figure(
        go.Waterfall(
            orientation="v",
            x=["Budget", "Ordered", "In Progress", "Remaining"],
            y=[budget, -ordered, -in_progress, 0],
            measure=["absolute", "relative", "relative", "total"],
            text=[
                fmt_num(budget, currency),
                fmt_num(ordered, currency),
                fmt_num(in_progress, currency),
                fmt_num(remaining, currency),
            ],
            textposition="outside",
            textfont=dict(size=11, family="Inter"),
            connector={"line": {"color": "rgba(0,0,0,0.1)", "width": 1.5}},
            increasing={"marker": {"color": pc["main"]}},
            decreasing={"marker": {"color": THEME["warning"]}},
            totals={"marker": {"color": THEME["success"] if remaining >= 0 else THEME["danger"]}},
        )
    )

    fig.update_layout(
        height=220,
        margin=dict(t=25, b=35, l=45, r=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 10}),
        xaxis=dict(tickfont={"size": 11, "color": THEME["text_light"]}),
        showlegend=False,
    )

    return fig


# ============================================================
# Risk Assessment Section
# ============================================================

def make_risk_section(all_data):
    fig = go.Figure()

    names = [d.get("project_name", "?") for d in all_data]
    risks = [d.get("risk_score", 0) for d in all_data]
    colors_list = [risk_color(r) for r in risks]

    fig.add_trace(
        go.Bar(
            x=names,
            y=risks,
            marker_color=colors_list,
            marker_line=dict(width=0),
            text=[f"{r}" for r in risks],
            textposition="outside",
            textfont=dict(size=13, family="Inter", weight=700),
            hovertemplate="<b>%{x}</b><br>Risk Score: %{y}/100<extra></extra>",
        )
    )

    # Threshold lines
    fig.add_hline(
        y=60,
        line=dict(color=THEME["danger"], width=1.5, dash="dash"),
        annotation_text="High Risk",
        annotation_position="top right",
        annotation_font=dict(size=10, color=THEME["danger"]),
    )
    fig.add_hline(
        y=30,
        line=dict(color=THEME["warning"], width=1.5, dash="dash"),
        annotation_text="Medium Risk",
        annotation_position="top right",
        annotation_font=dict(size=10, color=THEME["warning"]),
    )

    fig.update_layout(
        height=260,
        margin=dict(t=35, b=40, l=45, r=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        yaxis=dict(
            range=[0, 110],
            gridcolor="rgba(0,0,0,0.04)",
            tickfont={"size": 11},
            title="Risk Score",
            title_font=dict(size=11, color=THEME["text_light"]),
        ),
        xaxis=dict(tickfont={"size": 12, "color": THEME["text_light"]}),
        showlegend=False,
    )

    return fig


# ============================================================
# Procurement Funnel
# ============================================================

def make_procurement_funnel(data, pc_idx):
    pc = PROJECT_COLORS[pc_idx % len(PROJECT_COLORS)]
    total_pkg = data.get("total_packages", 0)
    pkg_done = data.get("packages_completed", 0)
    pkg_wip = data.get("packages_in_progress", 0)
    total_pos_val = data.get("total_pos", 0)
    delivered = data.get("delivered_pos", 0)

    stages = ["Total Packages", "Orders Completed", "Orders In Progress", "POs Raised", "Delivered"]
    values = [total_pkg, pkg_done, pkg_wip, total_pos_val, delivered]

    if max(values) == 0:
        return go.Figure().update_layout(
            height=200,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            annotations=[
                dict(
                    text="No procurement data yet",
                    xref="paper",
                    yref="paper",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=14, color=THEME["text_muted"]),
                )
            ],
        )

    color_scale = [
        hex_to_rgba(pc["main"], 0.3),
        hex_to_rgba(pc["main"], 0.5),
        hex_to_rgba(pc["main"], 0.65),
        hex_to_rgba(pc["main"], 0.8),
        pc["main"],
    ]

    fig = go.Figure(
        go.Funnel(
            y=stages,
            x=values,
            textinfo="value+percent initial",
            textfont=dict(size=12, family="Inter"),
            marker=dict(color=color_scale, line=dict(width=0)),
            connector=dict(line=dict(color="rgba(0,0,0,0.05)", width=1)),
        )
    )

    fig.update_layout(
        height=220,
        margin=dict(t=15, b=10, l=10, r=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
    )

    return fig


# ============================================================
# Unified Comparison Section (Enhanced)
# ============================================================

def make_unified_comparison(all_data):
    names = [d.get("project_name", "?") for d in all_data]
    colors = [PROJECT_COLORS[i % len(PROJECT_COLORS)]["main"] for i in range(len(all_data))]

    # Budget comparison
    budget_fig = go.Figure()
    for i, d in enumerate(all_data):
        cur = d.get("currency", "")
        budget = d.get("effective_budget", 0)
        placed = d.get("orders_placed", 0)
        prog = d.get("orders_in_progress", 0)
        remaining = max(0, budget - placed - prog)

        budget_fig.add_trace(
            go.Bar(
                x=["Budget", "Ordered", "In Progress", "Remaining"],
                y=[budget, placed, prog, remaining],
                name=f"{names[i]} ({cur})",
                marker_color=colors[i],
                marker_line=dict(width=0),
                text=[
                    fmt_num(budget, cur),
                    fmt_num(placed, cur),
                    fmt_num(prog, cur),
                    fmt_num(remaining, cur),
                ],
                textposition="outside",
                textfont={"size": 10, "family": "Inter"},
            )
        )
    budget_fig.update_layout(
        barmode="group",
        height=300,
        margin=dict(t=35, b=45, l=55, r=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        legend=dict(
            orientation="h", y=1.12, x=0.5, xanchor="center", font={"size": 11}
        ),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 11}),
        xaxis=dict(tickfont={"size": 12, "color": THEME["text_light"]}),
    )

    # Radar chart
    categories = ["Budget Util", "Packages", "Deliveries", "Overall"]
    radar_fig = go.Figure()
    for i, d in enumerate(all_data):
        vals = [
            d.get("orders_placed_pct", 0),
            d.get("pkg_completion_pct", 0),
            d.get("delivery_completion_pct", 0),
            d.get("overall_completion", 0),
        ]
        vals_closed = vals + [vals[0]]
        cats_closed = categories + [categories[0]]
        radar_fig.add_trace(
            go.Scatterpolar(
                r=vals_closed,
                theta=cats_closed,
                name=names[i],
                fill="toself",
                fillcolor=hex_to_rgba(colors[i], 0.15),
                line=dict(color=colors[i], width=2.5),
                marker=dict(size=6, color=colors[i]),
            )
        )
    radar_fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100],
                tickfont={"size": 10},
                gridcolor="rgba(0,0,0,0.06)",
            ),
            angularaxis=dict(tickfont={"size": 12, "color": THEME["text_light"]}),
        ),
        height=320,
        margin=dict(t=45, b=25, l=65, r=65),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h", y=-0.08, x=0.5, xanchor="center", font={"size": 12}
        ),
        font={"family": "Inter"},
    )

    # Packages & Deliveries stacked bar
    pkg_del_fig = go.Figure()
    x_labels = []
    done_pkgs, wip_pkgs, pend_pkgs = [], [], []
    done_pos, pend_pos = [], []

    for i, d in enumerate(all_data):
        x_labels.append(names[i])
        done_pkgs.append(d.get("packages_completed", 0))
        wip_pkgs.append(d.get("packages_in_progress", 0))
        pend_pkgs.append(d.get("packages_to_start", 0))
        done_pos.append(d.get("delivered_pos", 0))
        pend_pos.append(max(0, d.get("total_pos", 0) - d.get("delivered_pos", 0)))

    pkg_del_fig.add_trace(
        go.Bar(
            x=[f"{n}<br>Packages" for n in x_labels],
            y=done_pkgs,
            name="Completed",
            marker_color=THEME["success"],
            text=[f"{int(v)}" for v in done_pkgs],
            textposition="inside",
            textfont={"size": 11, "color": "#fff", "family": "Inter"},
        )
    )
    pkg_del_fig.add_trace(
        go.Bar(
            x=[f"{n}<br>Packages" for n in x_labels],
            y=wip_pkgs,
            name="In Progress",
            marker_color=THEME["warning"],
            text=[f"{int(v)}" if v > 0 else "" for v in wip_pkgs],
            textposition="inside",
            textfont={"size": 11, "color": "#fff", "family": "Inter"},
        )
    )
    pkg_del_fig.add_trace(
        go.Bar(
            x=[f"{n}<br>Packages" for n in x_labels],
            y=pend_pkgs,
            name="Pending",
            marker_color="#e2e8f0",
            text=[f"{int(v)}" if v > 0 else "" for v in pend_pkgs],
            textposition="inside",
            textfont={"size": 11, "color": "#94a3b8", "family": "Inter"},
        )
    )
    pkg_del_fig.add_trace(
        go.Bar(
            x=[f"{n}<br>POs" for n in x_labels],
            y=done_pos,
            name="Delivered",
            marker_color=THEME["info"],
            text=[f"{int(v)}" for v in done_pos],
            textposition="inside",
            textfont={"size": 11, "color": "#fff", "family": "Inter"},
            showlegend=True,
        )
    )
    pkg_del_fig.add_trace(
        go.Bar(
            x=[f"{n}<br>POs" for n in x_labels],
            y=pend_pos,
            name="PO Pending",
            marker_color="#cbd5e1",
            text=[f"{int(v)}" if v > 0 else "" for v in pend_pos],
            textposition="inside",
            textfont={"size": 11, "color": "#64748b", "family": "Inter"},
            showlegend=True,
        )
    )

    pkg_del_fig.update_layout(
        barmode="stack",
        height=300,
        margin=dict(t=35, b=55, l=45, r=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        legend=dict(
            orientation="h", y=1.15, x=0.5, xanchor="center", font={"size": 11}
        ),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 11}),
        xaxis=dict(tickfont={"size": 11, "color": THEME["text_light"]}),
    )

    # Timeline
    timeline_fig = make_timeline_chart(all_data)

    # Risk
    risk_fig = make_risk_section(all_data)

    def chart_card(title, fig, min_w="380px"):
        return html.Div(
            [
                html.P(
                    title,
                    style={
                        "fontSize": F["section"],
                        "color": THEME["text_light"],
                        "fontWeight": "700",
                        "letterSpacing": "1.5px",
                        "margin": "0 0 6px 0",
                        "textTransform": "uppercase",
                    },
                ),
                dcc.Graph(figure=fig, config={"displayModeBar": False}),
            ],
            style=glass_style(
                {"padding": "22px 24px", "flex": "1", "minWidth": min_w}
            ),
        )

    return html.Div(
        [
            html.P(
                "Cross-Project Comparison",
                style={
                    "fontSize": F["section"],
                    "color": THEME["text_light"],
                    "letterSpacing": "2.5px",
                    "fontWeight": "700",
                    "margin": "0 0 18px 0",
                    "textAlign": "center",
                    "textTransform": "uppercase",
                },
            ),
            html.Div(
                [
                    chart_card("Budget Allocation", budget_fig),
                    chart_card("Completion Radar", radar_fig),
                ],
                style={
                    "display": "flex",
                    "gap": "18px",
                    "flexWrap": "wrap",
                    "marginBottom": "18px",
                },
            ),
            html.Div(
                [
                    chart_card("Packages & Deliveries Breakdown", pkg_del_fig),
                ],
                style={"display": "flex", "gap": "18px", "marginBottom": "18px"},
            ),
            html.Div(
                [
                    chart_card("Opening Timeline", timeline_fig, "350px"),
                    chart_card("Risk Assessment", risk_fig, "350px"),
                ],
                style={"display": "flex", "gap": "18px", "flexWrap": "wrap"},
            ),
        ],
        style={"padding": "0 28px 24px 28px"},
    )


# ============================================================
# Summary Table (Enhanced)
# ============================================================

def make_data_table(all_data):
    rows = []
    for d in all_data:
        budget = d.get("effective_budget", 0)
        cur = d.get("currency", "USD")
        savings = d.get("savings_overrun", 0)
        rows.append(
            {
                "Project": d.get("project_name", "Unknown"),
                "Stage": d.get("current_stage", "N/A"),
                "Opening": d.get("opening_date", "TBD"),
                "Currency": cur,
                "Budget": f"{budget:,.0f}",
                "Ordered": f"{d.get('orders_placed', 0):,.0f}",
                "In Progress": f"{d.get('orders_in_progress', 0):,.0f}",
                "Savings/Overrun": f"{savings:,.0f}",
                "Packages": f"{int(d.get('packages_completed', 0))}/{int(d.get('total_packages', 0))}",
                "POs Del.": f"{int(d.get('delivered_pos', 0))}/{int(d.get('total_pos', 0))}",
                "Completion": f"{d.get('overall_completion', 0):.1f}%",
                "Risk": risk_label(d.get("risk_score", 0)),
            }
        )
    df_table = pd.DataFrame(rows)

    style_data_conditional = [
        {"if": {"row_index": "odd"}, "backgroundColor": "rgba(248,250,252,0.7)"},
        {"if": {"row_index": "even"}, "backgroundColor": "rgba(255,255,255,0.7)"},
    ]

    # Color the Risk column
    for i, d in enumerate(all_data):
        rc = risk_color(d.get("risk_score", 0))
        style_data_conditional.append(
            {
                "if": {"row_index": i, "column_id": "Risk"},
                "color": rc,
                "fontWeight": "700",
            }
        )
        # Color savings
        sav = d.get("savings_overrun", 0)
        style_data_conditional.append(
            {
                "if": {"row_index": i, "column_id": "Savings/Overrun"},
                "color": THEME["success"] if sav >= 0 else THEME["danger"],
                "fontWeight": "700",
            }
        )

    return html.Div(
        [
            html.P(
                "Project Summary",
                style={
                    "fontSize": F["section"],
                    "color": THEME["text_light"],
                    "letterSpacing": "2.5px",
                    "fontWeight": "700",
                    "margin": "0 0 14px 0",
                    "textAlign": "center",
                    "textTransform": "uppercase",
                },
            ),
            html.Div(
                [
                    dash_table.DataTable(
                        data=df_table.to_dict("records"),
                        columns=[{"name": c, "id": c} for c in df_table.columns],
                        style_table={"overflowX": "auto", "borderRadius": "14px"},
                        style_header={
                            "backgroundColor": THEME["dark"],
                            "color": "#fff",
                            "fontWeight": "700",
                            "fontSize": F["sm"],
                            "textAlign": "center",
                            "padding": "14px 12px",
                            "fontFamily": "Inter",
                            "borderBottom": "none",
                        },
                        style_cell={
                            "textAlign": "center",
                            "padding": "12px 14px",
                            "fontSize": F["sm"],
                            "fontFamily": "Inter",
                            "border": "none",
                            "borderBottom": f"1px solid {THEME['border']}",
                            "minWidth": "80px",
                            "color": THEME["text"],
                        },
                        style_data_conditional=style_data_conditional,
                    )
                ],
                style=glass_style({"overflow": "hidden", "padding": "0"}),
            ),
        ],
        style={"padding": "0 28px 28px 28px"},
    )


# ============================================================
# Project Card (Enhanced)
# ============================================================

def make_project_card(data, pc_idx):
    pc = PROJECT_COLORS[pc_idx % len(PROJECT_COLORS)]

    name = data.get("project_name", "Unknown")
    stage = data.get("current_stage", "N/A")
    currency = data.get("currency", "USD")
    opening = data.get("opening_date", "TBD")

    budget = data.get("effective_budget", 0)
    savings = data.get("savings_overrun", 0)

    pkgs = data.get("total_packages", 0)
    pkg_done = data.get("packages_completed", 0)
    pkg_wip = data.get("packages_in_progress", 0)

    pos = data.get("total_pos", 0)
    del_done = data.get("delivered_pos", 0)

    overall = data.get("overall_completion", 0)
    risk_score = data.get("risk_score", 0)

    concerns = data.get("concerns", [])

    ai_summary = generate_ai_summary(data)

    stage_lower = stage.lower()
    if any(k in stage_lower for k in ["closing", "handover"]):
        stage_bg = THEME["success"]
    elif any(k in stage_lower for k in ["po issuance", "deliveries", "order"]):
        stage_bg = THEME["info"]
    elif any(k in stage_lower for k in ["budget", "boq"]):
        stage_bg = THEME["warning"]
    else:
        stage_bg = THEME["text_muted"]

    oc = status_color(overall)
    rc = risk_color(risk_score)

    # Overall gauge
    gauge = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=round(overall, 1),
            number={
                "suffix": "%",
                "font": {"size": 36, "color": pc["main"], "family": "Inter"},
            },
            gauge={
                "axis": {"range": [0, 100], "tickfont": {"size": 10}},
                "bar": {"color": oc, "thickness": 0.7},
                "bgcolor": "#f1f5f9",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 40], "color": hex_to_rgba(THEME["danger"], 0.08)},
                    {"range": [40, 75], "color": hex_to_rgba(THEME["warning"], 0.08)},
                    {"range": [75, 100], "color": hex_to_rgba(THEME["success"], 0.08)},
                ],
            },
        )
    )
    gauge.update_layout(
        height=160,
        margin=dict(t=15, b=0, l=20, r=20),
        paper_bgcolor="rgba(0,0,0,0)",
    )

    # Budget waterfall
    waterfall_fig = make_budget_waterfall(data, pc_idx)

    # Procurement funnel
    funnel_fig = make_procurement_funnel(data, pc_idx)

    def kpi_chip(label, value, icon, color):
        return html.Div(
            [
                html.Span(icon, style={"fontSize": "18px"}),
                html.Div(
                    [
                        html.P(
                            label,
                            style={
                                "fontSize": F["xs"],
                                "color": THEME["text_muted"],
                                "margin": "0",
                            },
                        ),
                        html.P(
                            value,
                            style={
                                "fontSize": F["value"],
                                "color": THEME["text"],
                                "margin": "2px 0 0 0",
                                "fontWeight": "700",
                            },
                        ),
                    ]
                ),
            ],
            style={
                "display": "flex",
                "alignItems": "center",
                "gap": "8px",
                "backgroundColor": hex_to_rgba(color, 0.06),
                "padding": "10px 12px",
                "borderRadius": "12px",
                "flex": "1",
                "minWidth": "130px",
            },
        )

    sav_color = THEME["success"] if savings >= 0 else THEME["danger"]
    sav_label = "Savings" if savings >= 0 else "Overrun"

    kpi_row = html.Div(
        [
            kpi_chip("Budget", fmt_num(budget, currency), "💰", pc["main"]),
            kpi_chip(sav_label, fmt_num(abs(savings), currency), "📈" if savings >= 0 else "📉", sav_color),
            kpi_chip("Packages", f"{int(pkg_done)}/{int(pkgs)}", "📦", THEME["info"]),
            kpi_chip("Deliveries", f"{int(del_done)}/{int(pos)}", "🚚", THEME["success"]),
        ],
        style={
            "display": "flex",
            "gap": "8px",
            "flexWrap": "wrap",
            "marginBottom": "16px",
        },
    )

    # Proc / Delivery status badges
    proc_started = data.get("proc_started", "").lower() == "yes"
    delivery_started = data.get("delivery_started", "").lower() == "yes"

    def status_badge(label, active):
        color = THEME["success"] if active else THEME["text_muted"]
        return html.Span(
            f"{'✓' if active else '○'} {label}",
            style={
                "fontSize": F["xs"],
                "fontWeight": "700",
                "color": color,
                "backgroundColor": hex_to_rgba(color, 0.1),
                "padding": "4px 10px",
                "borderRadius": "20px",
                "border": f"1px solid {hex_to_rgba(color, 0.25)}",
            },
        )

    status_row = html.Div(
        [
            status_badge("Procurement Started", proc_started),
            status_badge("Delivery Started", delivery_started),
            html.Span(
                f"Opening: {opening}",
                style={
                    "fontSize": F["xs"],
                    "fontWeight": "600",
                    "color": THEME["info"],
                    "backgroundColor": hex_to_rgba(THEME["info"], 0.08),
                    "padding": "4px 10px",
                    "borderRadius": "20px",
                },
            ),
            html.Span(
                risk_label(risk_score),
                style={
                    "fontSize": F["xs"],
                    "fontWeight": "700",
                    "color": rc,
                    "backgroundColor": hex_to_rgba(rc, 0.12),
                    "padding": "4px 10px",
                    "borderRadius": "20px",
                    "border": f"1px solid {hex_to_rgba(rc, 0.3)}",
                },
            ),
        ],
        style={
            "display": "flex",
            "gap": "8px",
            "flexWrap": "wrap",
            "marginBottom": "16px",
        },
    )

    # Concerns
    if concerns:
        cblock = html.Div(
            [
                html.Div(
                    c,
                    style={
                        "padding": "10px 14px",
                        "backgroundColor": "#fffbeb",
                        "borderRadius": "10px",
                        "fontSize": F["sm"],
                        "borderLeft": f"3px solid {THEME['warning']}",
                        "marginBottom": "6px",
                        "color": "#92400e",
                        "lineHeight": "1.5",
                    },
                )
                for c in concerns
            ]
        )
    else:
        cblock = html.P(
            "No concerns reported",
            style={
                "color": THEME["success"],
                "fontSize": F["sm"],
                "margin": "0",
                "padding": "10px 14px",
                "backgroundColor": hex_to_rgba(THEME["success"], 0.06),
                "borderRadius": "10px",
                "borderLeft": f"3px solid {THEME['success']}",
            },
        )

    # AI block
    ai_block = html.Div(
        [
            html.P(
                "AI PROJECT SUMMARY",
                style={
                    "fontSize": F["xs"],
                    "color": THEME["text_muted"],
                    "letterSpacing": "1.5px",
                    "fontWeight": "700",
                    "margin": "16px 0 10px 0",
                },
            ),
            html.Div(
                ai_summary,
                style={
                    "backgroundColor": "#eef2ff",
                    "padding": "12px 14px",
                    "borderRadius": "10px",
                    "fontSize": F["sm"],
                    "color": "#3730a3",
                    "borderLeft": "3px solid #6366f1",
                    "lineHeight": "1.6",
                },
            ),
        ]
    )

    def section_label(text):
        return html.P(
            text,
            style={
                "fontSize": F["xs"],
                "color": THEME["text_muted"],
                "letterSpacing": "1.5px",
                "fontWeight": "700",
                "margin": "16px 0 8px 0",
                "textTransform": "uppercase",
            },
        )

    return html.Div(
        [
            # Header
            html.Div(
                [
                    html.H3(
                        name,
                        style={
                            "margin": "0",
                            "color": "#fff",
                            "fontSize": F["xl"],
                            "fontWeight": "800",
                        },
                    ),
                    html.Span(
                        stage,
                        style={
                            "backgroundColor": stage_bg,
                            "padding": "4px 12px",
                            "borderRadius": "20px",
                            "fontSize": F["xs"],
                            "color": "#fff",
                            "fontWeight": "600",
                            "whiteSpace": "nowrap",
                        },
                    ),
                ],
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "background": pc["gradient"],
                    "padding": "18px 22px",
                    "borderRadius": "18px 18px 0 0",
                    "gap": "12px",
                    "flexWrap": "wrap",
                },
            ),
            # Body
            html.Div(
                [
                    status_row,
                    kpi_row,
                    # Overall completion gauge
                    section_label("Overall Procurement Completion"),
                    dcc.Graph(figure=gauge, config={"displayModeBar": False}),
                    # Budget waterfall
                    section_label("Budget Breakdown"),
                    dcc.Graph(figure=waterfall_fig, config={"displayModeBar": False}),
                    # Procurement funnel
                    section_label("Procurement Funnel"),
                    dcc.Graph(figure=funnel_fig, config={"displayModeBar": False}),
                    # Concerns
                    section_label("Concerns & Notes"),
                    cblock,
                    # AI Summary
                    ai_block,
                ],
                style={
                    "padding": "20px 24px",
                    "backgroundColor": hex_to_rgba(pc["main"], 0.03),
                },
            ),
        ],
        style=glass_style(
            {
                "overflow": "hidden",
                "flex": "1",
                "minWidth": "440px",
                "maxWidth": "600px",
            }
        ),
    )


# ============================================================
# DASH APP
# ============================================================

app = dash.Dash(
    __name__,
    meta_tags=[
        {"name": "viewport", "content": "width=device-width, initial-scale=1.0"}
    ],
    suppress_callback_exceptions=True,
)

# ============================================================
# Fullscreen Button Styles
# ============================================================

FULLSCREEN_BTN_STYLE = {
    "backgroundColor": "rgba(99, 102, 241, 0.15)",
    "color": "#a5b4fc",
    "border": "1px solid rgba(99, 102, 241, 0.3)",
    "padding": "10px 20px",
    "borderRadius": "10px",
    "cursor": "pointer",
    "fontSize": F["sm"],
    "fontWeight": "600",
    "display": "flex",
    "alignItems": "center",
    "gap": "8px",
    "transition": "all 0.3s ease",
}

EXIT_FULLSCREEN_BTN_STYLE = {
    "backgroundColor": "rgba(239, 68, 68, 0.15)",
    "color": "#fca5a5",
    "border": "1px solid rgba(239, 68, 68, 0.3)",
    "padding": "10px 20px",
    "borderRadius": "10px",
    "cursor": "pointer",
    "fontSize": F["sm"],
    "fontWeight": "600",
    "display": "none",
    "alignItems": "center",
    "gap": "8px",
    "transition": "all 0.3s ease",
}

app.layout = html.Div(
    [
        # Upload page
        html.Div(
            id="upload-section",
            children=[
                html.Div(style={"height": "15vh"}),
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    style={
                                        "width": "72px",
                                        "height": "72px",
                                        "borderRadius": "20px",
                                        "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                                        "display": "flex",
                                        "alignItems": "center",
                                        "justifyContent": "center",
                                        "margin": "0 auto 20px auto",
                                        "boxShadow": "0 8px 24px rgba(99,102,241,0.35)",
                                    },
                                    children=[
                                        html.Img(
                                            src="/assets/luxurylogo.jpg",
                                            style={
                                                    "height": "100px",
                                                    "marginRight": "15px"
                                                }
                                        )
                                        
                                    ],
                                ),
                                html.H1(
                                    "Luxury Hospitality Dashboard",
                                    style={
                                        "margin": "0 0 8px 0",
                                        "color": THEME["text"],
                                        "fontSize": "32px",
                                        "fontWeight": "800",
                                        "letterSpacing": "-0.5px",
                                    },
                                ),
                                html.P(
                                    "Upload procurement data file to generate insights",
                                    style={
                                        "color": THEME["text_muted"],
                                        "fontSize": F["md"],
                                        "margin": "0 0 36px 0",
                                        "fontWeight": "400",
                                    },
                                ),
                            ]
                        ),
                        dcc.Upload(
                            id="upload-data",
                            children=html.Div(
                                [
                                    html.Div(
                                        [
                                            html.Div(
                                                style={
                                                    "width": "56px",
                                                    "height": "56px",
                                                    "borderRadius": "14px",
                                                    "backgroundColor": "#eef2ff",
                                                    "display": "flex",
                                                    "alignItems": "center",
                                                    "justifyContent": "center",
                                                    "margin": "0 auto 14px auto",
                                                },
                                                children=[
                                                    html.Span(
                                                        "📁",
                                                        style={"fontSize": "26px"},
                                                    )
                                                ],
                                            ),
                                            html.P(
                                                "Drag & Drop or Click to Upload",
                                                style={
                                                    "fontWeight": "700",
                                                    "color": "#6366f1",
                                                    "fontSize": F["lg"],
                                                    "margin": "0 0 6px 0",
                                                },
                                            ),
                                            html.P(
                                                ".xlsx  •  .xlsm  •  .xls  •  .csv",
                                                style={
                                                    "color": THEME["text_muted"],
                                                    "fontSize": F["sm"],
                                                    "margin": "0",
                                                },
                                            ),
                                        ]
                                    )
                                ]
                            ),
                            style={
                                "width": "100%",
                                "padding": "45px 20px",
                                "borderWidth": "2px",
                                "borderStyle": "dashed",
                                "borderColor": "#c7d2fe",
                                "borderRadius": "16px",
                                "textAlign": "center",
                                "backgroundColor": "rgba(250,251,255,0.8)",
                                "cursor": "pointer",
                            },
                            multiple=False,
                        ),
                        html.Div(id="upload-error", style={"marginTop": "18px"}),
                    ],
                    style=glass_style(
                        {
                            "padding": "52px 60px",
                            "maxWidth": "520px",
                            "margin": "0 auto",
                            "textAlign": "center",
                        }
                    ),
                ),
            ],
        ),
        # Dashboard page
        html.Div(
            id="dashboard-section",
            style={"display": "none"},
            children=[
                # Top bar
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    style={
                                        "width": "42px",
                                        "height": "42px",
                                        "borderRadius": "12px",
                                        "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                                        "display": "flex",
                                        "alignItems": "center",
                                        "justifyContent": "center",
                                        "boxShadow": "0 4px 12px rgba(99,102,241,0.3)",
                                    },
                                    children=[
                                        html.Img(
                                                src="/assets/luxurylogo.jpg",
                                                style={
                                                    "height": "60px",
                                                    "marginRight": "15px"
                                                }
                                            )
                                    ],
                                ),
                                html.Div(
                                    [
                                        html.H1(
                                            "Luxury Hospitality Dashboard",
                                            style={
                                                "margin": "0",
                                                "color": "#fff",
                                                "fontSize": F["xl"],
                                                "fontWeight": "800",
                                                "letterSpacing": "-0.3px",
                                            },
                                        ),
                                        html.P(
                                            "A Journey of Refined Luxury",
                                            style={
                                                "color": "rgba(255,255,255,0.5)",
                                                "fontSize": F["sm"],
                                                "margin": "0",
                                                "fontWeight": "500",
                                            },
                                        ),
                                    ]
                                ),
                            ],
                            style={
                                "display": "flex",
                                "alignItems": "center",
                                "gap": "14px",
                                "flex": "1",
                            },
                        ),
                        # Button container
                        html.Div(
                            [
                                html.Button(
                                    [
                                        html.Span(
                                            "⛶", style={"fontSize": "16px"}
                                        ),
                                        html.Span("Full Screen"),
                                    ],
                                    id="btn-fullscreen",
                                    n_clicks=0,
                                    style=FULLSCREEN_BTN_STYLE,
                                ),
                                html.Button(
                                    [
                                        html.Span(
                                            "✕",
                                            style={
                                                "fontSize": "14px",
                                                "fontWeight": "bold",
                                            },
                                        ),
                                        html.Span("Exit Full Screen"),
                                    ],
                                    id="btn-exit-fullscreen",
                                    n_clicks=0,
                                    style=EXIT_FULLSCREEN_BTN_STYLE,
                                ),
                                html.A(
                                    "← Upload New File",
                                    id="btn-back",
                                    href="/",
                                    style={
                                        "backgroundColor": "rgba(255,255,255,0.08)",
                                        "color": "#e2e8f0",
                                        "border": "1px solid rgba(255,255,255,0.12)",
                                        "padding": "10px 24px",
                                        "borderRadius": "10px",
                                        "cursor": "pointer",
                                        "fontSize": F["sm"],
                                        "fontWeight": "600",
                                        "textDecoration": "none",
                                        "display": "inline-block",
                                        "marginLeft": "10px",
                                    },
                                ),
                            ],
                            style={
                                "display": "flex",
                                "alignItems": "center",
                                "gap": "8px",
                            },
                        ),
                    ],
                    id="top-bar",
                    style={
                        "display": "flex",
                        "justifyContent": "space-between",
                        "alignItems": "center",
                        "background": f"linear-gradient(135deg, {THEME['dark']}, {THEME['dark2']})",
                        "padding": "16px 34px",
                        "marginBottom": "24px",
                        "boxShadow": "0 4px 20px rgba(0,0,0,0.15)",
                        "position": "sticky",
                        "top": "0",
                        "zIndex": "1000",
                    },
                ),
                html.Div(id="dashboard-body"),
            ],
        ),
    ],
    id="main-container",
    style={
        "fontFamily": "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        "backgroundImage": "url('/assets/background.jpg')",
        "backgroundSize": "cover",
        "backgroundPosition": "center",
        "backgroundRepeat": "no-repeat",
        "backgroundAttachment": "fixed",
        "minHeight": "100vh",
    },
)


# ============================================================
# Clientside Callback for Fullscreen
# ============================================================

app.clientside_callback(
    """
    function(n_clicks_enter, n_clicks_exit) {
        var elem = document.getElementById('main-container');
        
        var triggered = dash_clientside.callback_context.triggered;
        if (!triggered || triggered.length === 0) {
            return [window.dash_clientside.no_update, window.dash_clientside.no_update];
        }
        
        var triggeredId = triggered[0].prop_id.split('.')[0];
        
        var enterBtn = document.getElementById('btn-fullscreen');
        var exitBtn = document.getElementById('btn-exit-fullscreen');
        
        if (triggeredId === 'btn-fullscreen') {
            if (elem.requestFullscreen) {
                elem.requestFullscreen();
            } else if (elem.webkitRequestFullscreen) {
                elem.webkitRequestFullscreen();
            } else if (elem.msRequestFullscreen) {
                elem.msRequestFullscreen();
            }
            elem.style.overflow = 'auto';
            elem.style.height = '100vh';
            elem.style.maxHeight = '100vh';
            if (enterBtn) enterBtn.style.display = 'none';
            if (exitBtn) exitBtn.style.display = 'flex';
        } else if (triggeredId === 'btn-exit-fullscreen') {
            if (document.exitFullscreen) {
                document.exitFullscreen();
            } else if (document.webkitExitFullscreen) {
                document.webkitExitFullscreen();
            } else if (document.msExitFullscreen) {
                document.msExitFullscreen();
            }
            elem.style.overflow = '';
            elem.style.height = '';
            elem.style.maxHeight = '';
            if (enterBtn) enterBtn.style.display = 'flex';
            if (exitBtn) exitBtn.style.display = 'none';
        }
        
        document.onfullscreenchange = function() {
            var enterBtn = document.getElementById('btn-fullscreen');
            var exitBtn = document.getElementById('btn-exit-fullscreen');
            var elem = document.getElementById('main-container');
            if (!document.fullscreenElement) {
                if (enterBtn) enterBtn.style.display = 'flex';
                if (exitBtn) exitBtn.style.display = 'none';
                elem.style.overflow = '';
                elem.style.height = '';
                elem.style.maxHeight = '';
            } else {
                elem.style.overflow = 'auto';
                elem.style.height = '100vh';
                elem.style.maxHeight = '100vh';
            }
        };
        document.onwebkitfullscreenchange = document.onfullscreenchange;
        
        return [window.dash_clientside.no_update, window.dash_clientside.no_update];
    }
    """,
    Output("btn-fullscreen", "style"),
    Output("btn-exit-fullscreen", "style"),
    Input("btn-fullscreen", "n_clicks"),
    Input("btn-exit-fullscreen", "n_clicks"),
    prevent_initial_call=True,
)


# ============================================================
# Main Upload Callback
# ============================================================

@app.callback(
    Output("upload-section", "style"),
    Output("dashboard-section", "style"),
    Output("dashboard-body", "children"),
    Output("upload-error", "children"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, ""

    ext = ""
    if filename and "." in filename:
        ext = filename.lower().rsplit(".", 1)[-1]

    if ext not in ["xlsx", "xlsm", "xls", "csv"]:
        return (
            dash.no_update,
            dash.no_update,
            dash.no_update,
            html.P(
                f"Unsupported format: .{ext}",
                style={
                    "color": THEME["danger"],
                    "fontWeight": "600",
                    "fontSize": F["md"],
                },
            ),
        )

    try:
        all_data = parse_file(contents, filename)
    except Exception as e:
        return (
            dash.no_update,
            dash.no_update,
            dash.no_update,
            html.P(
                f"Error processing file: {str(e)}",
                style={"color": THEME["danger"], "fontSize": F["sm"]},
            ),
        )

    if not all_data:
        return (
            dash.no_update,
            dash.no_update,
            dash.no_update,
            html.P(
                "No valid project data found in this file.",
                style={"color": THEME["danger"], "fontSize": F["sm"]},
            ),
        )

    cards = []
    for idx, d in enumerate(all_data):
        cards.append(make_project_card(d, idx))

    header = make_portfolio_header(all_data)
    comparison = (
        make_unified_comparison(all_data) if len(all_data) >= 2 else html.Div()
    )
    table = make_data_table(all_data)

    body = html.Div(
        [
            header,
            comparison,
            table,
            html.Div(style={"height": "12px"}),
            html.P(
                "Project Details",
                style={
                    "fontSize": F["section"],
                    "color": THEME["text_light"],
                    "letterSpacing": "2.5px",
                    "fontWeight": "700",
                    "margin": "0 0 16px 0",
                    "textAlign": "center",
                    "textTransform": "uppercase",
                },
            ),
            html.Div(
                cards,
                style={
                    "display": "flex",
                    "justifyContent": "center",
                    "alignItems": "flex-start",
                    "gap": "24px",
                    "flexWrap": "wrap",
                    "padding": "0 28px 48px 28px",
                },
            ),
        ]
    )

    return {"display": "none"}, {"display": "block"}, body, ""


# ============================================================
# Run
# ============================================================

if __name__ == "__main__":
    port = 8050
    url = f"http://127.0.0.1:{port}"
    print("\n" + "=" * 50)
    print(f"  Dashboard running at: {url}")
    print("=" * 50 + "\n")
    Timer(2.0, open_browser, args=[url]).start()
    app.run(debug=False, port=port, host="127.0.0.1")
