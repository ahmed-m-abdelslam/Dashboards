import os
import pandas as pd
import dash
from dash import dcc, html, Input, Output, State, dash_table, clientside_callback
import plotly.graph_objects as go
import plotly.express as px
import io
import base64
import subprocess
import platform
from threading import Timer
from datetime import datetime, timedelta
import numpy as np
from groq import Groq
from dotenv import load_dotenv

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

        # Budget metrics
        committed = data["orders_placed"] + data["orders_in_progress"]
        data["committed_spend"] = committed
        data["budget_utilization_pct"] = (committed / budget * 100) if budget > 0 else 0
        data["budget_variance"] = budget - committed
        data["budget_variance_pct"] = (data["budget_variance"] / budget * 100) if budget > 0 else 0

        if budget > 0 and data["orders_placed"] > 0:
            data["savings_overrun"] = budget - committed
        else:
            data["savings_overrun"] = 0

        data["budget_uncommitted"] = max(0, budget - committed)

        # Projected final cost
        if data["packages_completed"] > 0 and data["total_packages"] > 0:
            avg_cost_per_pkg = data["orders_placed"] / data["packages_completed"]
            remaining_pkgs = data["total_packages"] - data["packages_completed"]
            data["projected_final_cost"] = data["orders_placed"] + data["orders_in_progress"] + (remaining_pkgs * avg_cost_per_pkg * 0.7)
        else:
            data["projected_final_cost"] = committed

        data["projected_overrun"] = data["projected_final_cost"] - budget
        data["budget_at_risk"] = data["projected_final_cost"] > budget * 0.95

        # Pipeline metrics
        initiated = data["packages_completed"] + data["packages_in_progress"]
        data["pipeline_initiation_rate"] = (initiated / data["total_packages"] * 100) if data["total_packages"] > 0 else 0
        data["pipeline_closure_rate"] = (data["packages_completed"] / initiated * 100) if initiated > 0 else 0
        data["ordering_completion_rate"] = pkg_completion

        # PO metrics
        data["po_delivery_rate"] = delivery_completion
        data["outstanding_pos"] = max(0, data["total_pos"] - data["delivered_pos"])
        data["pos_per_package"] = (data["total_pos"] / data["packages_completed"]) if data["packages_completed"] > 0 else 0
        data["delivery_gap"] = data["overall_completion"] - data["po_delivery_rate"]

        # Timeline metrics
        today = pd.Timestamp.now()
        if data["opening_date_parsed"]:
            days_left = (data["opening_date_parsed"] - today).days
            data["days_to_opening"] = days_left
            weeks_left = max(1, days_left / 7)
            data["delivery_pressure_index"] = data["outstanding_pos"] / weeks_left if days_left > 0 else 999
            remaining_completion = 100 - data["overall_completion"]
            data["required_daily_rate"] = remaining_completion / max(1, days_left) if days_left > 0 else 999

            if days_left < 0:
                data["urgency_category"] = "Overdue"
            elif days_left <= 90:
                data["urgency_category"] = "Imminent"
            elif days_left <= 180:
                data["urgency_category"] = "Near-term"
            elif days_left <= 365:
                data["urgency_category"] = "Medium-term"
            else:
                data["urgency_category"] = "Long-term"
        elif data["opening_date"] == "Opened":
            data["days_to_opening"] = 0
            data["delivery_pressure_index"] = 0
            data["required_daily_rate"] = 0
            data["urgency_category"] = "Opened"
        else:
            data["days_to_opening"] = None
            data["delivery_pressure_index"] = 0
            data["required_daily_rate"] = 0
            data["urgency_category"] = "TBD"

        # Schedule Performance Index
        if data["opening_date_parsed"] and data["overall_completion"] > 0:
            total_duration_est = 365
            elapsed_ratio = min(1.0, max(0.1, 1.0 - (data.get("days_to_opening", 365) / total_duration_est)))
            expected_completion = elapsed_ratio * 100
            data["schedule_performance_index"] = data["overall_completion"] / max(1, expected_completion)
        else:
            data["schedule_performance_index"] = 1.0

        # Completion status
        if data["overall_completion"] >= 75:
            data["completion_status"] = "On Track"
        elif data["overall_completion"] >= 40:
            data["completion_status"] = "Monitor"
        else:
            data["completion_status"] = "At Risk"

        # Risk Score (composite)
        risk = 0
        # Schedule risk
        if data["opening_date_parsed"]:
            dl = data.get("days_to_opening", 999)
            if dl < 0:
                risk += 15
            elif dl < 90 and data["overall_completion"] < 70:
                risk += 35
            elif dl < 180 and data["overall_completion"] < 40:
                risk += 25
            elif dl < 90:
                risk += 15
        # Budget risk
        if data["budget_at_risk"]:
            risk += 20
        if data["savings_overrun"] < 0:
            risk += 15
        # Delivery risk
        if data.get("delivery_pressure_index", 0) > 5:
            risk += 15
        elif data.get("delivery_pressure_index", 0) > 2:
            risk += 8
        # Pipeline risk
        if data["total_packages"] > 0 and data["packages_to_start"] / data["total_packages"] > 0.4:
            risk += 10
        elif data["total_packages"] > 0 and data["packages_to_start"] / data["total_packages"] > 0.2:
            risk += 5
        # Completion risk
        if data["overall_completion"] < 20 and data["opening_date"] not in ["TBD", "Opened"]:
            risk += 10

        data["risk_score"] = min(100, risk)

        # Risk breakdown
        data["schedule_risk"] = min(40, risk)
        data["budget_risk"] = 20 if data["budget_at_risk"] else (15 if data["savings_overrun"] < 0 else 0)
        data["delivery_risk"] = min(15, data.get("delivery_pressure_index", 0) * 3)

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
        data["concern_count"] = len(concerns)

        # Concern categorization
        concern_categories = set()
        concern_text = " ".join(concerns).lower()
        if any(w in concern_text for w in ["supplier", "vendor", "manufacturer"]):
            concern_categories.add("Supplier")
        if any(w in concern_text for w in ["delay", "late", "behind", "slow"]):
            concern_categories.add("Timeline")
        if any(w in concern_text for w in ["cost", "price", "budget", "expensive", "overrun"]):
            concern_categories.add("Cost")
        if any(w in concern_text for w in ["quality", "defect", "damage", "reject"]):
            concern_categories.add("Quality")
        if any(w in concern_text for w in ["approval", "pending", "waiting", "hold"]):
            concern_categories.add("Approval")
        if any(w in concern_text for w in ["lead time", "shipping", "logistics", "freight"]):
            concern_categories.add("Logistics")
        data["concern_categories"] = list(concern_categories)
        data["concern_severity"] = len(concern_categories)

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
                    subprocess.Popen([b, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
# Theme & Constants
# ============================================================

THEME = {
    "bg": "#f8fafc",
    "card_bg": "#ffffff",
    "dark": "#0f172a",
    "dark2": "#1e293b",
    "dark3": "#334155",
    "text": "#1e293b",
    "text_light": "#64748b",
    "text_muted": "#94a3b8",
    "border": "#e2e8f0",
    "border_light": "#f1f5f9",
    "success": "#10b981",
    "success_light": "#d1fae5",
    "warning": "#f59e0b",
    "warning_light": "#fef3c7",
    "danger": "#ef4444",
    "danger_light": "#fee2e2",
    "info": "#3b82f6",
    "info_light": "#dbeafe",
    "purple": "#8b5cf6",
    "purple_light": "#ede9fe",
    "indigo": "#6366f1",
    "cyan": "#06b6d4",
    "pink": "#ec4899",
    "teal": "#14b8a6",
    "orange": "#f97316",
}

PROJECT_COLORS = [
    {"main": "#6366f1", "light": "#eef2ff", "gradient": "linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%)"},
    {"main": "#0ea5e9", "light": "#f0f9ff", "gradient": "linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%)"},
    {"main": "#f97316", "light": "#fff7ed", "gradient": "linear-gradient(135deg, #f97316 0%, #fb923c 100%)"},
    {"main": "#ec4899", "light": "#fdf2f8", "gradient": "linear-gradient(135deg, #ec4899 0%, #f472b6 100%)"},
    {"main": "#14b8a6", "light": "#f0fdfa", "gradient": "linear-gradient(135deg, #14b8a6 0%, #2dd4bf 100%)"},
    {"main": "#8b5cf6", "light": "#f5f3ff", "gradient": "linear-gradient(135deg, #8b5cf6 0%, #a78bfa 100%)"},
    {"main": "#ef4444", "light": "#fef2f2", "gradient": "linear-gradient(135deg, #ef4444 0%, #f87171 100%)"},
]

F = {
    "xs": "11px", "sm": "12.5px", "md": "14px", "lg": "17px",
    "xl": "20px", "xxl": "26px", "xxxl": "34px",
}

SPACING = {"xs": "4px", "sm": "8px", "md": "12px", "lg": "16px", "xl": "20px", "xxl": "24px", "xxxl": "32px"}


def fmt_num(val, currency=""):
    prefix = f"{currency} " if currency else ""
    if abs(val) >= 1_000_000:
        return f"{prefix}{val / 1_000_000:,.2f}M"
    elif abs(val) >= 1_000:
        return f"{prefix}{val / 1_000:,.1f}K"
    else:
        return f"{prefix}{val:,.0f}"


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
        return "Medium"
    return "Low Risk"


def urgency_color(cat):
    mapping = {
        "Overdue": THEME["danger"],
        "Imminent": THEME["orange"],
        "Near-term": THEME["warning"],
        "Medium-term": THEME["info"],
        "Long-term": THEME["success"],
        "Opened": THEME["purple"],
        "TBD": THEME["text_muted"],
    }
    return mapping.get(cat, THEME["text_muted"])


def glass_card(extra=None):
    base = {
        "backgroundColor": "rgba(255,255,255,0.92)",
        "backdropFilter": "blur(20px)",
        "WebkitBackdropFilter": "blur(20px)",
        "borderRadius": "14px",
        "border": "1px solid rgba(226,232,240,0.8)",
        "boxShadow": "0 1px 3px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.04)",
        "transition": "box-shadow 0.2s ease",
    }
    if extra:
        base.update(extra)
    return base


def generate_ai_summary(project):
    try:
        data_text = f"""
Project: {project.get("project_name")} | Stage: {project.get("current_stage")}
Opening: {project.get("opening_date")} | Days Left: {project.get("days_to_opening", "N/A")}
Budget: {project.get("effective_budget")} {project.get("currency")}
Committed: {project.get("committed_spend")} | Utilization: {project.get("budget_utilization_pct", 0):.1f}%
Projected Final Cost: {project.get("projected_final_cost", 0):.0f}
Packages: {project.get("packages_completed")}/{project.get("total_packages")} completed, {project.get("packages_in_progress")} in progress
POs: {project.get("delivered_pos")}/{project.get("total_pos")} delivered
Overall Completion: {project.get("overall_completion")}%
Risk Score: {project.get("risk_score")}/100
Schedule Performance: {project.get("schedule_performance_index", 1):.2f}
Delivery Pressure: {project.get("delivery_pressure_index", 0):.1f}
Concerns: {'; '.join(project.get("concerns", [])) or 'None'}
"""
        completion = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{
                "role": "user",
                "content": f"""You are a senior procurement analyst. Provide a concise executive summary (3-4 sentences) covering:
1. Current procurement status and progress assessment
2. Key risks or concerns identified
3. Recommended immediate actions
Be specific with numbers. Use professional tone.
{data_text}"""
            }],
            temperature=0.3,
        )
        return completion.choices[0].message.content.strip()
    except Exception:
        return "AI summary unavailable — check API configuration."


def generate_portfolio_ai_summary(all_data):
    try:
        summary_lines = []
        for d in all_data:
            summary_lines.append(
                f"- {d['project_name']}: {d['overall_completion']:.0f}% complete, "
                f"Budget util: {d['budget_utilization_pct']:.0f}%, "
                f"Risk: {d['risk_score']}/100, "
                f"Opening: {d['opening_date']}, "
                f"Concerns: {len(d['concerns'])}"
            )
        portfolio_text = "\n".join(summary_lines)

        total_budget = sum(d["effective_budget"] for d in all_data)
        total_committed = sum(d["committed_spend"] for d in all_data)
        avg_completion = sum(d["overall_completion"] for d in all_data) / len(all_data)
        at_risk = sum(1 for d in all_data if d["risk_score"] >= 50)

        completion = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{
                "role": "user",
                "content": f"""You are a senior procurement portfolio manager. Analyze this portfolio and provide:
1. Portfolio health assessment (1 sentence)
2. Top 2-3 critical issues across all projects
3. Strategic recommendations (2-3 bullet points)

Portfolio Stats: {len(all_data)} projects, Avg completion: {avg_completion:.0f}%, 
Total budget: {total_budget:,.0f}, Committed: {total_committed:,.0f}, At-risk projects: {at_risk}

Projects:
{portfolio_text}

Keep it under 120 words. Be specific and actionable."""
            }],
            temperature=0.3,
        )
        return completion.choices[0].message.content.strip()
    except Exception:
        return "Portfolio AI analysis unavailable."


# ============================================================
# SECTION 1: Portfolio Header with KPIs
# ============================================================

def make_portfolio_header(all_data):
    n = len(all_data)
    active = sum(1 for d in all_data if d.get("overall_completion", 0) > 0)
    at_risk_count = sum(1 for d in all_data if d.get("risk_score", 0) >= 50)

    total_budget = sum(d.get("effective_budget", 0) for d in all_data)
    total_committed = sum(d.get("committed_spend", 0) for d in all_data)
    total_packages = sum(d.get("total_packages", 0) for d in all_data)
    total_pkg_done = sum(d.get("packages_completed", 0) for d in all_data)
    total_pos = sum(d.get("total_pos", 0) for d in all_data)
    total_delivered = sum(d.get("delivered_pos", 0) for d in all_data)

    avg_completion = sum(d.get("overall_completion", 0) for d in all_data) / n if n > 0 else 0
    portfolio_delivery_rate = (total_delivered / total_pos * 100) if total_pos > 0 else 0
    portfolio_budget_util = (total_committed / total_budget * 100) if total_budget > 0 else 0

    # Weighted completion
    weighted_completion = (
        sum(d["overall_completion"] * d["effective_budget"] for d in all_data) / total_budget
        if total_budget > 0 else avg_completion
    )

    def kpi_card(icon, value, label, sublabel="", trend_color=None):
        return html.Div([
            html.Div([
                html.Span(icon, style={"fontSize": "20px"}),
                html.Div([
                    html.P(value, style={
                        "fontSize": F["xl"], "fontWeight": "800", "color": "#fff",
                        "margin": "0", "lineHeight": "1.1",
                    }),
                    html.P(label, style={
                        "fontSize": "10px", "color": "rgba(255,255,255,0.5)", "margin": "3px 0 0 0",
                        "textTransform": "uppercase", "letterSpacing": "1px", "fontWeight": "600",
                    }),
                    html.P(sublabel, style={
                        "fontSize": "10px", "margin": "2px 0 0 0", "fontWeight": "600",
                        "color": trend_color or "rgba(255,255,255,0.35)",
                    }) if sublabel else html.Div(),
                ]),
            ], style={"display": "flex", "alignItems": "center", "gap": "10px"}),
        ], style={
            "backgroundColor": "rgba(255,255,255,0.06)",
            "padding": "14px 18px", "borderRadius": "12px",
            "border": "1px solid rgba(255,255,255,0.07)",
            "flex": "1", "minWidth": "155px",
        })

    kpis = html.Div([
        kpi_card("📊", f"{avg_completion:.0f}%", "Avg Completion",
                 f"Weighted: {weighted_completion:.0f}%", "#a5b4fc"),
        kpi_card("💰", fmt_num(total_budget), "Total Budget",
                 f"Util: {portfolio_budget_util:.0f}%",
                 "#fca5a5" if portfolio_budget_util > 90 else "#86efac"),
        kpi_card("📦", f"{int(total_pkg_done)}/{int(total_packages)}", "Packages",
                 f"{total_pkg_done/total_packages*100:.0f}% Done" if total_packages > 0 else ""),
        kpi_card("🚚", f"{int(total_delivered)}/{int(total_pos)}", "Deliveries",
                 f"Rate: {portfolio_delivery_rate:.0f}%",
                 "#fca5a5" if portfolio_delivery_rate < 50 else "#86efac"),
        kpi_card("⚡", str(active), "Active",
                 f"{at_risk_count} at risk",
                 "#fca5a5" if at_risk_count > 0 else "#86efac"),
    ], style={"display": "flex", "gap": "10px", "flexWrap": "wrap"})

    # Project mini cards
    project_minis = []
    for i, d in enumerate(all_data):
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        oc = status_color(d["overall_completion"])
        rc = risk_color(d["risk_score"])
        uc = urgency_color(d.get("urgency_category", "TBD"))

        project_minis.append(html.Div([
            html.Div(style={
                "width": "6px", "height": "6px", "borderRadius": "50%",
                "backgroundColor": pc["main"], "flexShrink": "0",
            }),
            html.Div([
                html.Span(d["project_name"], style={
                    "fontSize": F["sm"], "fontWeight": "600", "color": "#e2e8f0",
                }),
                html.Div([
                    html.Span(f"{d['overall_completion']:.0f}%", style={
                        "fontSize": "10px", "fontWeight": "800", "color": oc,
                        "backgroundColor": hex_to_rgba(oc, 0.15),
                        "padding": "1px 7px", "borderRadius": "5px",
                    }),
                    html.Span(d.get("urgency_category", "TBD"), style={
                        "fontSize": "9px", "fontWeight": "700", "color": uc,
                        "backgroundColor": hex_to_rgba(uc, 0.15),
                        "padding": "1px 7px", "borderRadius": "5px",
                    }),
                ], style={"display": "flex", "gap": "4px", "marginTop": "3px"}),
            ]),
        ], style={
            "display": "flex", "alignItems": "center", "gap": "8px",
            "backgroundColor": "rgba(255,255,255,0.05)",
            "padding": "8px 14px", "borderRadius": "9px",
            "border": "1px solid rgba(255,255,255,0.05)",
        }))

    return html.Div([
        html.Div([
            html.Div([
                html.H2("Portfolio Overview", style={
                    "margin": "0", "color": "#fff", "fontSize": F["xxl"],
                    "fontWeight": "800", "letterSpacing": "-0.5px",
                }),
                html.P(
                    f"{n} Projects  ·  {active} Active  ·  {at_risk_count} At Risk  ·  {datetime.now().strftime('%d %b %Y')}",
                    style={
                        "margin": "5px 0 0 0", "color": "rgba(255,255,255,0.4)",
                        "fontSize": F["sm"], "fontWeight": "500",
                    }
                ),
            ]),
            html.Div(project_minis, style={
                "display": "flex", "gap": "8px", "flexWrap": "wrap",
                "justifyContent": "flex-end", "flex": "1",
            }),
        ], style={
            "display": "flex", "justifyContent": "space-between",
            "alignItems": "center", "gap": "20px",
            "flexWrap": "wrap", "marginBottom": "16px",
        }),
        kpis,
    ], style={
        "background": f"linear-gradient(135deg, {THEME['dark']} 0%, {THEME['dark2']} 50%, rgba(99,102,241,0.1) 100%)",
        "padding": "24px 28px", "borderRadius": "16px",
        "margin": "0 20px 18px 20px",
        "boxShadow": "0 4px 24px rgba(0,0,0,0.12)",
    })


# ============================================================
# SECTION 2: AI Portfolio Summary
# ============================================================

def make_ai_portfolio_section(all_data):
    ai_text = generate_portfolio_ai_summary(all_data)

    return html.Div([
        html.Div([
            html.Div([
                html.Span("🤖", style={"fontSize": "18px"}),
                html.Span("AI Portfolio Analysis", style={
                    "fontSize": F["md"], "fontWeight": "700", "color": THEME["text"],
                }),
            ], style={"display": "flex", "alignItems": "center", "gap": "8px", "marginBottom": "12px"}),
            html.Div(ai_text, style={
                "fontSize": F["sm"], "lineHeight": "1.7", "color": THEME["dark3"],
                "whiteSpace": "pre-line",
            }),
        ], style=glass_card({
            "padding": "20px 24px", "borderLeft": f"4px solid {THEME['indigo']}",
        })),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 3: Project Progress Analysis Charts
# ============================================================

def make_progress_analysis(all_data):
    names = [d["project_name"] for d in all_data]
    colors = [PROJECT_COLORS[i % len(PROJECT_COLORS)]["main"] for i in range(len(all_data))]

    # Chart 1: Overall Completion Horizontal Bars
    completion_fig = go.Figure()
    sorted_data = sorted(all_data, key=lambda x: x["overall_completion"])
    for i, d in enumerate(sorted_data):
        oc = status_color(d["overall_completion"])
        completion_fig.add_trace(go.Bar(
            y=[d["project_name"]], x=[d["overall_completion"]],
            orientation="h",
            marker_color=oc, marker_line=dict(width=0),
            text=f"  {d['overall_completion']:.0f}%",
            textposition="outside",
            textfont=dict(size=12, family="Inter", weight=700, color=oc),
            showlegend=False,
            hovertemplate=f"<b>{d['project_name']}</b><br>Completion: {d['overall_completion']:.1f}%<br>Status: {d['completion_status']}<extra></extra>",
        ))

    completion_fig.add_vline(x=75, line=dict(color=THEME["success"], width=1, dash="dot"),
                             annotation_text="Target 75%", annotation_position="top",
                             annotation_font=dict(size=9, color=THEME["success"]))
    completion_fig.update_layout(
        height=max(200, len(all_data) * 55 + 60),
        margin=dict(t=30, b=25, l=10, r=60),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"}, xaxis=dict(range=[0, 110], gridcolor="rgba(0,0,0,0.04)", ticksuffix="%"),
        yaxis=dict(tickfont={"size": 12}),
    )

    # Chart 2: Schedule Performance Index
    spi_fig = go.Figure()
    spi_vals = [d.get("schedule_performance_index", 1) for d in all_data]
    spi_colors = [THEME["success"] if v >= 0.9 else THEME["warning"] if v >= 0.7 else THEME["danger"] for v in spi_vals]

    spi_fig.add_trace(go.Bar(
        x=names, y=spi_vals,
        marker_color=spi_colors, marker_line=dict(width=0),
        text=[f"{v:.2f}" for v in spi_vals],
        textposition="outside",
        textfont=dict(size=11, family="Inter", weight=700),
        hovertemplate="<b>%{x}</b><br>SPI: %{y:.2f}<br>(1.0 = On Schedule)<extra></extra>",
    ))
    spi_fig.add_hline(y=1.0, line=dict(color=THEME["success"], width=1.5, dash="dash"),
                      annotation_text="On Schedule", annotation_position="top right",
                      annotation_font=dict(size=9, color=THEME["success"]))
    spi_fig.add_hline(y=0.7, line=dict(color=THEME["danger"], width=1, dash="dot"),
                      annotation_text="Critical", annotation_position="bottom right",
                      annotation_font=dict(size=9, color=THEME["danger"]))
    spi_fig.update_layout(
        height=260, margin=dict(t=30, b=35, l=45, r=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"}, yaxis=dict(gridcolor="rgba(0,0,0,0.04)", title="SPI"),
        xaxis=dict(tickfont={"size": 11}),
    )

    # Chart 3: Completion Radar
    categories = ["Budget Util", "Pkg Completion", "Delivery Rate", "Overall", "Pipeline Closure"]
    radar_fig = go.Figure()
    for i, d in enumerate(all_data):
        vals = [
            min(100, d.get("budget_utilization_pct", 0)),
            d.get("pkg_completion_pct", 0),
            d.get("po_delivery_rate", 0),
            d.get("overall_completion", 0),
            d.get("pipeline_closure_rate", 0),
        ]
        vals_c = vals + [vals[0]]
        cats_c = categories + [categories[0]]
        radar_fig.add_trace(go.Scatterpolar(
            r=vals_c, theta=cats_c, name=names[i],
            fill="toself", fillcolor=hex_to_rgba(colors[i], 0.1),
            line=dict(color=colors[i], width=2.5),
            marker=dict(size=4, color=colors[i]),
        ))
    radar_fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], tickfont={"size": 9}, gridcolor="rgba(0,0,0,0.06)"),
            angularaxis=dict(tickfont={"size": 10, "color": THEME["text_light"]}),
        ),
        height=320, margin=dict(t=40, b=25, l=55, r=55),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center", font={"size": 11}),
        font={"family": "Inter"},
    )

    def card(title, fig, min_w="360px"):
        return html.Div([
            html.H3(title, style={
                "fontSize": F["md"], "color": THEME["text"], "fontWeight": "700",
                "margin": "0 0 6px 0",
            }),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Project Progress Analysis", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("Procurement Completion", completion_fig, "340px"),
            card("Schedule Performance Index", spi_fig, "340px"),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap", "marginBottom": "14px"}),
        html.Div([
            card("Multi-Dimensional Performance Radar", radar_fig, "100%"),
        ], style={"display": "flex", "gap": "14px"}),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 4: Budget Analysis
# ============================================================

def make_budget_analysis(all_data):
    names = [d["project_name"] for d in all_data]
    colors = [PROJECT_COLORS[i % len(PROJECT_COLORS)]["main"] for i in range(len(all_data))]

    # Chart 1: Budget Comparison (Proposed vs Client vs Committed)
    budget_compare = go.Figure()
    for i, d in enumerate(all_data):
        cur = d.get("currency", "")
        budget_compare.add_trace(go.Bar(
            x=["Proposed", "Client/Effective", "Committed", "Remaining"],
            y=[d["proposed_budget"], d["effective_budget"], d["committed_spend"], d["budget_uncommitted"]],
            name=f"{names[i]}",
            marker_color=colors[i], marker_line=dict(width=0),
            text=[fmt_num(v, cur) for v in [d["proposed_budget"], d["effective_budget"], d["committed_spend"], d["budget_uncommitted"]]],
            textposition="outside", textfont={"size": 9, "family": "Inter"},
        ))
    budget_compare.update_layout(
        barmode="group", height=280,
        margin=dict(t=25, b=40, l=50, r=15),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font={"size": 11}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 10}),
        xaxis=dict(tickfont={"size": 11}),
    )

    # Chart 2: Budget Utilization Gauge
    util_fig = go.Figure()
    for i, d in enumerate(all_data):
        util = d.get("budget_utilization_pct", 0)
        util_color = THEME["danger"] if util > 95 else THEME["warning"] if util > 80 else THEME["success"]
        util_fig.add_trace(go.Bar(
            x=[names[i]], y=[util],
            marker_color=util_color, marker_line=dict(width=0),
            text=f"{util:.0f}%", textposition="outside",
            textfont=dict(size=12, weight=700, family="Inter"),
            showlegend=False,
            hovertemplate=f"<b>{names[i]}</b><br>Budget Utilization: {util:.1f}%<br>Committed: {fmt_num(d['committed_spend'], d['currency'])}<extra></extra>",
        ))

    util_fig.add_hline(y=100, line=dict(color=THEME["danger"], width=1.5, dash="dash"),
                       annotation_text="100% Budget", annotation_position="top right",
                       annotation_font=dict(size=9, color=THEME["danger"]))
    util_fig.add_hline(y=80, line=dict(color=THEME["warning"], width=1, dash="dot"),
                       annotation_text="80% Threshold", annotation_position="top right",
                       annotation_font=dict(size=9, color=THEME["warning"]))
    util_fig.update_layout(
        height=260, margin=dict(t=30, b=35, l=45, r=20),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        yaxis=dict(range=[0, max(120, max(d["budget_utilization_pct"] for d in all_data) + 20)],
                   gridcolor="rgba(0,0,0,0.04)", ticksuffix="%"),
        xaxis=dict(tickfont={"size": 11}),
    )

    # Chart 3: Savings/Overrun Diverging Bar
    variance_fig = go.Figure()
    sorted_by_var = sorted(all_data, key=lambda x: x["budget_variance_pct"])
    for d in sorted_by_var:
        color = THEME["success"] if d["budget_variance_pct"] >= 0 else THEME["danger"]
        variance_fig.add_trace(go.Bar(
            y=[d["project_name"]], x=[d["budget_variance_pct"]],
            orientation="h", marker_color=color, marker_line=dict(width=0),
            text=f"  {d['budget_variance_pct']:+.1f}%",
            textposition="outside",
            textfont=dict(size=11, family="Inter", weight=700, color=color),
            showlegend=False,
            hovertemplate=f"<b>{d['project_name']}</b><br>Variance: {fmt_num(d['budget_variance'], d['currency'])}<br>({d['budget_variance_pct']:+.1f}%)<extra></extra>",
        ))

    variance_fig.add_vline(x=0, line=dict(color=THEME["text_muted"], width=1))
    variance_fig.update_layout(
        height=max(180, len(all_data) * 50 + 50),
        margin=dict(t=25, b=25, l=10, r=70),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        xaxis=dict(gridcolor="rgba(0,0,0,0.04)", ticksuffix="%", title="Budget Variance %"),
        yaxis=dict(tickfont={"size": 11}),
    )

    def card(title, fig, min_w="340px"):
        return html.Div([
            html.H3(title, style={"fontSize": F["md"], "color": THEME["text"], "fontWeight": "700", "margin": "0 0 6px 0"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Budget Analysis", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("Budget Comparison", budget_compare),
            card("Budget Utilization %", util_fig),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap", "marginBottom": "14px"}),
        html.Div([
            card("Budget Variance (Savings ← → Overrun)", variance_fig, "100%"),
        ]),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 5: Procurement Pipeline
# ============================================================

def make_procurement_pipeline(all_data):
    names = [d["project_name"] for d in all_data]

    # Stacked bar: Packages status
    pkg_fig = go.Figure()
    pkg_fig.add_trace(go.Bar(
        x=names,
        y=[d["packages_completed"] for d in all_data],
        name="Completed", marker_color=THEME["success"],
        text=[f"{int(d['packages_completed'])}" for d in all_data],
        textposition="inside", textfont={"size": 11, "color": "#fff"},
    ))
    pkg_fig.add_trace(go.Bar(
        x=names,
        y=[d["packages_in_progress"] for d in all_data],
        name="In Progress", marker_color=THEME["warning"],
        text=[f"{int(d['packages_in_progress'])}" if d["packages_in_progress"] > 0 else "" for d in all_data],
        textposition="inside", textfont={"size": 11, "color": "#fff"},
    ))
    pkg_fig.add_trace(go.Bar(
        x=names,
        y=[d["packages_to_start"] for d in all_data],
        name="Not Started", marker_color="#e2e8f0",
        text=[f"{int(d['packages_to_start'])}" if d["packages_to_start"] > 0 else "" for d in all_data],
        textposition="inside", textfont={"size": 11, "color": "#94a3b8"},
    ))
    pkg_fig.update_layout(
        barmode="stack", height=280,
        margin=dict(t=25, b=35, l=40, r=15),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font={"size": 11}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", title="Packages"),
    )

    # Pipeline rates
    pipeline_fig = go.Figure()
    pipeline_fig.add_trace(go.Bar(
        x=names, y=[d["pipeline_initiation_rate"] for d in all_data],
        name="Initiation Rate", marker_color=THEME["info"],
        text=[f"{d['pipeline_initiation_rate']:.0f}%" for d in all_data],
        textposition="outside", textfont={"size": 10, "family": "Inter"},
    ))
    pipeline_fig.add_trace(go.Bar(
        x=names, y=[d["pipeline_closure_rate"] for d in all_data],
        name="Closure Rate", marker_color=THEME["purple"],
        text=[f"{d['pipeline_closure_rate']:.0f}%" for d in all_data],
        textposition="outside", textfont={"size": 10, "family": "Inter"},
    ))
    pipeline_fig.update_layout(
        barmode="group", height=260,
        margin=dict(t=25, b=35, l=45, r=15),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font={"size": 11}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", ticksuffix="%", range=[0, 115]),
    )

    def card(title, fig, min_w="340px"):
        return html.Div([
            html.H3(title, style={"fontSize": F["md"], "color": THEME["text"], "fontWeight": "700", "margin": "0 0 6px 0"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Procurement Pipeline", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("Package Status Breakdown", pkg_fig),
            card("Pipeline Efficiency Rates", pipeline_fig),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap"}),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 6: PO Delivery Tracking
# ============================================================

def make_delivery_analysis(all_data):
    names = [d["project_name"] for d in all_data]

    # PO Delivery chart
    po_fig = go.Figure()
    po_fig.add_trace(go.Bar(
        x=names, y=[d["total_pos"] for d in all_data],
        name="Total POs", marker_color=THEME["info"],
        text=[f"{int(d['total_pos'])}" for d in all_data],
        textposition="outside", textfont={"size": 11},
    ))
    po_fig.add_trace(go.Bar(
        x=names, y=[d["delivered_pos"] for d in all_data],
        name="Delivered", marker_color=THEME["success"],
        text=[f"{int(d['delivered_pos'])}" for d in all_data],
        textposition="outside", textfont={"size": 11},
    ))
    po_fig.add_trace(go.Bar(
        x=names, y=[d["outstanding_pos"] for d in all_data],
        name="Outstanding", marker_color=THEME["danger"],
        text=[f"{int(d['outstanding_pos'])}" for d in all_data],
        textposition="outside", textfont={"size": 11},
    ))
    po_fig.update_layout(
        barmode="group", height=280,
        margin=dict(t=25, b=35, l=45, r=15),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font={"size": 11}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)"),
    )

    # Delivery Rate + Gap
    delivery_rate_fig = go.Figure()
    for i, d in enumerate(all_data):
        dr = d["po_delivery_rate"]
        dr_color = THEME["success"] if dr >= 70 else THEME["warning"] if dr >= 40 else THEME["danger"]
        delivery_rate_fig.add_trace(go.Bar(
            x=[names[i]], y=[dr],
            marker_color=dr_color, showlegend=False,
            text=f"{dr:.0f}%", textposition="outside",
            textfont=dict(size=12, weight=700),
        ))

    # Add delivery gap as line
    delivery_rate_fig.add_trace(go.Scatter(
        x=names, y=[d["delivery_gap"] for d in all_data],
        mode="lines+markers+text", name="Delivery Gap",
        line=dict(color=THEME["orange"], width=2.5, dash="dot"),
        marker=dict(size=8, color=THEME["orange"]),
        text=[f"{d['delivery_gap']:.0f}pp" for d in all_data],
        textposition="top center", textfont=dict(size=10, color=THEME["orange"]),
        yaxis="y2",
    ))
    delivery_rate_fig.update_layout(
        height=280, margin=dict(t=30, b=35, l=45, r=45),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        yaxis=dict(title="Delivery Rate %", gridcolor="rgba(0,0,0,0.04)", ticksuffix="%"),
        yaxis2=dict(title="Delivery Gap (pp)", overlaying="y", side="right", gridcolor="rgba(0,0,0,0)"),
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font={"size": 11}),
    )

    def card(title, fig, min_w="340px"):
        return html.Div([
            html.H3(title, style={"fontSize": F["md"], "color": THEME["text"], "fontWeight": "700", "margin": "0 0 6px 0"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Purchase Order & Delivery Tracking", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("PO Status Overview", po_fig),
            card("Delivery Rate & Procurement Gap", delivery_rate_fig),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap"}),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 7: Risk Analysis
# ============================================================

def make_risk_analysis(all_data):
    names = [d["project_name"] for d in all_data]

    # Risk scores bar
    risk_fig = go.Figure()
    sorted_risk = sorted(all_data, key=lambda x: x["risk_score"], reverse=True)
    for d in sorted_risk:
        rc = risk_color(d["risk_score"])
        risk_fig.add_trace(go.Bar(
            y=[d["project_name"]], x=[d["risk_score"]],
            orientation="h", marker_color=rc, marker_line=dict(width=0),
            text=f"  {d['risk_score']}", textposition="outside",
            textfont=dict(size=12, weight=700, color=rc),
            showlegend=False,
            hovertemplate=f"<b>{d['project_name']}</b><br>Risk: {d['risk_score']}/100<br>Category: {risk_label(d['risk_score'])}<extra></extra>",
        ))
    risk_fig.add_vline(x=60, line=dict(color=THEME["danger"], width=1, dash="dash"),
                       annotation_text="High Risk", annotation_font=dict(size=9, color=THEME["danger"]))
    risk_fig.add_vline(x=30, line=dict(color=THEME["warning"], width=1, dash="dot"),
                       annotation_text="Medium", annotation_font=dict(size=9, color=THEME["warning"]))
    risk_fig.update_layout(
        height=max(200, len(all_data) * 50 + 50),
        margin=dict(t=25, b=25, l=10, r=60),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"}, xaxis=dict(range=[0, 110], gridcolor="rgba(0,0,0,0.04)"),
    )

    # Risk scatter: Completion vs Days to Opening
    scatter_fig = go.Figure()
    for i, d in enumerate(all_data):
        dt = d.get("days_to_opening")
        if dt is None:
            dt = 400
        rc = risk_color(d["risk_score"])
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        scatter_fig.add_trace(go.Scatter(
            x=[dt], y=[d["overall_completion"]],
            mode="markers+text",
            marker=dict(size=max(12, d["risk_score"] / 3), color=rc,
                        line=dict(width=2, color="#fff"), opacity=0.85),
            text=[d["project_name"]],
            textposition="top center",
            textfont=dict(size=10, family="Inter"),
            showlegend=False,
            hovertemplate=f"<b>{d['project_name']}</b><br>Days to Opening: {dt}<br>Completion: {d['overall_completion']:.0f}%<br>Risk: {d['risk_score']}<extra></extra>",
        ))

    scatter_fig.add_hline(y=50, line=dict(color=THEME["warning"], width=1, dash="dot"))
    scatter_fig.add_vline(x=90, line=dict(color=THEME["danger"], width=1, dash="dot"))

    # Danger zone annotation
    scatter_fig.add_annotation(x=45, y=25, text="⚠ DANGER ZONE",
                               font=dict(size=12, color=THEME["danger"], family="Inter"),
                               showarrow=False, bgcolor=hex_to_rgba(THEME["danger"], 0.08),
                               borderpad=6)

    scatter_fig.update_layout(
        height=320, margin=dict(t=25, b=40, l=50, r=25),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        xaxis=dict(title="Days to Opening", gridcolor="rgba(0,0,0,0.04)"),
        yaxis=dict(title="Completion %", gridcolor="rgba(0,0,0,0.04)", ticksuffix="%"),
    )

    # Concern categories
    all_categories = {}
    for d in all_data:
        for cat in d.get("concern_categories", []):
            all_categories[cat] = all_categories.get(cat, 0) + 1

    if all_categories:
        cats = sorted(all_categories.items(), key=lambda x: x[1], reverse=True)
        cat_fig = go.Figure(go.Bar(
            x=[c[1] for c in cats], y=[c[0] for c in cats],
            orientation="h", marker_color=THEME["orange"],
            text=[str(c[1]) for c in cats], textposition="outside",
            textfont=dict(size=12, weight=700),
        ))
        cat_fig.update_layout(
            height=max(150, len(cats) * 40 + 50),
            margin=dict(t=15, b=20, l=10, r=40),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font={"family": "Inter"},
            xaxis=dict(gridcolor="rgba(0,0,0,0.04)", title="Occurrences"),
        )
    else:
        cat_fig = go.Figure()
        cat_fig.update_layout(height=150, paper_bgcolor="rgba(0,0,0,0)",
                              annotations=[dict(text="No concerns categorized", x=0.5, y=0.5,
                                                xref="paper", yref="paper", showarrow=False)])

    def card(title, fig, min_w="340px"):
        return html.Div([
            html.H3(title, style={"fontSize": F["md"], "color": THEME["text"], "fontWeight": "700", "margin": "0 0 6px 0"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Risk Analysis", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("Composite Risk Scores", risk_fig, "340px"),
            card("Risk Matrix: Completion vs Timeline", scatter_fig, "400px"),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap", "marginBottom": "14px"}),
        html.Div([
            card("Concern Categories Distribution", cat_fig, "100%"),
        ]),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 8: Timeline Analysis
# ============================================================

def make_timeline_analysis(all_data):
    today = pd.Timestamp.now()
    fig = go.Figure()

    for i, d in enumerate(all_data):
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        name = d["project_name"]
        parsed = d.get("opening_date_parsed")
        opening_str = d.get("opening_date", "TBD")
        uc = urgency_color(d.get("urgency_category", "TBD"))

        if opening_str == "Opened":
            fig.add_trace(go.Scatter(
                x=[today], y=[name], mode="markers+text",
                marker=dict(size=16, color=THEME["success"], symbol="star", line=dict(width=2, color="#fff")),
                text=["  OPENED"], textposition="middle right",
                textfont=dict(size=11, color=THEME["success"], weight=700),
                showlegend=False,
            ))
        elif parsed is not None:
            days_left = (parsed - today).days
            fig.add_trace(go.Scatter(
                x=[today, parsed], y=[name, name], mode="lines",
                line=dict(color=pc["main"], width=8),
                showlegend=False, hoverinfo="skip",
            ))
            label = f"  {days_left}d" if days_left > 0 else f"  {abs(days_left)}d overdue"
            fig.add_trace(go.Scatter(
                x=[parsed], y=[name], mode="markers+text",
                marker=dict(size=14, color=uc, symbol="diamond", line=dict(width=2, color="#fff")),
                text=[label], textposition="middle right",
                textfont=dict(size=11, color=uc, weight=700),
                showlegend=False,
                hovertemplate=f"<b>{name}</b><br>Opening: {parsed.strftime('%d %b %Y')}<br>Days: {days_left}<br>Urgency: {d.get('urgency_category','')}<extra></extra>",
            ))
        else:
            fig.add_trace(go.Scatter(
                x=[today], y=[name], mode="markers+text",
                marker=dict(size=12, color=THEME["text_muted"], symbol="circle"),
                text=["  TBD"], textposition="middle right",
                textfont=dict(size=11, color=THEME["text_muted"]),
                showlegend=False,
            ))

    fig.add_vline(x=today.timestamp() * 1000,
                  line=dict(color=THEME["danger"], width=1.5, dash="dot"),
                  annotation_text="Today", annotation_position="top")
    fig.update_layout(
        height=max(180, len(all_data) * 60 + 50),
        margin=dict(t=35, b=30, l=15, r=40),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        xaxis=dict(gridcolor="rgba(0,0,0,0.04)"),
        yaxis=dict(gridcolor="rgba(0,0,0,0.02)"),
    )

    # Urgency distribution
    urgency_counts = {}
    for d in all_data:
        cat = d.get("urgency_category", "TBD")
        urgency_counts[cat] = urgency_counts.get(cat, 0) + 1

    urg_fig = go.Figure(go.Pie(
        labels=list(urgency_counts.keys()),
        values=list(urgency_counts.values()),
        marker=dict(colors=[urgency_color(k) for k in urgency_counts.keys()]),
        hole=0.5,
        textinfo="label+value",
        textfont=dict(size=12, family="Inter"),
    ))
    urg_fig.update_layout(
        height=250, margin=dict(t=20, b=20, l=20, r=20),
        paper_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"},
        showlegend=False,
    )

    def card(title, fig, min_w="340px"):
        return html.Div([
            html.H3(title, style={"fontSize": F["md"], "color": THEME["text"], "fontWeight": "700", "margin": "0 0 6px 0"}),
            dcc.Graph(figure=fig, config={"displayModeBar": False}),
        ], style=glass_card({"padding": "18px 20px", "flex": "1", "minWidth": min_w}))

    return html.Div([
        html.H2("Timeline Analysis", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            card("Project Opening Timeline", fig, "55%"),
            card("Urgency Distribution", urg_fig, "280px"),
        ], style={"display": "flex", "gap": "14px", "flexWrap": "wrap"}),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 9: Summary Table
# ============================================================

def make_data_table(all_data):
    rows = []
    for d in all_data:
        rows.append({
            "Project": d.get("project_name", ""),
            "Stage": d.get("current_stage", ""),
            "Opening": d.get("opening_date", "TBD"),
            "Urgency": d.get("urgency_category", "TBD"),
            "Currency": d.get("currency", ""),
            "Budget": f"{d['effective_budget']:,.0f}",
            "Committed": f"{d['committed_spend']:,.0f}",
            "Util %": f"{d['budget_utilization_pct']:.0f}%",
            "Variance": f"{d['budget_variance']:+,.0f}",
            "Pkgs": f"{int(d['packages_completed'])}/{int(d['total_packages'])}",
            "POs Del": f"{int(d['delivered_pos'])}/{int(d['total_pos'])}",
            "Del Rate": f"{d['po_delivery_rate']:.0f}%",
            "Completion": f"{d['overall_completion']:.0f}%",
            "SPI": f"{d.get('schedule_performance_index', 1):.2f}",
            "Risk": f"{d['risk_score']}",
        })
    df_table = pd.DataFrame(rows)

    style_data_conditional = [
        {"if": {"row_index": "odd"}, "backgroundColor": "rgba(248,250,252,0.7)"},
    ]

    for i, d in enumerate(all_data):
        rc = risk_color(d["risk_score"])
        style_data_conditional.append({
            "if": {"row_index": i, "column_id": "Risk"},
            "color": rc, "fontWeight": "700",
        })
        vc = THEME["success"] if d["budget_variance"] >= 0 else THEME["danger"]
        style_data_conditional.append({
            "if": {"row_index": i, "column_id": "Variance"},
            "color": vc, "fontWeight": "700",
        })
        sc = status_color(d["overall_completion"])
        style_data_conditional.append({
            "if": {"row_index": i, "column_id": "Completion"},
            "color": sc, "fontWeight": "700",
        })
        uc = urgency_color(d.get("urgency_category", "TBD"))
        style_data_conditional.append({
            "if": {"row_index": i, "column_id": "Urgency"},
            "color": uc, "fontWeight": "700",
        })

    return html.Div([
        html.H2("Project Summary Table", style={
            "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
            "margin": "0 0 16px 0", "paddingBottom": "10px",
            "borderBottom": f"2px solid {THEME['border']}",
        }),
        html.Div([
            dash_table.DataTable(
                data=df_table.to_dict("records"),
                columns=[{"name": c, "id": c} for c in df_table.columns],
                style_table={"overflowX": "auto", "borderRadius": "12px"},
                style_header={
                    "backgroundColor": THEME["dark"], "color": "#fff",
                    "fontWeight": "700", "fontSize": F["sm"],
                    "textAlign": "center", "padding": "12px 8px",
                    "fontFamily": "Inter", "borderBottom": "none",
                },
                style_cell={
                    "textAlign": "center", "padding": "10px 8px",
                    "fontSize": F["sm"], "fontFamily": "Inter",
                    "border": "none",
                    "borderBottom": f"1px solid {THEME['border']}",
                    "minWidth": "70px", "color": THEME["text"],
                },
                style_data_conditional=style_data_conditional,
                sort_action="native",
                filter_action="native",
                page_size=20,
            )
        ], style=glass_card({"overflow": "hidden", "padding": "0"})),
    ], style={"padding": "0 20px 16px 20px"})


# ============================================================
# SECTION 10: Project Detail Cards
# ============================================================

def make_progress_bar(label, value, max_val, color, show_fraction=True):
    pct = (value / max_val * 100) if max_val > 0 else 0
    pct = min(100, pct)
    right_text = f"{int(value)}/{int(max_val)}" if show_fraction else f"{pct:.0f}%"

    return html.Div([
        html.Div([
            html.Span(label, style={"fontSize": F["sm"], "color": THEME["text"], "fontWeight": "600"}),
            html.Span(right_text, style={"fontSize": F["sm"], "color": color, "fontWeight": "700"}),
        ], style={"display": "flex", "justifyContent": "space-between", "marginBottom": "5px"}),
        html.Div([
            html.Div(style={
                "width": f"{pct}%", "height": "100%",
                "backgroundColor": color, "borderRadius": "5px",
                "transition": "width 0.5s ease",
            })
        ], style={
            "width": "100%", "height": "7px",
            "backgroundColor": "#f1f5f9", "borderRadius": "5px", "overflow": "hidden",
        }),
    ], style={"marginBottom": "12px"})


def make_project_card(data, pc_idx):
    pc = PROJECT_COLORS[pc_idx % len(PROJECT_COLORS)]
    name = data["project_name"]
    stage = data["current_stage"]
    currency = data["currency"]
    opening = data["opening_date"]
    budget = data["effective_budget"]
    overall = data["overall_completion"]
    risk_score = data["risk_score"]
    concerns = data["concerns"]

    oc = status_color(overall)
    rc = risk_color(risk_score)
    uc = urgency_color(data.get("urgency_category", "TBD"))

    ai_summary = generate_ai_summary(data)

    # Gauge
    gauge = go.Figure(go.Indicator(
        mode="gauge+number", value=round(overall, 1),
        number={"suffix": "%", "font": {"size": 30, "color": pc["main"], "family": "Inter"}},
        gauge={
            "axis": {"range": [0, 100], "tickfont": {"size": 9}, "dtick": 25},
            "bar": {"color": oc, "thickness": 0.7},
            "bgcolor": "#f1f5f9", "borderwidth": 0,
            "steps": [
                {"range": [0, 40], "color": hex_to_rgba(THEME["danger"], 0.05)},
                {"range": [40, 75], "color": hex_to_rgba(THEME["warning"], 0.05)},
                {"range": [75, 100], "color": hex_to_rgba(THEME["success"], 0.05)},
            ],
        },
    ))
    gauge.update_layout(height=130, margin=dict(t=8, b=0, l=15, r=15), paper_bgcolor="rgba(0,0,0,0)")

    # Budget waterfall
    ordered = data["orders_placed"]
    in_prog = data["orders_in_progress"]
    remaining = max(0, budget - ordered - in_prog)
    wf = go.Figure(go.Waterfall(
        orientation="v",
        x=["Budget", "Ordered", "In Progress", "Remaining"],
        y=[budget, -ordered, -in_prog, 0],
        measure=["absolute", "relative", "relative", "total"],
        text=[fmt_num(v, currency) for v in [budget, ordered, in_prog, remaining]],
        textposition="outside", textfont=dict(size=10, family="Inter"),
        connector={"line": {"color": "rgba(0,0,0,0.06)", "width": 1}},
        increasing={"marker": {"color": pc["main"]}},
        decreasing={"marker": {"color": THEME["warning"]}},
        totals={"marker": {"color": THEME["success"] if remaining >= 0 else THEME["danger"]}},
    ))
    wf.update_layout(
        height=190, margin=dict(t=18, b=28, l=35, r=12),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter"}, showlegend=False,
        yaxis=dict(gridcolor="rgba(0,0,0,0.03)", tickfont={"size": 9}),
        xaxis=dict(tickfont={"size": 10}),
    )

    # KPI chips
    def chip(label, value, icon, color):
        return html.Div([
            html.Span(icon, style={"fontSize": "14px"}),
            html.Div([
                html.P(label, style={"fontSize": "9px", "color": THEME["text_muted"], "margin": "0", "textTransform": "uppercase", "letterSpacing": "0.5px"}),
                html.P(value, style={"fontSize": F["md"], "color": THEME["text"], "margin": "1px 0 0 0", "fontWeight": "700"}),
            ]),
        ], style={
            "display": "flex", "alignItems": "center", "gap": "7px",
            "backgroundColor": hex_to_rgba(color, 0.05),
            "padding": "8px 10px", "borderRadius": "9px",
            "flex": "1", "minWidth": "110px",
        })

    sav = data["savings_overrun"]
    sav_color = THEME["success"] if sav >= 0 else THEME["danger"]

    # Badges
    proc_started = data.get("proc_started", "").lower() == "yes"
    del_started = data.get("delivery_started", "").lower() == "yes"

    def badge(text, active):
        color = THEME["success"] if active else THEME["text_muted"]
        return html.Span(f"{'✓' if active else '○'} {text}", style={
            "fontSize": "9px", "fontWeight": "700", "color": color,
            "backgroundColor": hex_to_rgba(color, 0.08),
            "padding": "3px 8px", "borderRadius": "12px",
        })

    def section_label(text):
        return html.H4(text, style={
            "fontSize": F["sm"], "color": THEME["text"], "fontWeight": "700",
            "margin": "14px 0 8px 0",
        })

    # Concern categories badges
    cat_badges = []
    for cat in data.get("concern_categories", []):
        cat_colors = {"Supplier": THEME["purple"], "Timeline": THEME["danger"],
                      "Cost": THEME["orange"], "Quality": THEME["pink"],
                      "Approval": THEME["warning"], "Logistics": THEME["cyan"]}
        cc = cat_colors.get(cat, THEME["text_muted"])
        cat_badges.append(html.Span(cat, style={
            "fontSize": "9px", "fontWeight": "700", "color": cc,
            "backgroundColor": hex_to_rgba(cc, 0.1),
            "padding": "2px 7px", "borderRadius": "8px",
        }))

    return html.Div([
        # Header
        html.Div([
            html.Div([
                html.H3(name, style={"margin": "0", "color": "#fff", "fontSize": F["lg"], "fontWeight": "800"}),
                html.P(stage, style={"margin": "3px 0 0 0", "color": "rgba(255,255,255,0.7)", "fontSize": F["xs"]}),
            ]),
            html.Div([
                html.Span(f"{risk_score}", style={
                    "fontSize": F["lg"], "fontWeight": "800", "color": "#fff",
                }),
                html.Span("Risk", style={
                    "fontSize": "9px", "color": "rgba(255,255,255,0.6)",
                    "textTransform": "uppercase",
                }),
            ], style={
                "display": "flex", "flexDirection": "column", "alignItems": "center",
                "backgroundColor": hex_to_rgba(rc, 0.3),
                "padding": "6px 14px", "borderRadius": "10px",
                "border": "1px solid rgba(255,255,255,0.2)",
            }),
        ], style={
            "display": "flex", "justifyContent": "space-between", "alignItems": "center",
            "background": pc["gradient"], "padding": "16px 20px",
            "borderRadius": "14px 14px 0 0",
        }),

        # Body
        html.Div([
            # Status row
            html.Div([
                badge("Procurement", proc_started),
                badge("Delivery", del_started),
                html.Span(f"📅 {opening}", style={
                    "fontSize": "9px", "fontWeight": "600", "color": THEME["info"],
                    "backgroundColor": hex_to_rgba(THEME["info"], 0.07),
                    "padding": "3px 8px", "borderRadius": "12px",
                }),
                html.Span(data.get("urgency_category", "TBD"), style={
                    "fontSize": "9px", "fontWeight": "700", "color": uc,
                    "backgroundColor": hex_to_rgba(uc, 0.1),
                    "padding": "3px 8px", "borderRadius": "12px",
                }),
            ], style={"display": "flex", "gap": "5px", "flexWrap": "wrap", "marginBottom": "12px"}),

            # KPIs
            html.Div([
                chip("Budget", fmt_num(budget, currency), "💰", pc["main"]),
                chip("Savings" if sav >= 0 else "Overrun", fmt_num(abs(sav), currency), "📊", sav_color),
                chip("Util", f"{data['budget_utilization_pct']:.0f}%", "📈", THEME["info"]),
            ], style={"display": "flex", "gap": "6px", "flexWrap": "wrap", "marginBottom": "14px"}),

            # Gauge
            section_label("Overall Completion"),
            html.Div(
                dcc.Graph(figure=gauge, config={"displayModeBar": False}),
                style={"backgroundColor": "#fafbfc", "borderRadius": "10px", "padding": "4px 0", "marginBottom": "12px"},
            ),

            # Progress bars
            section_label("Progress Tracking"),
            html.Div([
                make_progress_bar("Packages Completed", data["packages_completed"], data["total_packages"], THEME["success"]),
                make_progress_bar("Packages In Progress", data["packages_in_progress"], data["total_packages"], THEME["warning"]),
                make_progress_bar("POs Delivered", data["delivered_pos"], data["total_pos"], THEME["info"]),
            ], style={"backgroundColor": "#fafbfc", "padding": "14px", "borderRadius": "10px", "marginBottom": "12px"}),

            # Advanced metrics
            section_label("Advanced Metrics"),
            html.Div([
                html.Div([
                    html.Div([
                        html.Span("SPI", style={"fontSize": "10px", "color": THEME["text_muted"], "fontWeight": "600"}),
                        html.Span(f"{data.get('schedule_performance_index', 1):.2f}", style={
                            "fontSize": F["md"], "fontWeight": "800",
                            "color": THEME["success"] if data.get("schedule_performance_index", 1) >= 0.9 else THEME["danger"],
                        }),
                    ], style={"textAlign": "center", "flex": "1"}),
                    html.Div(style={"width": "1px", "backgroundColor": THEME["border"], "alignSelf": "stretch"}),
                    html.Div([
                        html.Span("Pipeline", style={"fontSize": "10px", "color": THEME["text_muted"], "fontWeight": "600"}),
                        html.Span(f"{data['pipeline_closure_rate']:.0f}%", style={
                            "fontSize": F["md"], "fontWeight": "800", "color": THEME["purple"],
                        }),
                    ], style={"textAlign": "center", "flex": "1"}),
                    html.Div(style={"width": "1px", "backgroundColor": THEME["border"], "alignSelf": "stretch"}),
                    html.Div([
                        html.Span("Del. Rate", style={"fontSize": "10px", "color": THEME["text_muted"], "fontWeight": "600"}),
                        html.Span(f"{data['po_delivery_rate']:.0f}%", style={
                            "fontSize": F["md"], "fontWeight": "800",
                            "color": THEME["success"] if data["po_delivery_rate"] >= 60 else THEME["danger"],
                        }),
                    ], style={"textAlign": "center", "flex": "1"}),
                    html.Div(style={"width": "1px", "backgroundColor": THEME["border"], "alignSelf": "stretch"}),
                    html.Div([
                        html.Span("Del. Gap", style={"fontSize": "10px", "color": THEME["text_muted"], "fontWeight": "600"}),
                        html.Span(f"{data['delivery_gap']:.0f}pp", style={
                            "fontSize": F["md"], "fontWeight": "800", "color": THEME["orange"],
                        }),
                    ], style={"textAlign": "center", "flex": "1"}),
                ], style={
                    "display": "flex", "gap": "0", "padding": "12px 8px",
                    "backgroundColor": "#fafbfc", "borderRadius": "10px",
                }),
            ], style={"marginBottom": "12px"}),

            # Budget waterfall
            section_label("Budget Breakdown"),
            html.Div(
                dcc.Graph(figure=wf, config={"displayModeBar": False}),
                style={"backgroundColor": "#fafbfc", "borderRadius": "10px", "marginBottom": "12px"},
            ),

            # Concerns
            section_label("Concerns & Risk Categories"),
            html.Div(cat_badges, style={"display": "flex", "gap": "4px", "flexWrap": "wrap", "marginBottom": "8px"}) if cat_badges else html.Div(),
            html.Div([
                html.Div(c, style={
                    "padding": "9px 12px", "backgroundColor": "#fffbeb",
                    "borderRadius": "8px", "fontSize": F["sm"],
                    "borderLeft": f"3px solid {THEME['warning']}",
                    "marginBottom": "5px", "color": "#92400e", "lineHeight": "1.5",
                }) for c in concerns
            ]) if concerns else html.P("No concerns reported", style={
                "color": THEME["success"], "fontSize": F["sm"], "margin": "0",
                "padding": "9px 12px", "backgroundColor": hex_to_rgba(THEME["success"], 0.05),
                "borderRadius": "8px", "borderLeft": f"3px solid {THEME['success']}",
            }),

            # AI Summary
            html.Div([
                html.H4("AI Analysis", style={
                    "fontSize": F["sm"], "color": THEME["text"], "fontWeight": "700",
                    "margin": "14px 0 8px 0",
                }),
                html.Div(ai_summary, style={
                    "backgroundColor": "#eef2ff", "padding": "12px 14px",
                    "borderRadius": "8px", "fontSize": F["sm"],
                    "color": "#3730a3", "borderLeft": "3px solid #6366f1",
                    "lineHeight": "1.6",
                }),
            ]),
        ], style={"padding": "16px 20px"}),
    ], style=glass_card({
        "overflow": "hidden", "flex": "1",
        "minWidth": "400px", "maxWidth": "560px",
    }))


# ============================================================
# DASH APP
# ============================================================

app = dash.Dash(
    __name__,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1.0"}],
    suppress_callback_exceptions=True,
)

FULLSCREEN_BTN = {
    "backgroundColor": "rgba(99,102,241,0.15)", "color": "#a5b4fc",
    "border": "1px solid rgba(99,102,241,0.3)",
    "padding": "8px 18px", "borderRadius": "9px", "cursor": "pointer",
    "fontSize": F["sm"], "fontWeight": "600",
    "display": "flex", "alignItems": "center", "gap": "6px",
}

EXIT_FS_BTN = {
    "backgroundColor": "rgba(239,68,68,0.15)", "color": "#fca5a5",
    "border": "1px solid rgba(239,68,68,0.3)",
    "padding": "8px 18px", "borderRadius": "9px", "cursor": "pointer",
    "fontSize": F["sm"], "fontWeight": "600",
    "display": "none", "alignItems": "center", "gap": "6px",
}

app.layout = html.Div([
    # Upload
    html.Div(id="upload-section", children=[
        html.Div(style={"height": "12vh"}),
        html.Div([
            html.Div([
                html.Div(style={
                    "width": "68px", "height": "68px", "borderRadius": "18px",
                    "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                    "display": "flex", "alignItems": "center", "justifyContent": "center",
                    "margin": "0 auto 18px auto", "overflow": "hidden",
                    "boxShadow": "0 8px 24px rgba(99,102,241,0.3)",
                }, children=[
                    html.Img(src="/assets/luxurylogo.jpg", style={"height": "68px", "width": "68px", "objectFit": "cover"})
                ]),
                html.H1("Luxury Hospitality Dashboard", style={
                    "margin": "0 0 6px 0", "color": THEME["text"],
                    "fontSize": "28px", "fontWeight": "800", "letterSpacing": "-0.5px",
                }),
                html.P("Upload your procurement data to generate insights", style={
                    "color": THEME["text_muted"], "fontSize": F["md"],
                    "margin": "0 0 28px 0",
                }),
            ]),
            dcc.Upload(
                id="upload-data",
                children=html.Div([
                    html.Div(style={
                        "width": "48px", "height": "48px", "borderRadius": "12px",
                        "backgroundColor": "#eef2ff", "display": "flex",
                        "alignItems": "center", "justifyContent": "center",
                        "margin": "0 auto 12px auto",
                    }, children=[html.Span("📁", style={"fontSize": "22px"})]),
                    html.P("Drag & Drop or Click to Upload", style={
                        "fontWeight": "700", "color": "#6366f1", "fontSize": F["lg"], "margin": "0 0 4px 0",
                    }),
                    html.P(".xlsx  ·  .xlsm  ·  .xls  ·  .csv", style={
                        "color": THEME["text_muted"], "fontSize": F["sm"], "margin": "0",
                    }),
                ]),
                style={
                    "width": "100%", "padding": "36px 20px",
                    "borderWidth": "2px", "borderStyle": "dashed",
                    "borderColor": "#c7d2fe", "borderRadius": "14px",
                    "textAlign": "center", "backgroundColor": "rgba(250,251,255,0.8)",
                    "cursor": "pointer",
                },
                multiple=False,
            ),
            html.Div(id="upload-error", style={"marginTop": "14px"}),
        ], style=glass_card({
            "padding": "44px 52px", "maxWidth": "480px",
            "margin": "0 auto", "textAlign": "center",
        })),
    ]),

    # Dashboard
    html.Div(id="dashboard-section", style={"display": "none"}, children=[
        # Top bar
        html.Div([
            html.Div([
                html.Div(style={
                    "width": "36px", "height": "36px", "borderRadius": "9px",
                    "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                    "overflow": "hidden",
                }, children=[
                    html.Img(src="/assets/luxurylogo.jpg", style={"height": "36px", "width": "36px", "objectFit": "cover"})
                ]),
                html.Div([
                    html.H1("Luxury Hospitality Dashboard", style={
                        "margin": "0", "color": "#fff", "fontSize": F["lg"],
                        "fontWeight": "800", "letterSpacing": "-0.3px",
                    }),
                    html.P("Procurement Analytics & Risk Intelligence", style={
                        "color": "rgba(255,255,255,0.4)", "fontSize": F["xs"],
                        "margin": "1px 0 0 0",
                    }),
                ]),
            ], style={"display": "flex", "alignItems": "center", "gap": "10px", "flex": "1"}),
            html.Div([
                html.Button([html.Span("⛶", style={"fontSize": "14px"}), html.Span("Fullscreen")],
                            id="btn-fullscreen", n_clicks=0, style=FULLSCREEN_BTN),
                html.Button([html.Span("✕", style={"fontSize": "12px"}), html.Span("Exit")],
                            id="btn-exit-fullscreen", n_clicks=0, style=EXIT_FS_BTN),
                html.A("← New File", href="/", style={
                    "backgroundColor": "rgba(255,255,255,0.08)", "color": "#e2e8f0",
                    "border": "1px solid rgba(255,255,255,0.1)",
                    "padding": "8px 18px", "borderRadius": "9px",
                    "fontSize": F["sm"], "fontWeight": "600", "textDecoration": "none",
                }),
            ], style={"display": "flex", "alignItems": "center", "gap": "6px"}),
        ], id="top-bar", style={
            "display": "flex", "justifyContent": "space-between", "alignItems": "center",
            "background": f"linear-gradient(135deg, {THEME['dark']}, {THEME['dark2']})",
            "padding": "12px 24px", "marginBottom": "18px",
            "boxShadow": "0 2px 12px rgba(0,0,0,0.1)",
            "position": "sticky", "top": "0", "zIndex": "1000",
        }),
        html.Div(id="dashboard-body"),
    ]),
], id="main-container", style={
    "fontFamily": "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
    "backgroundImage": "url('/assets/background.jpg')",
    "backgroundSize": "cover", "backgroundPosition": "center",
    "backgroundRepeat": "no-repeat", "backgroundAttachment": "fixed",
    "minHeight": "100vh",
})


# ============================================================
# Fullscreen Callback
# ============================================================

app.clientside_callback(
    """
    function(n1, n2) {
        var elem = document.getElementById('main-container');
        var triggered = dash_clientside.callback_context.triggered;
        if (!triggered || triggered.length === 0) return [window.dash_clientside.no_update, window.dash_clientside.no_update];
        var id = triggered[0].prop_id.split('.')[0];
        var enterBtn = document.getElementById('btn-fullscreen');
        var exitBtn = document.getElementById('btn-exit-fullscreen');
        if (id === 'btn-fullscreen') {
            if (elem.requestFullscreen) elem.requestFullscreen();
            else if (elem.webkitRequestFullscreen) elem.webkitRequestFullscreen();
            elem.style.overflow = 'auto'; elem.style.height = '100vh';
            if (enterBtn) enterBtn.style.display = 'none';
            if (exitBtn) exitBtn.style.display = 'flex';
        } else {
            if (document.exitFullscreen) document.exitFullscreen();
            else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
            elem.style.overflow = ''; elem.style.height = '';
            if (enterBtn) enterBtn.style.display = 'flex';
            if (exitBtn) exitBtn.style.display = 'none';
        }
        document.onfullscreenchange = function() {
            var e = document.getElementById('btn-fullscreen');
            var x = document.getElementById('btn-exit-fullscreen');
            var m = document.getElementById('main-container');
            if (!document.fullscreenElement) {
                if (e) e.style.display='flex'; if (x) x.style.display='none';
                m.style.overflow=''; m.style.height='';
            } else { m.style.overflow='auto'; m.style.height='100vh'; }
        };
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
# Upload Callback
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

    ext = filename.lower().rsplit(".", 1)[-1] if filename and "." in filename else ""
    if ext not in ["xlsx", "xlsm", "xls", "csv"]:
        return dash.no_update, dash.no_update, dash.no_update, html.P(
            f"Unsupported: .{ext}", style={"color": THEME["danger"], "fontWeight": "600"})

    try:
        all_data = parse_file(contents, filename)
    except Exception as e:
        return dash.no_update, dash.no_update, dash.no_update, html.P(
            f"Error: {str(e)}", style={"color": THEME["danger"]})

    if not all_data:
        return dash.no_update, dash.no_update, dash.no_update, html.P(
            "No valid project data found.", style={"color": THEME["danger"]})

    # Build all sections
    body = html.Div([
        # 1. Portfolio Header
        make_portfolio_header(all_data),

        # 2. AI Portfolio Summary
        make_ai_portfolio_section(all_data),

        # 3. Progress Analysis
        make_progress_analysis(all_data),

        # 4. Budget Analysis
        make_budget_analysis(all_data),

        # 5. Procurement Pipeline
        make_procurement_pipeline(all_data),

        # 6. PO & Delivery
        make_delivery_analysis(all_data),

        # 7. Risk Analysis
        make_risk_analysis(all_data),

        # 8. Timeline
        make_timeline_analysis(all_data),

        # 9. Summary Table
        make_data_table(all_data),

        # 10. Project Detail Cards
        html.Div([
            html.H2("Project Detail Cards", style={
                "fontSize": F["lg"], "color": THEME["text"], "fontWeight": "700",
                "margin": "0 0 16px 0", "paddingBottom": "10px",
                "borderBottom": f"2px solid {THEME['border']}",
            }),
        ], style={"padding": "0 20px"}),

        html.Div(
            [make_project_card(d, i) for i, d in enumerate(all_data)],
            style={
                "display": "flex", "justifyContent": "center",
                "alignItems": "flex-start", "gap": "18px",
                "flexWrap": "wrap", "padding": "0 20px 40px 20px",
            },
        ),
    ])

    return {"display": "none"}, {"display": "block"}, body, ""


# ============================================================
# Run
# ============================================================

if __name__ == "__main__":
    port = 8050
    url = f"http://127.0.0.1:{port}"
    print(f"\n{'='*50}\n  Dashboard: {url}\n{'='*50}\n")
    Timer(2.0, open_browser, args=[url]).start()
    app.run(debug=False, port=port, host="127.0.0.1")
