import pandas as pd # type: ignore
import dash # type: ignore
from dash import dcc, html, Input, Output, State, dash_table # type: ignore
import plotly.graph_objects as go # type: ignore
import io
import base64
import subprocess
import platform
from threading import Timer

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
            else:
                try:
                    parsed = pd.to_datetime(od, dayfirst=True)
                    data["opening_date"] = parsed.strftime("%Y-%m-%d")
                except:
                    data["opening_date"] = od
        else:
            data["opening_date"] = "TBD"

        data["proposed_budget"] = safe_float(row.get(col_map.get("proposed_budget", ""), 0))
        data["client_budget"] = safe_float(row.get(col_map.get("client_budget", ""), 0))
        data["orders_placed"] = safe_float(row.get(col_map.get("orders_placed", ""), 0))
        data["orders_in_progress"] = safe_float(row.get(col_map.get("orders_in_progress", ""), 0))
        data["currency"] = safe_str(row.get(col_map.get("currency", ""), "USD")) or "USD"
        data["proc_started"] = safe_str(row.get(col_map.get("proc_started", ""), ""))
        data["total_packages"] = safe_float(row.get(col_map.get("total_packages", ""), 0))
        data["packages_completed"] = safe_float(row.get(col_map.get("packages_completed", ""), 0))
        data["packages_in_progress"] = safe_float(row.get(col_map.get("packages_in_progress", ""), 0))
        data["packages_to_start"] = max(0, data["total_packages"] - data["packages_completed"] - data["packages_in_progress"])
        data["delivery_started"] = safe_str(row.get(col_map.get("delivery_started", ""), ""))
        data["total_pos"] = safe_float(row.get(col_map.get("total_pos", ""), 0))
        data["delivered_pos"] = safe_float(row.get(col_map.get("delivered_pos", ""), 0))
        data["delivery_in_progress"] = max(0, data["total_pos"] - data["delivered_pos"])

        overall_proc_val = safe_float(row.get(col_map.get("overall_proc", ""), 0))
        data["overall_proc_from_file"] = overall_proc_val * 100 if overall_proc_val <= 1 else overall_proc_val

        budget = data["client_budget"] if data["client_budget"] > 0 else data["proposed_budget"]
        data["orders_placed_pct"] = (data["orders_placed"] / budget * 100) if budget > 0 else 0

        pkg_completion = (data["packages_completed"] / data["total_packages"] * 100) if data["total_packages"] > 0 else 0
        delivery_completion = (data["delivered_pos"] / data["total_pos"] * 100) if data["total_pos"] > 0 else 0

        data["overall_completion"] = data["overall_proc_from_file"]
        data["pkg_completion_pct"] = pkg_completion
        data["delivery_completion_pct"] = delivery_completion

        if budget > 0 and data["orders_placed"] > 0:
            data["savings_overrun"] = budget - data["orders_placed"] - data["orders_in_progress"]
        else:
            data["savings_overrun"] = 0

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
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    buf = io.BytesIO(decoded)
    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''

    all_projects = []
    if ext == 'csv':
        df = pd.read_csv(buf)
        all_projects.extend(extract_projects_from_sheet(df))
    elif ext in ['xlsx', 'xlsm', 'xls']:
        engine = 'openpyxl' if ext in ['xlsx', 'xlsm'] else 'xlrd'
        xls = pd.ExcelFile(buf, engine=engine)
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.dropna(how='all').empty:
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
    except:
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
}

PROJECT_COLORS = [
    {"main": "#6366f1", "light": "#eef2ff", "gradient": "linear-gradient(135deg, #6366f1, #8b5cf6)"},
    {"main": "#0ea5e9", "light": "#f0f9ff", "gradient": "linear-gradient(135deg, #0ea5e9, #06b6d4)"},
    {"main": "#f97316", "light": "#fff7ed", "gradient": "linear-gradient(135deg, #f97316, #fb923c)"},
    {"main": "#ec4899", "light": "#fdf2f8", "gradient": "linear-gradient(135deg, #ec4899, #f472b6)"},
    {"main": "#14b8a6", "light": "#f0fdfa", "gradient": "linear-gradient(135deg, #14b8a6, #2dd4bf)"},
]

# Font sizes
F = {
    "xs": "11px",
    "sm": "13px",
    "md": "15px",
    "lg": "18px",
    "xl": "22px",
    "xxl": "28px",
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


def status_color(pct):
    if pct >= 75:
        return THEME["success"]
    elif pct >= 40:
        return THEME["warning"]
    return THEME["danger"]


# ============================================================
# Glass card style helper
# ============================================================

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
# Portfolio Summary Header
# ============================================================

def make_portfolio_header(all_data):
    n = len(all_data)
    active = sum(1 for d in all_data if d.get("overall_completion", 0) > 0)
    planning = n - active

    project_pills = []
    for i, d in enumerate(all_data):
        pc = PROJECT_COLORS[i % len(PROJECT_COLORS)]
        overall = d.get("overall_completion", 0)
        sc = status_color(overall)

        project_pills.append(html.Div([
            html.Div(style={
                "width": "10px", "height": "10px", "borderRadius": "50%",
                "backgroundColor": pc["main"], "flexShrink": "0",
                "boxShadow": f"0 0 8px {hex_to_rgba(pc['main'], 0.5)}"
            }),
            html.Span(d.get("project_name", ""), style={
                "fontSize": F["sm"], "fontWeight": "600", "color": "#e2e8f0", "flex": "1"
            }),
            html.Span(f"{overall:.0f}%", style={
                "fontSize": F["sm"], "fontWeight": "800", "color": sc,
                "backgroundColor": hex_to_rgba(sc, 0.18), "padding": "4px 12px",
                "borderRadius": "12px"
            })
        ], style={
            "display": "flex", "alignItems": "center", "gap": "10px",
            "backgroundColor": "rgba(255,255,255,0.07)", "padding": "12px 18px",
            "borderRadius": "12px", "border": "1px solid rgba(255,255,255,0.08)",
            "backdropFilter": "blur(10px)"
        }))

    return html.Div([
        html.Div([
            html.Div([
                html.H2("Portfolio Overview", style={
                    "margin": "0", "color": "#fff", "fontSize": F["xxl"], "fontWeight": "800",
                    "letterSpacing": "-0.5px"
                }),
                html.P(f"{n} Projects  •  {active} Active  •  {planning} Planning", style={
                    "margin": "6px 0 0 0", "color": "rgba(255,255,255,0.5)", "fontSize": F["sm"],
                    "fontWeight": "500"
                })
            ]),
            html.Div(project_pills, style={
                "display": "flex", "gap": "12px", "flexWrap": "wrap", "flex": "1",
                "justifyContent": "flex-end"
            })
        ], style={
            "display": "flex", "justifyContent": "space-between", "alignItems": "center",
            "gap": "30px", "flexWrap": "wrap"
        })
    ], style={
        "background": f"linear-gradient(135deg, {THEME['dark']} 0%, {THEME['dark2']} 50%, rgba(99,102,241,0.15) 100%)",
        "padding": "28px 36px", "borderRadius": "20px", "margin": "0 28px 24px 28px",
        "boxShadow": "0 8px 40px rgba(0,0,0,0.18)",
        "border": "1px solid rgba(255,255,255,0.05)"
    })


# ============================================================
# Unified Comparison Section
# ============================================================

def make_unified_comparison(all_data):
    names = [d.get("project_name", "?") for d in all_data]
    colors = [PROJECT_COLORS[i % len(PROJECT_COLORS)]["main"] for i in range(len(all_data))]

    # === Chart 1: Budget Comparison ===
    budget_fig = go.Figure()
    for i, d in enumerate(all_data):
        cur = d.get("currency", "")
        budget = d.get("client_budget", 0) if d.get("client_budget", 0) > 0 else d.get("proposed_budget", 0)
        placed = d.get("orders_placed", 0)
        prog = d.get("orders_in_progress", 0)

        budget_fig.add_trace(go.Bar(
            x=["Budget", "Ordered", "In Progress"],
            y=[budget, placed, prog],
            name=f"{names[i]} ({cur})",
            marker_color=colors[i],
            marker_line=dict(width=0),
            text=[fmt_num(budget, cur), fmt_num(placed, cur), fmt_num(prog, cur)],
            textposition="outside", textfont={"size": 11, "family": "Inter"}
        ))
    budget_fig.update_layout(
        barmode="group", height=300,
        margin=dict(t=35, b=45, l=55, r=25),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        legend=dict(orientation="h", y=1.12, x=0.5, xanchor="center", font={"size": 12}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 11}),
        xaxis=dict(tickfont={"size": 12, "color": THEME["text_light"]}),
    )

    # === Chart 2: Completion Radar ===
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
        radar_fig.add_trace(go.Scatterpolar(
            r=vals_closed, theta=cats_closed, name=names[i],
            fill="toself",
            fillcolor=hex_to_rgba(colors[i], 0.15),
            line=dict(color=colors[i], width=2.5),
            marker=dict(size=6, color=colors[i])
        ))
    radar_fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], tickfont={"size": 10}, gridcolor="rgba(0,0,0,0.06)"),
            angularaxis=dict(tickfont={"size": 12, "color": THEME["text_light"]})
        ),
        height=320, margin=dict(t=45, b=25, l=65, r=65),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center", font={"size": 12}),
        font={"family": "Inter"}
    )

    # === Chart 3: Packages & Deliveries ===
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

    pkg_del_fig.add_trace(go.Bar(
        x=[f"{n}<br>Packages" for n in x_labels], y=done_pkgs,
        name="Completed", marker_color=THEME["success"],
        text=[f"{int(v)}" for v in done_pkgs], textposition="inside",
        textfont={"size": 11, "color": "#fff", "family": "Inter"}
    ))
    pkg_del_fig.add_trace(go.Bar(
        x=[f"{n}<br>Packages" for n in x_labels], y=wip_pkgs,
        name="In Progress", marker_color=THEME["warning"],
        text=[f"{int(v)}" if v > 0 else "" for v in wip_pkgs], textposition="inside",
        textfont={"size": 11, "color": "#fff", "family": "Inter"}
    ))
    pkg_del_fig.add_trace(go.Bar(
        x=[f"{n}<br>Packages" for n in x_labels], y=pend_pkgs,
        name="Pending", marker_color="#e2e8f0",
        text=[f"{int(v)}" if v > 0 else "" for v in pend_pkgs], textposition="inside",
        textfont={"size": 11, "color": "#94a3b8", "family": "Inter"}
    ))
    pkg_del_fig.add_trace(go.Bar(
        x=[f"{n}<br>POs" for n in x_labels], y=done_pos,
        name="Delivered", marker_color=THEME["info"],
        text=[f"{int(v)}" for v in done_pos], textposition="inside",
        textfont={"size": 11, "color": "#fff", "family": "Inter"}, showlegend=True
    ))
    pkg_del_fig.add_trace(go.Bar(
        x=[f"{n}<br>POs" for n in x_labels], y=pend_pos,
        name="PO Pending", marker_color="#cbd5e1",
        text=[f"{int(v)}" if v > 0 else "" for v in pend_pos], textposition="inside",
        textfont={"size": 11, "color": "#64748b", "family": "Inter"}, showlegend=True
    ))

    pkg_del_fig.update_layout(
        barmode="stack", height=300,
        margin=dict(t=35, b=55, l=45, r=25),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter", "size": 12},
        legend=dict(orientation="h", y=1.15, x=0.5, xanchor="center", font={"size": 11}),
        yaxis=dict(gridcolor="rgba(0,0,0,0.04)", tickfont={"size": 11}),
        xaxis=dict(tickfont={"size": 11, "color": THEME["text_light"]}),
    )

    def chart_card(title, fig):
        return html.Div([
            html.P(title, style={
                "fontSize": F["section"], "color": THEME["text_light"], "fontWeight": "700",
                "letterSpacing": "1.5px", "margin": "0 0 6px 0", "textTransform": "uppercase"
            }),
            dcc.Graph(figure=fig, config={"displayModeBar": False})
        ], style=glass_style({
            "padding": "22px 24px", "flex": "1", "minWidth": "380px"
        }))

    return html.Div([
        html.P("Cross-Project Comparison", style={
            "fontSize": F["section"], "color": THEME["text_light"], "letterSpacing": "2.5px",
            "fontWeight": "700", "margin": "0 0 18px 0", "textAlign": "center",
            "textTransform": "uppercase"
        }),
        html.Div([
            chart_card("Budget Allocation", budget_fig),
            chart_card("Completion Radar", radar_fig),
        ], style={"display": "flex", "gap": "18px", "flexWrap": "wrap", "marginBottom": "18px"}),
        html.Div([
            chart_card("Packages & Deliveries Breakdown", pkg_del_fig),
        ], style={"display": "flex", "gap": "18px"})
    ], style={"padding": "0 28px 24px 28px"})


# ============================================================
# Summary Table
# ============================================================

def make_data_table(all_data):
    rows = []
    for d in all_data:
        budget = d.get("client_budget", 0) if d.get("client_budget", 0) > 0 else d.get("proposed_budget", 0)
        cur = d.get("currency", "USD")
        rows.append({
            "Project": d.get("project_name", "Unknown"),
            "Stage": d.get("current_stage", "N/A"),
            "Opening": d.get("opening_date", "TBD"),
            "Currency": cur,
            "Budget": f"{budget:,.0f}",
            "Ordered": f"{d.get('orders_placed', 0):,.0f}",
            "In Progress": f"{d.get('orders_in_progress', 0):,.0f}",
            "Packages": f"{int(d.get('packages_completed', 0))}/{int(d.get('total_packages', 0))}",
            "POs Del.": f"{int(d.get('delivered_pos', 0))}/{int(d.get('total_pos', 0))}",
            "Completion": f"{d.get('overall_completion', 0):.1f}%"
        })
    df_table = pd.DataFrame(rows)

    return html.Div([
        html.P("Project Summary", style={
            "fontSize": F["section"], "color": THEME["text_light"], "letterSpacing": "2.5px",
            "fontWeight": "700", "margin": "0 0 14px 0", "textAlign": "center",
            "textTransform": "uppercase"
        }),
        html.Div([
            dash_table.DataTable(
                data=df_table.to_dict('records'),
                columns=[{"name": c, "id": c} for c in df_table.columns],
                style_table={"overflowX": "auto", "borderRadius": "14px"},
                style_header={
                    "backgroundColor": THEME["dark"], "color": "#fff", "fontWeight": "700",
                    "fontSize": F["sm"], "textAlign": "center", "padding": "14px 12px",
                    "fontFamily": "Inter", "borderBottom": "none"
                },
                style_cell={
                    "textAlign": "center", "padding": "12px 14px", "fontSize": F["sm"],
                    "fontFamily": "Inter", "border": "none",
                    "borderBottom": f"1px solid {THEME['border']}", "minWidth": "80px",
                    "color": THEME["text"]
                },
                style_data_conditional=[
                    {"if": {"row_index": "odd"}, "backgroundColor": "rgba(248,250,252,0.7)"},
                    {"if": {"row_index": "even"}, "backgroundColor": "rgba(255,255,255,0.7)"},
                ]
            )
        ], style=glass_style({"overflow": "hidden", "padding": "0"}))
    ], style={"padding": "0 28px 28px 28px"})


# ============================================================
# Project Card
# ============================================================

def make_project_card(data, pc_idx):
    pc = PROJECT_COLORS[pc_idx % len(PROJECT_COLORS)]
    name = data.get("project_name", "Unknown")
    stage = data.get("current_stage", "N/A")
    opening = data.get("opening_date", "TBD")
    currency = data.get("currency", "USD")
    proposed = data.get("proposed_budget", 0)
    client_b = data.get("client_budget", 0)
    budget = client_b if client_b > 0 else proposed
    placed = data.get("orders_placed", 0)
    progress = data.get("orders_in_progress", 0)
    savings = data.get("savings_overrun", 0)
    pkgs = data.get("total_packages", 0)
    pkg_done = data.get("packages_completed", 0)
    pos = data.get("total_pos", 0)
    del_done = data.get("delivered_pos", 0)
    overall = data.get("overall_completion", 0)
    pkg_pct = data.get("pkg_completion_pct", 0)
    del_pct = data.get("delivery_completion_pct", 0)
    concerns = data.get("concerns", [])

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

    # === Gauge ===
    gauge = go.Figure(go.Indicator(
        mode="gauge+number", value=round(overall, 1),
        number={"suffix": "%", "font": {"size": 38, "color": pc["main"], "family": "Inter"}},
        gauge={
            "axis": {"range": [0, 100], "dtick": 25, "tickcolor": "#e2e8f0",
                     "tickfont": {"size": 10, "color": "#94a3b8"}, "tickwidth": 1},
            "bar": {"color": oc, "thickness": 0.3},
            "bgcolor": "#f8fafc", "borderwidth": 0,
            "steps": [{"range": [0, 100], "color": "#f1f5f9"}],
            "threshold": {"line": {"color": pc["main"], "width": 2}, "thickness": 0.8, "value": overall}
        }
    ))
    gauge.update_layout(height=175, margin=dict(t=18, b=0, l=28, r=28),
                        paper_bgcolor="rgba(0,0,0,0)", font={"family": "Inter"})

    # === Progress bars ===
    def prog_row(label, val, total, color):
        pct = (val / total * 100) if total > 0 else 0
        return html.Div([
            html.Div([
                html.Span(label, style={
                    "fontSize": F["sm"], "color": THEME["text_light"], "fontWeight": "600"
                }),
                html.Span(f"{val:,.0f} / {total:,.0f}", style={
                    "fontSize": F["xs"], "color": THEME["text_muted"]
                }),
            ], style={"display": "flex", "justifyContent": "space-between", "marginBottom": "5px"}),
            html.Div([
                html.Div(style={
                    "width": f"{min(pct, 100)}%", "height": "100%",
                    "backgroundColor": color, "borderRadius": "5px",
                    "transition": "width 0.6s ease",
                    "boxShadow": f"0 0 8px {hex_to_rgba(color, 0.3)}" if pct > 10 else "none"
                })
            ], style={
                "width": "100%", "height": "9px", "backgroundColor": "#f1f5f9",
                "borderRadius": "5px", "overflow": "hidden"
            })
        ], style={"marginBottom": "12px"})

    progress_children = [
        prog_row("Budget Utilization", placed + progress, budget, pc["main"]),
        prog_row("Packages Completed", pkg_done, pkgs, THEME["info"]),
    ]
    if pos > 0:
        progress_children.append(prog_row("POs Delivered", del_done, pos, THEME["success"]))

    # === KPI Row ===
    def kpi_chip(label, value, icon, color):
        return html.Div([
            html.Span(icon, style={"fontSize": "20px"}),
            html.Div([
                html.P(label, style={
                    "fontSize": F["xs"], "color": THEME["text_muted"], "margin": "0",
                    "textTransform": "uppercase", "letterSpacing": "0.5px", "fontWeight": "700"
                }),
                html.P(value, style={
                    "fontSize": F["value"], "color": THEME["text"], "margin": "2px 0 0 0",
                    "fontWeight": "700"
                })
            ])
        ], style={
            "display": "flex", "alignItems": "center", "gap": "10px",
            "backgroundColor": hex_to_rgba(color, 0.06), "padding": "10px 14px",
            "borderRadius": "12px", "flex": "1", "minWidth": "110px",
            "border": f"1px solid {hex_to_rgba(color, 0.12)}"
        })

    sav_color = THEME["success"] if savings >= 0 else THEME["danger"]
    sav_label = "Savings" if savings >= 0 else "Overrun"

    kpi_row = html.Div([
        kpi_chip("Budget", fmt_num(budget, currency), "💰", pc["main"]),
        kpi_chip(sav_label, fmt_num(abs(savings), currency), "📈" if savings >= 0 else "📉", sav_color),
        kpi_chip("Packages", f"{int(pkg_done)}/{int(pkgs)}", "📦", THEME["info"]),
        kpi_chip("Deliveries", f"{int(del_done)}/{int(pos)}", "🚚", THEME["success"]),
    ], style={"display": "flex", "gap": "8px", "flexWrap": "wrap", "marginBottom": "16px"})

    # === Concerns ===
    if concerns:
        cblock = html.Div([
            html.Div([
                html.Span(c, style={"lineHeight": "1.6"})
            ], style={
                "padding": "10px 14px", "backgroundColor": "#fffbeb", "borderRadius": "10px",
                "fontSize": F["sm"], "borderLeft": f"3px solid {THEME['warning']}",
                "marginBottom": "6px", "color": "#92400e"
            }) for c in concerns
        ])
    else:
        cblock = html.P("No concerns reported", style={
            "color": THEME["success"], "fontSize": F["sm"], "margin": "0",
            "padding": "10px 14px", "backgroundColor": hex_to_rgba(THEME["success"], 0.06),
            "borderRadius": "10px", "borderLeft": f"3px solid {THEME['success']}",
            "fontWeight": "500"
        })

    return html.Div([
        # Header
        html.Div([
            html.Div([
                html.H3(name, style={
                    "margin": "0", "color": "#fff", "fontSize": F["xl"],
                    "fontWeight": "800", "letterSpacing": "-0.3px"
                }),
                html.Div([
                    html.Span(stage, style={
                        "backgroundColor": stage_bg, "padding": "4px 14px", "borderRadius": "20px",
                        "fontSize": F["xs"], "color": "#fff", "fontWeight": "600",
                        "boxShadow": f"0 2px 8px {hex_to_rgba(stage_bg, 0.4)}"
                    }),
                ], style={"marginTop": "8px"})
            ], style={"flex": "1"}),
            html.Div([
                html.P("OPENING", style={
                    "margin": "0", "fontSize": "9px", "color": "rgba(255,255,255,0.5)",
                    "letterSpacing": "1.5px", "fontWeight": "700"
                }),
                html.P(str(opening), style={
                    "margin": "3px 0 0 0", "color": "#fff", "fontSize": F["md"], "fontWeight": "600"
                })
            ], style={"textAlign": "right"})
        ], style={
            "display": "flex", "justifyContent": "space-between", "alignItems": "center",
            "background": pc["gradient"], "padding": "20px 24px",
            "borderRadius": "18px 18px 0 0"
        }),

        # Body
        html.Div([
            kpi_row,

            # Gauge & Progress side by side
            html.Div([
                html.Div([
                    html.P("OVERALL COMPLETION", style={
                        "textAlign": "center", "fontSize": F["xs"], "color": THEME["text_muted"],
                        "letterSpacing": "1.5px", "fontWeight": "700", "margin": "0"
                    }),
                    dcc.Graph(figure=gauge, config={"displayModeBar": False}),
                ], style={"flex": "1", "minWidth": "170px"}),

                html.Div([
                    html.P("PROGRESS DETAILS", style={
                        "fontSize": F["xs"], "color": THEME["text_muted"],
                        "letterSpacing": "1.5px", "fontWeight": "700", "margin": "0 0 14px 0"
                    }),
                    *progress_children
                ], style={"flex": "1.2", "minWidth": "200px"})
            ], style={
                "display": "flex", "gap": "16px", "marginBottom": "16px",
                "backgroundColor": "rgba(255,255,255,0.9)", "borderRadius": "14px",
                "padding": "16px", "boxShadow": "0 2px 8px rgba(0,0,0,0.03)", "flexWrap": "wrap"
            }),

            # Concerns
            html.Div([
                html.P("CONCERNS & NOTES", style={
                    "fontSize": F["xs"], "color": THEME["text_muted"],
                    "letterSpacing": "1.5px", "fontWeight": "700", "margin": "0 0 10px 0"
                }),
                cblock
            ], style={
                "backgroundColor": "rgba(255,255,255,0.9)", "borderRadius": "14px",
                "padding": "16px", "boxShadow": "0 2px 8px rgba(0,0,0,0.03)"
            })
        ], style={
            "padding": "20px 24px 24px 24px",
            "backgroundColor": hex_to_rgba(pc["main"], 0.04)
        })
    ], style=glass_style({
        "overflow": "hidden", "flex": "1", "minWidth": "400px", "maxWidth": "560px",
        "backgroundColor": "rgba(255, 255, 255, 0.75)",
    }))


# ============================================================
# DASH APP
# ============================================================

app = dash.Dash(
    __name__,
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1.0"}]
)

app.layout = html.Div([
    # Upload page
    html.Div(id="upload-section", children=[
        html.Div(style={"height": "15vh"}),
        html.Div([
            html.Div([
                html.Div(style={
                    "width": "72px", "height": "72px", "borderRadius": "20px",
                    "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                    "display": "flex", "alignItems": "center", "justifyContent": "center",
                    "margin": "0 auto 20px auto",
                    "boxShadow": "0 8px 24px rgba(99,102,241,0.35)"
                }, children=[
                    html.Span("📊", style={"fontSize": "32px"})
                ]),
                html.H1("Dashboard", style={
                    "margin": "0 0 8px 0", "color": THEME["text"],
                    "fontSize": "32px", "fontWeight": "800", "letterSpacing": "-0.5px"
                }),
                html.P("Upload data file to generate insights", style={
                    "color": THEME["text_muted"], "fontSize": F["md"], "margin": "0 0 36px 0",
                    "fontWeight": "400"
                }),
            ]),
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    html.Div([
                        html.Div(style={
                            "width": "56px", "height": "56px", "borderRadius": "14px",
                            "backgroundColor": "#eef2ff", "display": "flex",
                            "alignItems": "center", "justifyContent": "center",
                            "margin": "0 auto 14px auto"
                        }, children=[
                            html.Span("📁", style={"fontSize": "26px"})
                        ]),
                        html.P("Drag & Drop or Click to Upload", style={
                            "fontWeight": "700", "color": "#6366f1",
                            "fontSize": F["lg"], "margin": "0 0 6px 0"
                        }),
                        html.P(".xlsx  •  .xlsm  •  .xls  •  .csv", style={
                            "color": THEME["text_muted"], "fontSize": F["sm"], "margin": "0"
                        })
                    ])
                ]),
                style={
                    'width': '100%', 'padding': '45px 20px', 'borderWidth': '2px',
                    'borderStyle': 'dashed', 'borderColor': '#c7d2fe', 'borderRadius': '16px',
                    'textAlign': 'center', 'backgroundColor': 'rgba(250,251,255,0.8)',
                    'cursor': 'pointer',
                },
                multiple=False
            ),
            html.Div(id='upload-error', style={"marginTop": "18px"})
        ], style=glass_style({
            "padding": "52px 60px", "maxWidth": "520px",
            "margin": "0 auto", "textAlign": "center"
        }))
    ]),

    # Dashboard page
    html.Div(id="dashboard-section", style={"display": "none"}, children=[
        # Top bar
        html.Div([
            html.Div([
                html.Div(style={
                    "width": "42px", "height": "42px", "borderRadius": "12px",
                    "background": "linear-gradient(135deg, #6366f1, #8b5cf6)",
                    "display": "flex", "alignItems": "center", "justifyContent": "center",
                    "boxShadow": "0 4px 12px rgba(99,102,241,0.3)"
                }, children=[html.Span("📊", style={"fontSize": "18px"})]),
                html.Div([
                    html.H1("Dashboard", style={
                        "margin": "0", "color": "#fff", "fontSize": F["xl"],
                        "fontWeight": "800", "letterSpacing": "-0.3px"
                    }),
                    html.P("Project insights and analysis", style={
                        "color": "rgba(255,255,255,0.5)", "fontSize": F["sm"], "margin": "0",
                        "fontWeight": "500"
                    })
                ])
            ], style={"display": "flex", "alignItems": "center", "gap": "14px", "flex": "1"}),
            html.A("← Upload New File", id="btn-back", href="/", style={
                "backgroundColor": "rgba(255,255,255,0.08)", "color": "#e2e8f0",
                "border": "1px solid rgba(255,255,255,0.12)", "padding": "10px 24px",
                "borderRadius": "10px", "cursor": "pointer", "fontSize": F["sm"],
                "fontWeight": "600", "textDecoration": "none", "display": "inline-block",
            })
        ], style={
            "display": "flex", "justifyContent": "space-between", "alignItems": "center",
            "background": f"linear-gradient(135deg, {THEME['dark']}, {THEME['dark2']})",
            "padding": "16px 34px", "marginBottom": "24px",
            "boxShadow": "0 4px 20px rgba(0,0,0,0.15)"
        }),
        html.Div(id="dashboard-body")
    ])
], style={
    "fontFamily": "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
    "backgroundImage": "url('/assets/background.jpg')",
    "backgroundSize": "cover",
    "backgroundPosition": "center",
    "backgroundRepeat": "no-repeat",
    "backgroundAttachment": "fixed",
    "minHeight": "100vh",
})


# ============================================================
# Callback
# ============================================================

@app.callback(
    Output("upload-section", "style"),
    Output("dashboard-section", "style"),
    Output("dashboard-body", "children"),
    Output("upload-error", "children"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True
)
def handle_upload(contents, filename):
    if contents is None:
        return dash.no_update, dash.no_update, dash.no_update, ""

    ext = ""
    if filename and '.' in filename:
        ext = filename.lower().rsplit('.', 1)[-1]

    if ext not in ['xlsx', 'xlsm', 'xls', 'csv']:
        return (dash.no_update, dash.no_update, dash.no_update,
                html.P(f"Unsupported format: .{ext}",
                       style={"color": THEME["danger"], "fontWeight": "600", "fontSize": F["md"]}))

    try:
        all_data = parse_file(contents, filename)
    except Exception as e:
        return (dash.no_update, dash.no_update, dash.no_update,
                html.P(f"Error processing file: {str(e)}",
                       style={"color": THEME["danger"], "fontSize": F["sm"]}))

    if not all_data:
        return (dash.no_update, dash.no_update, dash.no_update,
                html.P("No valid project data found in this file.",
                       style={"color": THEME["danger"], "fontSize": F["sm"]}))

    cards = []
    for idx, d in enumerate(all_data):
        cards.append(make_project_card(d, idx))

    header = make_portfolio_header(all_data)
    comparison = make_unified_comparison(all_data) if len(all_data) >= 2 else html.Div()
    table = make_data_table(all_data)

    body = html.Div([
        header,
        comparison,
        table,
        html.Div(style={"height": "12px"}),
        html.P("Project Details", style={
            "fontSize": F["section"], "color": THEME["text_light"], "letterSpacing": "2.5px",
            "fontWeight": "700", "margin": "0 0 16px 0", "textAlign": "center",
            "textTransform": "uppercase"
        }),
        html.Div(cards, style={
            "display": "flex", "justifyContent": "center", "alignItems": "flex-start",
            "gap": "24px", "flexWrap": "wrap", "padding": "0 28px 48px 28px"
        })
    ])

    return {"display": "none"}, {"display": "block"}, body, ""


# ============================================================
# Run
# ============================================================

if __name__ == '__main__':
    port = 8050
    url = f"http://127.0.0.1:{port}"
    print("\n" + "=" * 50)
    print(f"  Dashboard running at: {url}")
    print("=" * 50 + "\n")
    Timer(2.0, open_browser, args=[url]).start()
    app.run(debug=False, port=port, host="127.0.0.1")
