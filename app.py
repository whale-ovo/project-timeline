
from time import strftime
import dash
from dash import dcc, html, Input, Output
import plotly.graph_objects as go
import pandas as pd
import locale
import os


app = dash.Dash(__name__)
server = app.server

# 设置中文语言环境（兼容多系统）
try:
    locale.setlocale(locale.LC_TIME, 'zh_CN.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'zh-CN')  # Windows备用
    except:
        print("警告：中文语言环境设置失败，日期可能显示为英文")


# ----------------------
# 1. 示例数据生成（适配延期列为时间类型）
# ----------------------
def create_sample_data():
    data = {
        "项目": ["A", "B", "C", "D", "E", "F"],
        "TR1": ["2024-12-01", "2024-12-15", "2025-01-10", "2024-11-20", "2025-02-05", "2025-03-01"],
        "TR1延期": [None, "2025-03-01", None, "2025-09-01", None, None],  # 时间类型：None=无延期，日期=延期时间
        "TR2": ["2025-01-20", "2025-02-28", "2025-03-15", "2025-01-05", "2025-04-10", "2025-05-10"],
        "TR2延期": ["2025-01-23", None, "2025-03-25", None, None, "2025-05-15"],
        "TR3": ["2025-03-10", "2025-04-30", "2025-05-20", "2025-03-01", "2025-06-15", "2025-07-10"],
        "TR3延期": [None, "2025-05-05", None, None, "2025-06-22", None],
        "TR3A": ["2025-04-01", "2025-05-15", "2025-06-10", "2025-04-05", "2025-07-20", "2025-08-05"],
        "TR3A延期": ["2025-04-03", None, "2025-06-13", None, None, None],
        "TR4": ["2025-05-15", "2025-06-30", "2025-08-01", "2025-05-20", "2025-09-10", "2025-09-30"],
        "TR4延期": [None, None, "2025-08-06", "2025-05-22", None, None],
        "TR4A": ["2025-06-01", "2025-07-15", "2025-08-20", "2025-06-05", "2025-09-25", "2025-10-15"],
        "TR4A延期": [None, "2025-07-18", None, None, "2025-09-30", None],
        "TR5": ["2025-07-10", "2025-08-30", "2025-10-01", "2025-07-20", "2025-11-05", "2025-11-20"],
        "TR5延期": ["2025-07-15", None, None, "2025-07-23", None, None],
        "TR6": ["2025-08-20", "2025-10-15", "2025-12-01", "2025-09-10", "2025-12-15", "2025-12-30"],
        "TR6延期": [None, None, "2025-12-11", None, None, None]
    }
    # 转换所有TR列（含延期列）为时间类型，None/0转为NaT
    for col in data:
        if col.startswith("TR"):  # 包括"TR1"和"TR1延期"
            data[col] = pd.to_datetime(data[col], errors='coerce').tolist()
    return data


# ----------------------
# 2. Excel数据读取（适配延期列为时间类型）
# ----------------------
def load_excel_data(file_path="D:\\1032.xlsx"):
    actual_path = os.path.abspath(file_path)
    print(f"正在查找Excel文件：{actual_path}")

    # 1. 检查文件是否存在
    if not os.path.exists(actual_path):
        print(f"警告：未找到文件 {actual_path}，使用示例数据")
        return create_sample_data()

    # 2. 读取Excel文件（保留原始列名）
    file_ext = os.path.splitext(actual_path)[1].lower()
    try:
        if file_ext == ".xls":
            df = pd.read_excel(actual_path, engine="xlrd")
        else:
            df = pd.read_excel(actual_path, engine="openpyxl")
        print(f"Excel读取成功")
    except Exception as e:
        print(f"读取Excel失败：{str(e)}，使用示例数据")
        return create_sample_data()

    # 3. 查找"项目"列
    clean_col_map = {col.strip().lower(): col for col in df.columns}
    project_candidates = ["项目", "项目名", "项目名称", "产品", "产品名"]
    project_original_col = None
    for candidate in project_candidates:
        if candidate.lower() in clean_col_map.keys():
            project_original_col = clean_col_map[candidate.lower()]
            break
    if project_original_col is None:
        print(f"错误：未找到项目相关列！Excel中的列名：{df.columns.tolist()}")
        return create_sample_data()

    # 4. 检查TR相关列（含延期列）
    required_tr_cols = [
        "TR1", "TR1延期", "TR2", "TR2延期",
        "TR3", "TR3延期", "TR3A", "TR3A延期",
        "TR4", "TR4延期", "TR4A", "TR4A延期",
        "TR5", "TR5延期", "TR6", "TR6延期"
    ]
    missing_tr_cols = []
    tr_col_map = {}
    for tr_col in required_tr_cols:
        clean_tr_col = tr_col.strip().lower()
        if clean_tr_col not in clean_col_map.keys():
            missing_tr_cols.append(tr_col)
        else:
            tr_col_map[tr_col] = clean_col_map[clean_tr_col]
    if missing_tr_cols:
        print(f"Excel缺少TR列：{missing_tr_cols}，使用示例数据")
        return create_sample_data()

    # 5. 提取并清理数据
    keep_original_cols = [project_original_col] + list(tr_col_map.values())
    df_filtered = df[keep_original_cols].copy()
    # 重命名列（统一为代码预期列名）
    rename_dict = {project_original_col: "项目"}
    for code_col, original_col in tr_col_map.items():
        rename_dict[original_col] = code_col
    df_renamed = df_filtered.rename(columns=rename_dict)
    # 删除项目为空的行
    df_renamed = df_renamed.dropna(subset=["项目"]).reset_index(drop=True)
    if len(df_renamed) == 0:
        print("Excel无有效项目数据，使用示例数据")
        return create_sample_data()

    # 6. 转换数据格式：所有TR列（含延期列）转为时间类型
    columns_as_lists = {}
    # 项目列
    columns_as_lists["项目"] = df_renamed["项目"].astype(str).tolist()
    # TR列（含延期列）：转为datetime，无效值为NaT
    for tr_col in required_tr_cols:
        columns_as_lists[tr_col] = pd.to_datetime(df_renamed[tr_col], errors='coerce').tolist()

    return columns_as_lists


# ----------------------
# 3. 数据初始化（适配延期时间逻辑）
# ----------------------
excel_data = load_excel_data()
projects = []
y_pos = 0
y_step = 30  # 行间距，避免矩形重叠
colors = [
    "#2E8B57", "#FF8C00", "#1E90FF", "#FF6347", "#9370DB", "#00CED1",
    "#8B4513", "#FF69B4", "#00FF7F", "#FFD700", "#191970", "#FF4500",
    "#20B2AA", "#DDA0DD", "#F08080", "#98FB98", "#FFA07A"
]

# 生成项目列表：计算延期天数（延期时间 - TR节点时间）
for i, project_name in enumerate(excel_data["项目"]):
    # 定义TR节点顺序（用于判断start_date和end_date）
    tr_order = ["TR1", "TR1延期", "TR2", "TR2延期", "TR3", "TR3延期",
                "TR3A", "TR3A延期", "TR4", "TR4延期", "TR4A", "TR4A延期",
                "TR5", "TR5延期", "TR6", "TR6延期"]  # 从早到晚的TR节点
    reverse_tr_order = tr_order[::-1]  # 反转顺序（从晚到早：TR6→TR5→...→TR1）

    # ----------------------
    # 计算start_date：从TR1开始找第一个非空节点，减15天
    # ----------------------
    start_base_date = None
    for tr in tr_order:
        tr_date = excel_data[tr][i]
        if not pd.isna(tr_date):  # 找到第一个非空的TR节点
            start_base_date = tr_date
            break
    # 若所有TR节点都为空，用当前时间兜底
    if start_base_date is None:
        start_base_date = pd.Timestamp.now()
    start_date = start_base_date - pd.Timedelta(days=15)

    # ----------------------
    # 计算end_date：从TR6开始找第一个非空节点，加15天
    # ----------------------
    end_base_date = None
    for tr in reverse_tr_order:
        tr_date = excel_data[tr][i]
        if not pd.isna(tr_date):  # 找到第一个非空的TR节点
            end_base_date = tr_date
            break
    # 若所有TR节点都为空，用当前时间+90天兜底
    if end_base_date is None:
        end_base_date = pd.Timestamp.now() + pd.Timedelta(days=90)
    end_date = end_base_date + pd.Timedelta(days=15)

    # 过滤无效TR节点 + 计算延期天数
    valid_trs_delay = {}
    tr_keys = ["TR1", "TR2", "TR3", "TR3A", "TR4", "TR4A", "TR5", "TR6",
               "TR1延期", "TR2延期", "TR3延期", "TR3A延期", "TR4延期", "TR4A延期", "TR5延期", "TR6延期"]

    for tr_key in tr_keys:
        tr_date = excel_data[tr_key][i]
        delay_date = excel_data[f"{tr_key}"][i]  # 延期时间（时间类型）

        # 仅保留TR节点时间有效的数据
        if not pd.isna(tr_date):
            # 计算延期天数：延期时间存在且晚于TR节点时间 → 计算差值；否则0
            if not pd.isna(delay_date) and delay_date > tr_date:
                delay_days = (delay_date - tr_date).days  # 延期天数（整数）
            else:
                delay_days = 0  # 无延期或延期时间无效
            valid_trs_delay[tr_key] = {"date": tr_date, "delay": delay_days}  # 存储"TR时间+延期天数"

    projects.append({
        "name": str(project_name).strip(),
        "start_date": start_date,
        "end_date": end_date,
        "y_pos": y_pos,
        "trs_delay": valid_trs_delay
    })
    y_pos += y_step

# 计算全局时间范围
if projects:
    min_date = min(p["start_date"] for p in projects)
    max_date = max(p["end_date"] for p in projects)
else:
    min_date = pd.Timestamp.now() - pd.Timedelta(days=30)
    max_date = pd.Timestamp.now() + pd.Timedelta(days=30)
tick_dates = pd.date_range(start=min_date, end=max_date, freq="MS")
tick_texts = [f"{d.strftime('%Y')}<br>{d.strftime('%m.%d')}" for d in tick_dates]


# ----------------------
# 4. 滚动区域图表生成（功能不变，延期逻辑已适配）
# ----------------------
def create_scrollable_fig(selected_projects, xaxis_range=None):
    if not selected_projects:
        selected_projects = [p["name"] for p in projects]
    filtered_projects = [p for p in projects if p["name"] in selected_projects]

    fig = go.Figure()
    for p in filtered_projects:
        # 项目时间线
        fig.add_trace(go.Scatter(
            x=[p["start_date"], p["end_date"]],
            y=[p["y_pos"], p["y_pos"]],
            mode="lines",
            line=dict(color=colors[projects.index(p)], width=8),
            name=p["name"],
            showlegend=True,
            hoverinfo="none"
        ))

        # TR节点：延期节点贴合时间线，红色矩形
        for tr_name, tr_info in p["trs_delay"].items():
            fig.add_trace(go.Scatter())
            is_delayed = "延期" in tr_name  # 基于计算出的延期天数判断
            node_color = "#FF0000" if is_delayed else "#56C440"#colors[projects.index(p)]
            text_color = "white" if is_delayed else "black"
            formatted_tr_date = tr_info["date"].strftime("%Y年%m月%d日")
            # 延期节点显示"延期X天"，正常节点显示"无延期"
            delay_text = f"延期{tr_info['delay']}天" if is_delayed else ""

            # 添加矩形+文字标注
            fig.add_trace(go.Scatter(
                x=[tr_info["date"]],
                y=[p["y_pos"]],
                mode="markers+text",
                marker=dict(
                    color=node_color,
                    symbol=("diamond-wide"if is_delayed else "square"),
                    size=30,
                    line=dict(width=1, color="black")
                ),
                text=[tr_name.replace("延期", "") if is_delayed else tr_name],
                textposition="middle center",
                textfont=dict(color=text_color, size=11, weight="bold"),
                # 悬停提示：显示TR时间、延期状态和延期天数
                customdata=[[p["name"], tr_name, formatted_tr_date, delay_text]],
                hovertemplate="""
                    %{customdata[0]} %{customdata[1]}<br>
                    TR时间: %{customdata[2]}}
                """,
                hoverinfo="none",
                name=f"{p['name']}-{tr_name}",
                showlegend=False
            ))

    # X轴配置 - 支持滚动
    xaxis_config = dict(
        range=xaxis_range if xaxis_range else [min_date, max_date],
        dtick="M1",
        tickvals=pd.date_range(start=min_date, end=max_date, freq="MS"),
        showticklabels=False,
        showgrid=True,
        gridcolor="#CCCCCC",
        gridwidth=1,
        linecolor="#FF8C00",
        fixedrange=False  # 允许X轴缩放和滚动
    )

    # Y轴配置 - 固定
    yaxis_config = dict(
        tickvals=[p["y_pos"] for p in filtered_projects],
        ticktext=[p["name"] for p in filtered_projects],
        showgrid=True,
        gridcolor="#CCCCCC",
        gridwidth=1,
        linecolor="#CCCCCC",
        tickfont=dict(size=20, weight="bold", color="RED"),
        fixedrange=True  # 锁定Y轴
    )

    fig.update_layout(
        xaxis=xaxis_config,
        yaxis=yaxis_config,
        plot_bgcolor="white",
        legend=dict(orientation="h", y=1.05),
        margin=dict(l=0, r=0, t=0, b=0),
        dragmode="pan"  # 设置拖拽模式为平移
    )
    return fig


# ----------------------
# 5. 固定X轴图表生成（支持滚动）
# ----------------------
def create_fixed_xaxis_fig(xaxis_range=None):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=[min_date, max_date], y=[0, 0], mode="lines", line=dict(width=0)))

    xaxis_config = dict(
        range=xaxis_range if xaxis_range else [min_date, max_date],
        dtick="M1",
        tickvals=tick_dates,
        ticktext=tick_texts,
        tickangle=0,
        linecolor="#FF8C00",
        tickfont=dict(size=10),
        side="top",
        fixedrange=False  # 允许X轴缩放和滚动
    )

    fig.update_xaxes(xaxis_config)
    fig.update_yaxes(showticklabels=False, showgrid=False, linecolor="rgba(0,0,0,0)", fixedrange=True)
    fig.update_layout(
        plot_bgcolor="white",
        margin=dict(l=100, r=0, t=0, b=0),
        height=60,
        dragmode="pan"  # 设置拖拽模式为平移
    )
    return fig


# ----------------------
# 6. Dash应用构建与回调
# ----------------------
app = dash.Dash(__name__)

app.layout = html.Div([
    html.Div([
        dcc.Dropdown(
            id="project-filter",
            options=[{"label": p["name"], "value": p["name"]} for p in projects],
            value=[p["name"] for p in projects],
            multi=True,
            placeholder="选择要显示的项目",
            style={"width": "80%", "margin": "0 auto"}
        )
    ], style={"padding": "15px"}),

    # 横向滚动容器
    html.Div([
        # 固定X轴图表
        html.Div([
            dcc.Graph(
                id="fixed-xaxis-fig",
                config={
                    "displayModeBar": False,
                    "staticPlot": False,
                    "locale": "zh-cn",
                    "scrollZoom": True
                }
            )
        ], style={"position": "sticky", "top": 0, "zIndex": 100, "background": "white"}),

        # 主图表区域
        html.Div([
            dcc.Graph(
                id="scrollable-fig",
                config={
                    "displayModeBar": False,
                    "locale": "zh-cn",
                    "scrollZoom": True
                }
            )
        ], style={"height": "800px"})
    ], style={"overflowX": "auto", "width": "100%"})
])


# 同步两个图表的X轴范围
@app.callback(
    [Output("scrollable-fig", "figure"),
     Output("fixed-xaxis-fig", "figure")],
    [Input("project-filter", "value"),
     Input("scrollable-fig", "relayoutData"),
     Input("fixed-xaxis-fig", "relayoutData")]
)
def update_figures(selected_projects, scrollable_relayout, fixed_relayout):
    ctx = dash.callback_context
    trigger_id = ctx.triggered[0]["prop_id"].split(".")[0] if ctx.triggered else None

    xaxis_range = None

    # 根据触发源获取X轴范围
    if trigger_id == "scrollable-fig" and scrollable_relayout:
        if "xaxis.range[0]" in scrollable_relayout and "xaxis.range[1]" in scrollable_relayout:
            xaxis_range = [scrollable_relayout["xaxis.range[0]"], scrollable_relayout["xaxis.range[1]"]]
        elif "xaxis.range" in scrollable_relayout:
            xaxis_range = scrollable_relayout["xaxis.range"]

    elif trigger_id == "fixed-xaxis-fig" and fixed_relayout:
        if "xaxis.range[0]" in fixed_relayout and "xaxis.range[1]" in fixed_relayout:
            xaxis_range = [fixed_relayout["xaxis.range[0]"], fixed_relayout["xaxis.range[1]"]]
        elif "xaxis.range" in fixed_relayout:
            xaxis_range = fixed_relayout["xaxis.range"]

    # 创建同步的图表
    scrollable_fig = create_scrollable_fig(selected_projects, xaxis_range)
    fixed_fig = create_fixed_xaxis_fig(xaxis_range)

    return scrollable_fig, fixed_fig


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8050))
    app.run(debug=False, host='0.0.0.0', port=port)