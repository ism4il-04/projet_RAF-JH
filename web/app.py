import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from shiny import App, ui, render, reactive

app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.h2("📁 Data Uploader"),
        ui.input_file("file", "Choose CSV or Excel file", accept=[".csv", ".xlsx", ".xls"]),
        ui.hr(),
        ui.p("Upload a file to view its summary, charts, and table."),
    ),
    ui.div(
        ui.h3("File Information"),
        ui.output_ui("file_info"),
        ui.hr(),
        ui.layout_columns(
            ui.div(
                ui.h4("Bar Chart: Charge JH par consultant"),
                ui.output_ui("bar_chart_ui"),
                class_="card p-3 mb-3"
            ),
            ui.div(
                ui.h4("Pie Chart: Distribution de l'ecarts"),
                ui.output_ui("pie_chart_ui"),
                class_="card p-3 mb-3"
            ),
            col_widths=(6, 6)
        ),
        ui.hr(),
        ui.h4("Data Table"),
        ui.div(
            ui.output_ui("table_ui"),
            style="max-height:400px;overflow:auto; border:1px solid #eee; border-radius:8px;"
        ),
        class_="container-fluid"
    )
)

def read_uploaded_file(fileinfo):
    if fileinfo is None or len(fileinfo) == 0:
        return None
    file = fileinfo[0]
    ext = Path(file["name"]).suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(file["datapath"])
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file["datapath"])
    else:
        return None
    return df

def plot_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return f'<img src="data:image/png;base64,{img_base64}" style="max-width:100%;height:auto;"/>'

def server(input, output, session):
    @reactive.Calc
    def df():
        return read_uploaded_file(input.file())

    def get_column(data, colname):
        norm_cols = {c.strip().lower(): c for c in data.columns}
        return norm_cols.get(colname.strip().lower(), None)

    @output()
    @render.ui
    def file_info():
        fileinfo = input.file()
        data = df()
        if not fileinfo or data is None:
            return ui.p("No file uploaded.")
        file = fileinfo[0]
        return ui.div(
            ui.p(f"<b>File name:</b> {file['name']}", class_="mb-1"),
            ui.p(f"<b>Rows:</b> {data.shape[0]}, <b>Columns:</b> {data.shape[1]}", class_="mb-1"),
            ui.p(f"<b>Columns:</b> {', '.join(data.columns)}", class_="mb-1")
        )

    @output()
    @render.ui
    def bar_chart_ui():
        data = df()
        if data is None:
            return ui.p("Bar chart: No data uploaded.")
        col_proj = get_column(data, "Resource/ PROJET")
        col_jh = get_column(data, "Somme de Charge JH")
        if col_proj is None or col_jh is None:
            return ui.p(f"Bar chart: Required columns not found. Available columns: {list(data.columns)}")
        chart_data = data.dropna(subset=[col_jh])
        if chart_data.empty:
            return ui.p("Bar chart: No data to display after filtering NaN values.")
        fig, ax = plt.subplots(figsize=(6, 3))
        ax.bar(chart_data[col_proj].astype(str), chart_data[col_jh])
        ax.set_xlabel("Consultants")
        ax.set_ylabel(col_jh)
        ax.set_title("Charge JH par consultant")
        plt.xticks(rotation=45, ha="right")
        return ui.HTML(plot_to_base64(fig))

    @output()
    @render.ui
    def pie_chart_ui():
        data = df()
        if data is None:
            return ui.p("Pie chart: No data uploaded.")
        col_ecart = get_column(data, "ecart")
        if col_ecart is None:
            return ui.p(f"Pie chart: 'ecart' column not found. Available columns: {list(data.columns)}")
        ecart = data[col_ecart].dropna()
        if ecart.empty:
            return ui.p("Pie chart: No data to display after filtering NaN values.")
        categories = [
            (ecart > 0).sum(),
            (ecart < 0).sum(),
            (ecart == 0).sum(),
        ]
        labels = ["Positive", "Negative", "Zero"]
        fig, ax = plt.subplots()
        ax.pie(categories, labels=labels, autopct="%1.1f%%", startangle=90)
        ax.set_title("Distribution de l'ecarts")
        return ui.HTML(plot_to_base64(fig))

    @output()
    @render.ui
    def table_ui():
        data = df()
        if data is None:
            return ui.p("No file uploaded or unsupported file type.")
        return ui.HTML(data.to_html(index=False, classes="table table-striped"))

app = App(app_ui, server)
