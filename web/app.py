import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from shiny import App, ui, render, reactive

app_ui = ui.page_fluid(
    ui.h2("Upload a file and view its table and charts"),
    ui.input_file("file", "Choose CSV or Excel file", accept=[".csv", ".xlsx", ".xls"]),
    ui.output_ui("bar_chart_ui"),
    ui.output_ui("pie_chart_ui"),
    ui.output_ui("table_ui"),
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
        # Normalize columns: strip and lower
        norm_cols = {c.strip().lower(): c for c in data.columns}
        return norm_cols.get(colname.strip().lower(), None)

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
        fig, ax = plt.subplots(figsize=(8, 4))
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
