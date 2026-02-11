from pathlib import Path
import base64
import io
from datetime import datetime

import dash
from dash import Input, Output, State, dash_table, dcc, html
import pandas as pd
import plotly.express as px
import plotly.io as pio


FILE_PATH = Path("CAMOTE.xlsx")


def _empty_fig(title: str):
    fig = px.bar(title=title)
    fig.update_layout(
        template="plotly_white",
        xaxis={"visible": False},
        yaxis={"visible": False},
        annotations=[
            {
                "text": "Sin datos para mostrar",
                "xref": "paper",
                "yref": "paper",
                "showarrow": False,
                "font": {"size": 16},
            }
        ],
    )
    return fig


def load_data(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        return pd.DataFrame(
            columns=[
                "Fecha",
                "U$ FOB Tot",
                "Kg Neto",
                "Pais de Destino",
                "Presentación",
                "Partida Aduanera",
                "Exportador",
                "Año",
                "Mes",
                "Año-Mes",
            ]
        )

    df = pd.read_excel(file_path)
    df.columns = [str(col).strip() for col in df.columns]

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df = df.dropna(subset=["Fecha"]).copy()

    df["U$ FOB Tot"] = pd.to_numeric(df["U$ FOB Tot"], errors="coerce").fillna(0)
    df["Kg Neto"] = pd.to_numeric(df["Kg Neto"], errors="coerce").fillna(0)
    df["Pais de Destino"] = df["Pais de Destino"].fillna("Sin especificar")

    if "Presentación" in df.columns:
        df["Presentación"] = df["Presentación"].fillna("Sin especificar")
    elif "Presentacion" in df.columns:
        df["Presentación"] = df["Presentacion"].fillna("Sin especificar")
    elif len(df.columns) >= 25:
        # Columna Y (índice 24) según la estructura del Excel original.
        df["Presentación"] = df.iloc[:, 24].fillna("Sin especificar")
    else:
        df["Presentación"] = "Sin especificar"

    df["Año"] = df["Fecha"].dt.year
    df["Mes"] = df["Fecha"].dt.month
    df["Año-Mes"] = df["Fecha"].dt.to_period("M").dt.to_timestamp()

    return df


def format_number(value: float, decimals: int = 2) -> str:
    fmt = f"{{value:,.{decimals}f}}"
    return fmt.format(value=value).replace(",", "X").replace(".", ",").replace("X", ".")


def _build_yoy_variation_fig(monthly_df: pd.DataFrame, value_col: str, title: str):
    if monthly_df.empty:
        return _empty_fig(title)

    pivot = monthly_df.pivot_table(index="Mes", columns="Año", values=value_col, aggfunc="sum")
    if 2023 not in pivot.columns or 2024 not in pivot.columns:
        fig = _empty_fig(title)
        fig.update_layout(
            annotations=[
                {
                    "text": "No hay datos suficientes para comparar 2023 vs 2024",
                    "xref": "paper",
                    "yref": "paper",
                    "showarrow": False,
                    "font": {"size": 14},
                }
            ]
        )
        return fig

    month_labels = {
        1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
    }
    yoy_df = pivot[[2023, 2024]].reset_index()
    yoy_df["Variación % 24 vs 23"] = yoy_df.apply(
        lambda r: ((r[2024] - r[2023]) / r[2023] * 100) if r[2023] not in (0, None) else None,
        axis=1,
    )
    yoy_df["Mes Nombre"] = yoy_df["Mes"].map(month_labels)
    yoy_df = yoy_df.dropna(subset=["Variación % 24 vs 23"])

    if yoy_df.empty:
        fig = _empty_fig(title)
        fig.update_layout(
            annotations=[
                {
                    "text": "No hay base válida (año 2023 con valores > 0) para calcular variación",
                    "xref": "paper",
                    "yref": "paper",
                    "showarrow": False,
                    "font": {"size": 13},
                }
            ]
        )
        return fig

    fig = px.bar(
        yoy_df,
        x="Mes Nombre",
        y="Variación % 24 vs 23",
        title=title,
        text_auto=".1f",
        category_orders={"Mes Nombre": ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]},
    )
    fig.update_layout(template="plotly_white", xaxis_title="Mes", yaxis_title="Variación (%)", yaxis_tickformat=",.1f")
    fig.add_hline(y=0, line_dash="dash", line_color="#666")
    return fig


def _build_outputs(selected_years, selected_countries, selected_presentations):
    if df_data.empty:
        warning = "No se encontró CAMOTE.xlsx en la raíz del proyecto o no contiene datos válidos."
        empty = _empty_fig("Sin datos")
        table_columns = [
            {"name": "Pais de Destino", "id": "Pais de Destino"},
            {"name": "FOB total", "id": "FOB total"},
            {"name": "Kg Neto total", "id": "Kg Neto total"},
            {"name": "Participación FOB (%)", "id": "Participación FOB (%)"},
        ]
        return {
            "card_fob": [html.H4("FOB total"), html.H2("0,00")],
            "card_kg": [html.H4("Kg Neto total"), html.H2("0,00")],
            "fob_ts": empty,
            "kg_ts": empty,
            "pie": empty,
            "top5_fob": empty,
            "top5_kg": empty,
            "var_fob_2324": empty,
            "var_kg_2324": empty,
            "table_data": [],
            "table_columns": table_columns,
            "warning": warning,
            "filtered_df": pd.DataFrame(),
        }

    filtered = df_data.copy()
    if selected_years:
        filtered = filtered[filtered["Año"].isin(selected_years)]
    if selected_countries:
        filtered = filtered[filtered["Pais de Destino"].isin(selected_countries)]
    if selected_presentations:
        filtered = filtered[filtered["Presentación"].isin(selected_presentations)]

    total_fob = filtered["U$ FOB Tot"].sum()
    total_kg = filtered["Kg Neto"].sum()

    card_fob = [html.H4("FOB total"), html.H2(f"US$ {format_number(total_fob)}")]
    card_kg = [html.H4("Kg Neto total"), html.H2(format_number(total_kg))]

    monthly = (
        filtered.groupby(["Año", "Mes"], as_index=False)[["U$ FOB Tot", "Kg Neto"]]
        .sum()
        .sort_values(["Año", "Mes"])
    )

    month_labels = {
        1: "Ene",
        2: "Feb",
        3: "Mar",
        4: "Abr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Ago",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dic",
    }

    if monthly.empty:
        fob_ts = _empty_fig("Evolución temporal del FOB")
        kg_ts = _empty_fig("Evolución temporal del Kg Neto")
    else:
        monthly["Mes Nombre"] = monthly["Mes"].map(month_labels)

        fob_ts = px.line(
            monthly,
            x="Mes Nombre",
            y="U$ FOB Tot",
            color="Año",
            markers=True,
            category_orders={"Mes Nombre": list(month_labels.values())},
            title="Evolución mensual del FOB por Año",
        )
        fob_ts.update_layout(template="plotly_white", xaxis_title="Mes", yaxis_title="FOB", yaxis_tickformat=",.0f")

        kg_ts = px.line(
            monthly,
            x="Mes Nombre",
            y="Kg Neto",
            color="Año",
            markers=True,
            category_orders={"Mes Nombre": list(month_labels.values())},
            title="Evolución mensual del Kg Neto por Año",
        )
        kg_ts.update_layout(template="plotly_white", xaxis_title="Mes", yaxis_title="Kg Neto", yaxis_tickformat=",.0f")

    var_fob_2324 = _build_yoy_variation_fig(
        monthly, "U$ FOB Tot", "Evolución variación FOB 2023 - 2024 (%)"
    )
    var_kg_2324 = _build_yoy_variation_fig(
        monthly, "Kg Neto", "Evolución variación Kg Neto 2023 - 2024 (%)"
    )

    by_country = (
        filtered.groupby("Pais de Destino", as_index=False)[["U$ FOB Tot", "Kg Neto"]]
        .sum()
        .sort_values("U$ FOB Tot", ascending=False)
    )

    if by_country.empty or total_fob == 0:
        pie = _empty_fig("Participación FOB por País de Destino")
    else:
        top10_country = by_country.nlargest(10, "U$ FOB Tot").copy()
        remainder_fob = by_country["U$ FOB Tot"].sum() - top10_country["U$ FOB Tot"].sum()

        if remainder_fob > 0:
            top10_country = pd.concat(
                [
                    top10_country,
                    pd.DataFrame(
                        [{"Pais de Destino": "Los demás", "U$ FOB Tot": remainder_fob, "Kg Neto": 0}]
                    ),
                ],
                ignore_index=True,
            )

        top10_country["Participación FOB (%)"] = top10_country["U$ FOB Tot"] / total_fob * 100

        pie = px.pie(
            top10_country,
            names="Pais de Destino",
            values="Participación FOB (%)",
            title="Participación porcentual del FOB por País de Destino (Top 10 + Los demás)",
            hole=0.35,
        )
        pie.update_traces(
            textposition="outside",
            texttemplate="%{label}: %{value:.2f}%",
            showlegend=False,
        )
        pie.update_layout(template="plotly_white")

    top5_fob_df = by_country.nlargest(5, "U$ FOB Tot")
    top5_kg_df = by_country.nlargest(5, "Kg Neto")

    top5_fob = (
        _empty_fig("Top 5 países por FOB total")
        if top5_fob_df.empty
        else px.bar(
            top5_fob_df,
            x="Pais de Destino",
            y="U$ FOB Tot",
            title="Top 5 países por FOB total",
            text_auto=".2s",
        )
    )
    top5_fob.update_layout(template="plotly_white", xaxis_title="País", yaxis_title="FOB total", yaxis_tickformat=",.0f")

    top5_kg = (
        _empty_fig("Top 5 países por Kg Neto total")
        if top5_kg_df.empty
        else px.bar(
            top5_kg_df,
            x="Pais de Destino",
            y="Kg Neto",
            title="Top 5 países por Kg Neto total",
            text_auto=".2s",
        )
    )
    top5_kg.update_layout(template="plotly_white", xaxis_title="País", yaxis_title="Kg Neto total", yaxis_tickformat=",.0f")

    if not by_country.empty:
        by_country["Participación FOB (%)"] = (
            by_country["U$ FOB Tot"] / total_fob * 100 if total_fob > 0 else 0
        )

        table_df = by_country[["Pais de Destino", "U$ FOB Tot", "Kg Neto", "Participación FOB (%)"]].copy()
        table_df = table_df.rename(columns={"U$ FOB Tot": "FOB total", "Kg Neto": "Kg Neto total"})
        table_df["FOB total"] = table_df["FOB total"].map(lambda x: format_number(float(x), 2))
        table_df["Kg Neto total"] = table_df["Kg Neto total"].map(lambda x: format_number(float(x), 2))
        table_df["Participación FOB (%)"] = table_df["Participación FOB (%)"].map(
            lambda x: format_number(float(x), 2)
        )
        table_data = table_df.to_dict("records")
    else:
        table_data = []

    table_columns = [
        {"name": "Pais de Destino", "id": "Pais de Destino"},
        {"name": "FOB total", "id": "FOB total"},
        {"name": "Kg Neto total", "id": "Kg Neto total"},
        {"name": "Participación FOB (%)", "id": "Participación FOB (%)"},
    ]

    return {
        "card_fob": card_fob,
        "card_kg": card_kg,
        "fob_ts": fob_ts,
        "kg_ts": kg_ts,
        "pie": pie,
        "top5_fob": top5_fob,
        "top5_kg": top5_kg,
        "var_fob_2324": var_fob_2324,
        "var_kg_2324": var_kg_2324,
        "table_data": table_data,
        "table_columns": table_columns,
        "warning": "",
        "filtered_df": filtered,
    }


def build_pdf_report(selected_years, selected_countries, selected_presentations):
    outputs = _build_outputs(selected_years, selected_countries, selected_presentations)

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas
    except Exception:
        return None

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    page_w, page_h = A4

    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    filtered_df = outputs["filtered_df"]
    last_date = (
        filtered_df["Fecha"].max().strftime("%d/%m/%Y")
        if not filtered_df.empty
        else "Sin datos"
    )

    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(40, page_h - 40, "Reporte CAMOTE - Dashboard")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(40, page_h - 58, f"Generado: {generated_at}")
    pdf.drawString(40, page_h - 74, f"Último registro (filtro actual): {last_date}")

    chart_specs = [
        (outputs["fob_ts"], "Evolución FOB"),
        (outputs["kg_ts"], "Evolución Kg Neto"),
        (outputs["pie"], "Participación FOB"),
        (outputs["top5_fob"], "Top 5 FOB"),
        (outputs["top5_kg"], "Top 5 Kg Neto"),
        (outputs["var_fob_2324"], "Variación FOB 2023-2024"),
        (outputs["var_kg_2324"], "Variación Kg Neto 2023-2024"),
    ]

    x_positions = [35, page_w / 2 + 5]
    y_top = page_h - 100
    chart_w = (page_w - 80) / 2
    chart_h = 210

    for i, (fig, title) in enumerate(chart_specs):
        if i > 0 and i % 4 == 0:
            pdf.showPage()
            y_top = page_h - 50

        panel = i % 4
        row = panel // 2
        col = panel % 2

        x = x_positions[col]
        y = y_top - row * (chart_h + 40) - chart_h

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(x, y + chart_h + 10, title)

        img_bytes = pio.to_image(fig, format="png", width=900, height=500, scale=1)
        img = ImageReader(io.BytesIO(img_bytes))
        pdf.drawImage(img, x, y, width=chart_w, height=chart_h, preserveAspectRatio=True, mask="auto")

    pdf.save()
    buffer.seek(0)
    return buffer.read()


df_data = load_data(FILE_PATH)
years = sorted(df_data["Año"].dropna().unique()) if not df_data.empty else []
countries = sorted(df_data["Pais de Destino"].dropna().unique()) if not df_data.empty else []
presentations = sorted(df_data["Presentación"].dropna().unique()) if not df_data.empty else []
last_record_date_text = (
    df_data["Fecha"].max().strftime("%d/%m/%Y") if not df_data.empty else "Sin datos"
)

app = dash.Dash(__name__)
app.title = "Dashboard CAMOTE"

app.layout = html.Div(
    style={"fontFamily": "Arial, sans-serif", "padding": "20px", "backgroundColor": "#f7f9fc"},
    children=[
        html.Div(
            style={
                "display": "flex",
                "justifyContent": "space-between",
                "alignItems": "flex-start",
                "gap": "14px",
                "flexWrap": "wrap",
                "marginBottom": "8px",
            },
            children=[
                html.Div(
                    children=[
                        html.H1("Dashboard de Exportaciones - CAMOTE", style={"marginBottom": "5px"}),
                        html.P(
                            "Análisis de FOB, Kg Neto y participación por país de destino.",
                            style={"color": "#555", "marginTop": "0"},
                        ),
                    ]
                ),
                html.Div(
                    style={
                        "background": "white",
                        "borderRadius": "8px",
                        "padding": "10px 14px",
                        "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
                        "minWidth": "250px",
                    },
                    children=[
                        html.Div("Último registro", style={"fontSize": "12px", "color": "#666"}),
                        html.Div(
                            last_record_date_text,
                            style={"fontWeight": "bold", "fontSize": "18px", "marginTop": "2px"},
                        ),
                        html.Button(
                            "Descargar toda la información en PDF",
                            id="download-pdf-btn",
                            style={
                                "marginTop": "10px",
                                "width": "100%",
                                "padding": "8px",
                                "border": "none",
                                "borderRadius": "6px",
                                "backgroundColor": "#0d6efd",
                                "color": "white",
                                "cursor": "pointer",
                            },
                        ),
                        dcc.Download(id="download-pdf"),
                    ],
                ),
            ],
        ),
        html.Div(
            style={"display": "flex", "gap": "16px", "marginBottom": "20px", "flexWrap": "wrap"},
            children=[
                html.Div(
                    style={"minWidth": "230px", "flex": "1"},
                    children=[
                        html.Label("Filtro por Año"),
                        dcc.Dropdown(
                            id="year-filter",
                            options=[{"label": str(y), "value": int(y)} for y in years],
                            value=years,
                            multi=True,
                            placeholder="Seleccione año(s)",
                        ),
                    ],
                ),
                html.Div(
                    style={"minWidth": "260px", "flex": "2"},
                    children=[
                        html.Label("Filtro por Pais de Destino"),
                        dcc.Dropdown(
                            id="country-filter",
                            options=[{"label": c, "value": c} for c in countries],
                            value=countries,
                            multi=True,
                            placeholder="Seleccione país(es)",
                        ),
                    ],
                ),
                html.Div(
                    style={"minWidth": "260px", "flex": "2"},
                    children=[
                        html.Label("Filtro por Presentación"),
                        dcc.Dropdown(
                            id="presentation-filter",
                            options=[{"label": p, "value": p} for p in presentations],
                            value=presentations,
                            multi=True,
                            placeholder="Seleccione presentación(es)",
                        ),
                    ],
                ),
            ],
        ),
        html.Div(id="data-warning", style={"color": "#b00020", "marginBottom": "16px"}),
        html.Div(
            style={
                "display": "grid",
                "gridTemplateColumns": "repeat(auto-fit, minmax(220px, 1fr))",
                "gap": "12px",
                "marginBottom": "18px",
            },
            children=[
                html.Div(
                    id="card-fob",
                    style={
                        "background": "white",
                        "padding": "16px",
                        "borderRadius": "8px",
                        "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
                    },
                ),
                html.Div(
                    id="card-kg",
                    style={
                        "background": "white",
                        "padding": "16px",
                        "borderRadius": "8px",
                        "boxShadow": "0 1px 4px rgba(0,0,0,.08)",
                    },
                ),
            ],
        ),
        html.Div(
            style={"display": "grid", "gridTemplateColumns": "1fr", "gap": "18px"},
            children=[
                dcc.Graph(id="fob-time-series"),
                dcc.Graph(id="kg-time-series"),
                dcc.Graph(id="fob-country-share"),
                html.Div(
                    style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px"},
                    children=[dcc.Graph(id="top5-fob"), dcc.Graph(id="top5-kg")],
                ),
                html.Div(
                    style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "16px"},
                    children=[dcc.Graph(id="var-fob-2324"), dcc.Graph(id="var-kg-2324")],
                ),
            ],
        ),
        html.H3("Tabla resumen por país", style={"marginTop": "24px"}),
        dash_table.DataTable(
            id="summary-table",
            page_size=10,
            style_table={"overflowX": "auto", "backgroundColor": "white"},
            style_cell={"padding": "8px", "textAlign": "left", "fontFamily": "Arial"},
            style_cell_conditional=[
                {"if": {"column_id": "FOB total"}, "textAlign": "right"},
                {"if": {"column_id": "Kg Neto total"}, "textAlign": "right"},
                {"if": {"column_id": "Participación FOB (%)"}, "textAlign": "right"},
            ],
            style_header={"backgroundColor": "#e9edf5", "fontWeight": "bold"},
        ),
    ],
)


@app.callback(
    Output("card-fob", "children"),
    Output("card-kg", "children"),
    Output("fob-time-series", "figure"),
    Output("kg-time-series", "figure"),
    Output("fob-country-share", "figure"),
    Output("top5-fob", "figure"),
    Output("top5-kg", "figure"),
    Output("var-fob-2324", "figure"),
    Output("var-kg-2324", "figure"),
    Output("summary-table", "data"),
    Output("summary-table", "columns"),
    Output("data-warning", "children"),
    Input("year-filter", "value"),
    Input("country-filter", "value"),
    Input("presentation-filter", "value"),
)
def update_dashboard(selected_years, selected_countries, selected_presentations):
    outputs = _build_outputs(selected_years, selected_countries, selected_presentations)
    return (
        outputs["card_fob"],
        outputs["card_kg"],
        outputs["fob_ts"],
        outputs["kg_ts"],
        outputs["pie"],
        outputs["top5_fob"],
        outputs["top5_kg"],
        outputs["var_fob_2324"],
        outputs["var_kg_2324"],
        outputs["table_data"],
        outputs["table_columns"],
        outputs["warning"],
    )


@app.callback(
    Output("download-pdf", "data"),
    Input("download-pdf-btn", "n_clicks"),
    State("year-filter", "value"),
    State("country-filter", "value"),
    State("presentation-filter", "value"),
    prevent_initial_call=True,
)
def download_pdf(_, selected_years, selected_countries, selected_presentations):
    pdf_bytes = build_pdf_report(selected_years, selected_countries, selected_presentations)
    if pdf_bytes is None:
        fallback = (
            "No fue posible generar el PDF con gráficos porque faltan dependencias "
            "de exportación (reportlab y/o kaleido)."
        )
        return {
            "content": fallback,
            "filename": "reporte_camote_error.txt",
            "type": "text/plain",
        }

    pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")
    return {
        "content": pdf_base64,
        "filename": "reporte_camote.pdf",
        "type": "application/pdf",
        "base64": True,
    }


if __name__ == "__main__":
    app.run(debug=True)
