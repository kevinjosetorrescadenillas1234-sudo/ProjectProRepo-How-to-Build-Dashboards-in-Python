import base64
import io

import dash
from dash import Dash, Input, Output, State, dcc, html
import pandas as pd
import plotly.express as px


app = Dash(__name__)
app.title = "Dashboard desde Excel"


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _encontrar_columna(df: pd.DataFrame, opciones: list[str]) -> str | None:
    for opcion in opciones:
        if opcion in df.columns:
            return opcion
    return None


def _primera_columna_categorica(df: pd.DataFrame) -> str | None:
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            return col
    return None


def _columna_pais(df: pd.DataFrame) -> str | None:
    return _encontrar_columna(
        df,
        ["pais", "país", "paises", "países", "country", "countries"],
    )


app.layout = html.Div(
    [
        html.H1("Dashboard básico desde Excel"),
        html.P(
            "Sube un archivo Excel (.xlsx o .xls) con columnas como: "
            "participación, fob total y kg nt totales."
        ),
        dcc.Upload(
            id="upload-excel",
            children=html.Div(["Arrastra o selecciona tu archivo Excel"]),
            style={
                "width": "100%",
                "height": "90px",
                "lineHeight": "90px",
                "borderWidth": "1px",
                "borderStyle": "dashed",
                "borderRadius": "10px",
                "textAlign": "center",
                "marginBottom": "20px",
            },
            multiple=False,
        ),
        html.Div(id="mensaje"),
        html.Div(
            id="tarjetas",
            style={
                "display": "grid",
                "gridTemplateColumns": "repeat(auto-fit, minmax(220px, 1fr))",
                "gap": "12px",
                "marginTop": "10px",
            },
        ),
        dcc.Graph(id="grafico-participacion"),
        dcc.Graph(id="grafico-fob"),
        dcc.Graph(id="grafico-kg"),
    ],
    style={"maxWidth": "1100px", "margin": "0 auto", "padding": "24px"},
)


@app.callback(
    Output("mensaje", "children"),
    Output("tarjetas", "children"),
    Output("grafico-participacion", "figure"),
    Output("grafico-fob", "figure"),
    Output("grafico-kg", "figure"),
    Input("upload-excel", "contents"),
    State("upload-excel", "filename"),
)
def procesar_archivo(contents, filename):
    fig_vacia = px.scatter(title="Sin datos cargados")

    if not contents:
        return (
            html.Div("Esperando archivo Excel..."),
            [],
            fig_vacia,
            fig_vacia,
            fig_vacia,
        )

    try:
        _, content_string = contents.split(",")
        decoded = base64.b64decode(content_string)
        df = pd.read_excel(io.BytesIO(decoded))
        df = _normalizar_columnas(df)

        col_fob = _encontrar_columna(df, ["fob total", "fob", "valor fob", "total fob"])
        col_kg = _encontrar_columna(df, ["kg nt totales", "kg nt total", "kg", "kilogramos"])
        col_categoria = _primera_columna_categorica(df)
        col_pais = _columna_pais(df)

        if not col_categoria:
            return (
                html.Div("No se encontró ninguna columna categórica para segmentar gráficos."),
                [],
                fig_vacia,
                fig_vacia,
                fig_vacia,
            )

        tarjetas = []
        if col_fob:
            fob_total = pd.to_numeric(df[col_fob], errors="coerce").fillna(0).sum()
            tarjetas.append(
                html.Div(
                    [html.H4("FOB Total"), html.P(f"{fob_total:,.2f}")],
                    style={"padding": "12px", "border": "1px solid #ddd", "borderRadius": "8px"},
                )
            )

        if col_kg:
            kg_total = pd.to_numeric(df[col_kg], errors="coerce").fillna(0).sum()
            tarjetas.append(
                html.Div(
                    [html.H4("KG NT Totales"), html.P(f"{kg_total:,.2f}")],
                    style={"padding": "12px", "border": "1px solid #ddd", "borderRadius": "8px"},
                )
            )

        # Participación calculada por FOB:
        # (suma FOB por país / FOB total de todos los países) * 100
        if col_fob:
            categoria_participacion = col_pais or col_categoria
            part_df = (
                df[[categoria_participacion, col_fob]]
                .assign(**{col_fob: pd.to_numeric(df[col_fob], errors="coerce").fillna(0)})
                .groupby(categoria_participacion, as_index=False)[col_fob]
                .sum()
            )
            total_fob = part_df[col_fob].sum()
            part_df["participacion"] = (part_df[col_fob] / total_fob * 100) if total_fob else 0
            titulo_participacion = "Participación por país (%)" if col_pais else "Participación por categoría (%)"
            fig_part = px.pie(
                part_df,
                values="participacion",
                names=categoria_participacion,
                title=titulo_participacion,
            )
        else:
            fig_part = px.scatter(title="No se pudo calcular participación: falta columna FOB")

        # FOB
        if col_fob:
            categoria_fob = col_pais or col_categoria
            fob_df = (
                df[[categoria_fob, col_fob]]
                .assign(**{col_fob: pd.to_numeric(df[col_fob], errors="coerce").fillna(0)})
                .groupby(categoria_fob, as_index=False)[col_fob]
                .sum()
                .sort_values(by=col_fob, ascending=False)
            )
            titulo_fob = "FOB Total por país" if col_pais else "FOB Total por categoría"
            fig_fob = px.bar(fob_df, x=categoria_fob, y=col_fob, title=titulo_fob)
        else:
            fig_fob = px.scatter(title="No se encontró columna FOB")

        # KG
        if col_kg:
            categoria_kg = col_pais or col_categoria
            kg_df = (
                df[[categoria_kg, col_kg]]
                .assign(**{col_kg: pd.to_numeric(df[col_kg], errors="coerce").fillna(0)})
                .groupby(categoria_kg, as_index=False)[col_kg]
                .sum()
                .sort_values(by=col_kg, ascending=False)
            )
            titulo_kg = "KG NT Totales por país" if col_pais else "KG NT Totales por categoría"
            fig_kg = px.bar(kg_df, x=categoria_kg, y=col_kg, title=titulo_kg)
        else:
            fig_kg = px.scatter(title="No se encontró columna KG NT")

        mensaje = html.Div(f"Archivo cargado: {filename} | Filas leídas: {len(df)}")
        return mensaje, tarjetas, fig_part, fig_fob, fig_kg

    except Exception as exc:
        return (
            html.Div(f"Error al procesar el archivo: {exc}"),
            [],
            fig_vacia,
            fig_vacia,
            fig_vacia,
        )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
