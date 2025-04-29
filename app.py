from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

def procesar_excel(file):
    # Leer el archivo Excel
    df = pd.read_excel(file)

    # Filtrar días laborados (excluyendo permisos, incapacidades y vacaciones)
    
    df_laborado = df[~df["ACTIVIDAD"].isin(["PERMISO", "INCAPACIDAD", "VACACIONES","FUNC ADMON","RETIRO"])]

    # Contar los días laborados por cada inspector
    dias_laborados = df_laborado.groupby("CEDULA INSPECTOR")["FECHA"].nunique().reset_index()
    dias_laborados.columns = ["CEDULA INSPECTOR", "Total Dias Laborados"]

    # Contar total de inspecciones por inspector
    total_inspecciones = df.groupby("CEDULA INSPECTOR")["TOTAL REVISIONES"].sum().reset_index()
    total_inspecciones.columns = ["CEDULA INSPECTOR", "TOTAL_INSPECCIONES"]

    # Calcular total de "LM" por inspector
    total_lm = df.groupby("CEDULA INSPECTOR")["LM"].sum().reset_index()
    total_lm.columns = ["CEDULA INSPECTOR", "Total_LM"]

    # Calcular el total por suspensiones (3,000 pesos por cada suspensión)
    total_suspensiones = df.groupby("CEDULA INSPECTOR")["TOTAL SUSPENSIONES"].sum() * 3000

    # Filtrar solo el personal operativo para el auxilio de moto
    df_operativo = df[df["CENTRO DE VINCULACIÓN"].str.contains("operativo", case=False, na=False)]

    # Calcular el auxilio de moto solo para personal operativo
    dias_laborados_operativo = dias_laborados[dias_laborados["CEDULA INSPECTOR"].isin(df_operativo["CEDULA INSPECTOR"])]
    auxilio_moto = dias_laborados_operativo.set_index("CEDULA INSPECTOR")["Total Dias Laborados"] * 22000

    # Crear DataFrame de bonos
    df_bonos = df.groupby("CEDULA INSPECTOR").agg(
        NOMBRE_INSPECTOR=('NOMBRE INSPECTOR', 'first'),
        CENTRO_DE_VINCULACION=('CENTRO DE VINCULACIÓN', 'first'),
        TOTAL_INSPECCIONES=('TOTAL REVISIONES', 'sum'),
        TOTAL_SUSPENSIONES=('TOTAL SUSPENSIONES', 'sum')
    ).reset_index()

    # Unir los días laborados y el total de LM
    df_bonos = (
        df_bonos
        .merge(dias_laborados, on="CEDULA INSPECTOR", how="left")
        .merge(total_lm, on="CEDULA INSPECTOR", how="left")
    )

    # Calcular bonos de gestión y adicionales
    df_bonos["Bono_Gestion"] = df_bonos["TOTAL_INSPECCIONES"].apply(calcular_bono_gestion)
    df_bonos["Bono_Adicional"] = df_bonos["TOTAL_INSPECCIONES"].apply(calcular_bono_adicional)

    # Agregar auxilio de moto y suspensiones
    df_bonos["Auxilio_Moto"] = df_bonos["Total Dias Laborados"] * 22000
    df_bonos["Auxilio_Suspensiones"] = df_bonos["TOTAL_SUSPENSIONES"] * 3000

    # Calcular el total a liquidar por inspector
    df_bonos["Bono_Total"] = df_bonos["Bono_Gestion"] + df_bonos["Bono_Adicional"]
    df_bonos["Auxilio_Total"] = df_bonos["Auxilio_Moto"] + df_bonos["Auxilio_Suspensiones"]

    # Categorizar los inspectores según el total de inspecciones
    df_bonos["Categoria"] = df_bonos["TOTAL_INSPECCIONES"].apply(categorizar_inspector)

    # Seleccionar columnas y ajustar el formato
    output_df = df_bonos[[
        "CENTRO_DE_VINCULACION", "CEDULA INSPECTOR", "NOMBRE_INSPECTOR", 
        "Total Dias Laborados", "TOTAL_SUSPENSIONES", "TOTAL_INSPECCIONES", 
        "Total_LM", "Bono_Gestion", "Bono_Adicional", "Auxilio_Moto", 
        "Auxilio_Suspensiones", "Bono_Total", "Auxilio_Total", "Categoria"
    ]]

    # Guardar el resultado en un archivo Excel
    output = BytesIO()
    output_df.to_excel(output, index=False, sheet_name="Liquidacion")
    output.seek(0)
    return output

def calcular_bono_gestion(inspecciones):
    if inspecciones > 210:
        return (inspecciones - 160) * 15000
    elif inspecciones > 180:
        return (inspecciones - 160) * 13000
    elif inspecciones > 160:
        return (inspecciones - 160) * 10000
    else:
        return 0

def calcular_bono_adicional(inspecciones):
    if inspecciones > 250:
        return 500000
    elif inspecciones > 230:
        return 330000
    elif inspecciones > 210:
        return 180000
    else:
        return 0

def categorizar_inspector(inspecciones):
    if inspecciones > 250:
        return "ORO"
    elif inspecciones > 230:
        return "PLATA"
    elif inspecciones > 210 :
        return "BRONCE"
    else:
        return "SIN CATEGORIA"

@app.route("/nomina", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        processed_file = procesar_excel(file)
        return send_file(processed_file, download_name="Liquidacion__Bonificaciones.xlsx", as_attachment=True)
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
