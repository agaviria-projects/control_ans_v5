"""
MAPA ANS PROFESIONAL – v6.3 (Ultra-Blindado)
Google Maps + Tooltip Dirección
Héctor + IA – 2025
"""

import pandas as pd
import folium
from branca.element import Template, MacroElement
from pathlib import Path
import webbrowser
import re

# ============================================================
# 1. RUTAS
# ============================================================
script_path = Path(__file__).resolve()
base_path = script_path.parent

ruta_fenix = base_path / "data_clean" / "FENIX_ANS.xlsx"
ruta_salida = base_path / "data_output" / "mapa_ans.html"

# ============================================================
# 2. CARGAR EXCEL
# ============================================================
df = pd.read_excel(ruta_fenix, sheet_name="FENIX_ANS", dtype=str)
df.columns = df.columns.str.upper().str.strip()

# ============================================================
# 2.1 NORMALIZAR ESTADOS (ULTRA BLINDADO)
# ============================================================
def normalizar_estado(e):
    if not isinstance(e, str):
        return "SIN FECHA"

    # eliminar caracteres invisibles
    e = re.sub(r"[\u200B-\u200D\uFEFF\u00A0]", "", e)

    e = (
        e.upper().strip()
         .replace("Í", "I")
         .replace(" 0 DIAS", "_0 DIAS")
         .replace("0 DIAS", "_0 DIAS")
         .replace("SIN DATO", "SIN FECHA")
    )

    e = e.replace("  ", " ").replace("\r", "").replace("\n", "")

    ESTADOS_VALIDOS = {
        "A TIEMPO",
        "ALERTA",
        "ALERTA_0 DIAS",
        "VENCIDO",
        "SIN FECHA"
    }

    if e in ESTADOS_VALIDOS:
        return e

    return "SIN FECHA"

df["ESTADO"] = df["ESTADO"].apply(normalizar_estado)

# ============================================================
# 2.2 LIMPIAR COORDENADAS
# ============================================================
def limpiar_coord(x):
    if x is None: 
        return None
    x = str(x).strip().replace(",", ".")
    try:
        return float(x)
    except:
        return None

df["COORDENADAX"] = df["COORDENADAX"].apply(limpiar_coord)
df["COORDENADAY"] = df["COORDENADAY"].apply(limpiar_coord)

df = df.dropna(subset=["COORDENADAX", "COORDENADAY"])

# ============================================================
# 2.3 ELIMINAR DUPLICADOS SOLO PARA EL MAPA (NO AFECTA EL EXCEL)
# ============================================================
df_mapa = df.drop_duplicates(
    subset=["PEDIDO", "COORDENADAX", "COORDENADAY"],
    keep="first"
)

print(f"📌 Total pedidos visibles en el mapa: {len(df_mapa)}")

# ============================================================
# 3. MAPA BASE (GOOGLE MAPS SATÉLITE + ETIQUETAS)
# ============================================================
mapa = folium.Map(
    location=[6.24, -75.57],
    zoom_start=13,
    tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
    attr="Google"
)

mapa_id = mapa.get_name()

mapa.get_root().html.add_child(folium.Element(f"""
<script>
document.addEventListener("DOMContentLoaded", function() {{
    window.mapa = {mapa_id};
    window.marcadores = {{}};
    window.estadoMarcadores = {{
        "A TIEMPO": [],
        "ALERTA": [],
        "ALERTA_0 DIAS": [],
        "VENCIDO": [],
        "SIN FECHA": []
    }};
}});
</script>
"""))

# ============================================================
# 4. ICONOS
# ============================================================
ICON_SIZE = [20, 33]

colores = {
    "A TIEMPO": "green",
    "ALERTA": "yellow",
    "ALERTA_0 DIAS": "orange",
    "VENCIDO": "red",
    "SIN FECHA": "grey"
}

# ============================================================
# 5. CREAR MARCADORES (CON TOOLTIP DIRECCIÓN)
# ============================================================
markers_js = """
<script>
document.addEventListener("DOMContentLoaded", function() {
    window.marcadores = {};
    window.estadoMarcadores = {
        "A TIEMPO": [],
        "ALERTA": [],
        "ALERTA_0 DIAS": [],
        "VENCIDO": [],
        "SIN FECHA": []
    };
"""

for _, row in df_mapa.iterrows():
    pedido = row["PEDIDO"]
    estado = row["ESTADO"]
    lat = row["COORDENADAY"]
    lon = row["COORDENADAX"]
    direccion = row["DIRECCION"] if "DIRECCION" in df.columns else "SIN DIRECCIÓN"

    popup = f"""
<b>PEDIDO:</b> {pedido}<br>
<b>ESTADO:</b> {estado}<br>
<b>DIRECCIÓN:</b> {direccion}
"""

    color = colores.get(estado, "grey")

    markers_js += f"""
var mk_{pedido} = L.marker([{lat}, {lon}], {{
    icon: L.icon({{
        iconUrl: "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-{color}.png",
        shadowUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-shadow.png",
        iconSize: [{ICON_SIZE[0]}, {ICON_SIZE[1]}],
        iconAnchor: [10, 33],
        popupAnchor: [0, -28]
    }})
}}).bindTooltip("{direccion}", {{permanent:false}}).bindPopup(`{popup}`).addTo(window.mapa);

window.marcadores["{pedido}"] = mk_{pedido};
window.estadoMarcadores["{estado}"].push("{pedido}");
"""

markers_js += """
});
</script>
"""

mapa.get_root().html.add_child(folium.Element(markers_js))

# ============================================================
# 6. PANEL LATERAL
# ============================================================
panel_html = Template("""
{% macro html(this, kwargs) %}

<style>
#panelANS{
    position: fixed;
    right:20px;
    top:20px;
    width:240px;
    background:white;
    padding:15px;
    border-radius:12px;
    box-shadow:0 0 12px rgba(0,0,0,0.3);
    z-index:999999;
    font-family:Arial;
}
.filtroBtn{
    width:100%;
    padding:8px;
    border-radius:6px;
    margin-top:6px;
    cursor:pointer;
    font-weight:bold;
    text-align:center;
}
</style>

<div id="panelANS">
<b style="font-size:18px;">📊 ANS Control</b><br><br>

<b>Buscar pedido:</b><br>
<input id="buscarPedido" style="width:100%;padding:6px;"><br>
<button onclick="buscarPedido()" class="filtroBtn" style="background:#e0e0e0;">Buscar</button>
<button onclick="limpiarBusqueda()" class="filtroBtn" style="background:#cccccc;">Limpiar</button>

<hr>
<b>Filtros:</b><br>

<div onclick="filtrarEstado('A TIEMPO')" class="filtroBtn" style="background:#00C853;color:white;">A TIEMPO</div>
<div onclick="filtrarEstado('ALERTA')" class="filtroBtn" style="background:#FFD600;">ALERTA</div>
<div onclick="filtrarEstado('ALERTA_0 DIAS')" class="filtroBtn" style="background:#FF8F00;">ALERTA 0 DÍAS</div>
<div onclick="filtrarEstado('VENCIDO')" class="filtroBtn" style="background:#D50000;color:white;">VENCIDO</div>
<div onclick="filtrarEstado('SIN FECHA')" class="filtroBtn" style="background:#616161;color:white;">SIN FECHA</div>
<div onclick="mostrarTodos()" class="filtroBtn" style="background:#bbdefb;">MOSTRAR TODOS</div>
</div>

<script>

function ocultarTodos(){
    Object.values(window.marcadores).forEach(m => window.mapa.removeLayer(m));
}

function mostrarTodos(){
    Object.values(window.marcadores).forEach(m => window.mapa.addLayer(m));
}

function filtrarEstado(estado){
    ocultarTodos();
    window.estadoMarcadores[estado].forEach(p => {
        window.mapa.addLayer(window.marcadores[p]);
    });
}

function buscarPedido(){
    let p = document.getElementById("buscarPedido").value.trim();
    if(!p){ mostrarTodos(); return; }

    if(window.marcadores[p]){
        ocultarTodos();
        let mk = window.marcadores[p];
        window.mapa.addLayer(mk);
        window.mapa.setView(mk.getLatLng(), 18);
        mk.openPopup();
    }else{
        alert("Pedido no encontrado");
    }
}

function limpiarBusqueda(){
    document.getElementById("buscarPedido").value = "";
    mostrarTodos();
}

</script>

{% endmacro %}
""")

panel = MacroElement()
panel._template = panel_html
mapa.get_root().add_child(panel)

# ============================================================
# 7. GUARDAR MAPA
# ============================================================
ruta_salida.parent.mkdir(exist_ok=True)
mapa.save(ruta_salida)

print("🟢 Mapa ANS v6.3 generado correctamente.")
webbrowser.open(str(ruta_salida))
