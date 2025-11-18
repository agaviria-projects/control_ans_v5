"""
MAPA ANS PROFESIONAL – v7.4 (ULTRA-ESTABLE)
Google Maps + Panel ANS + Modal Elegante
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
# 2.1 NORMALIZAR ESTADOS (VERSIÓN ULTRA-ROBUSTA)
# ============================================================
def normalizar_estado(e):
    if not isinstance(e, str):
        return "SIN FECHA"

    # Eliminar caracteres invisibles
    e = re.sub(r"[\u200B-\u200D\uFEFF\u00A0]", "", e)

    # Normalizar tildes y espacios
    e = (
        e.upper()
        .replace("Í", "I")
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("SIN DATO", "SIN FECHA")
        .strip()
    )

    # Unificar variantes de ALERTA 0 DIAS
    if "ALERTA" in e and ("0" in e or " 0 " in e):
        return "ALERTA_0 DIAS"

    # ALERTA normal
    if "ALERTA" in e:
        return "ALERTA"

    # A TIEMPO
    if "TIEMPO" in e:
        return "A TIEMPO"

    # VENCIDO
    if "VENCID" in e:
        return "VENCIDO"

    # SIN FECHA final
    return "SIN FECHA"

df["ESTADO"] = df["ESTADO"].apply(normalizar_estado)

# ============================================================
# 2.2 LIMPIAR COORDENADAS
# ============================================================
def limpiar_coord(x):
    if x is None:
        return None
    x = str(x).replace(",", ".").strip()
    try:
        return float(x)
    except:
        return None

df["COORDENADAX"] = df["COORDENADAX"].apply(limpiar_coord)
df["COORDENADAY"] = df["COORDENADAY"].apply(limpiar_coord)

df = df.dropna(subset=["COORDENADAX", "COORDENADAY"])

# ============================================================
# 2.3 ELIMINAR DUPLICADOS
# ============================================================
df_mapa = df.drop_duplicates(
    subset=["PEDIDO", "COORDENADAX", "COORDENADAY"],
    keep="first"
)

print(f"[INFO] Total pedidos visibles en el mapa: {len(df_mapa)}")


# ============================================================
# 3. MAPA BASE (SATÉLITE)
# ============================================================
mapa = folium.Map(
    location=[6.24, -75.57],
    zoom_start=13,
    tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
    attr="Google"
)

mapa_id = mapa.get_name()

# ============================================================
# 3.1 Variables JS Globales
# ============================================================
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
    "SIN FECHA": "violet"
}

# ============================================================
# 5. MARCADORES
# ============================================================
markers_js = """
<script>
document.addEventListener("DOMContentLoaded", function() {
"""

for _, row in df_mapa.iterrows():
    pedido = row["PEDIDO"]
    estado = row["ESTADO"]
    lat = row["COORDENADAY"]
    lon = row["COORDENADAX"]

    color = colores.get(estado, "grey")

    markers_js += f"""
var mk_{pedido} = L.marker([{lat}, {lon}], {{
    icon: L.icon({{
        iconUrl: "https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-{color}.png",
        shadowUrl: "https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.7.1/images/marker-shadow.png",
        iconSize: [{ICON_SIZE[0]}, {ICON_SIZE[1]}],
        iconAnchor: [10, 33], popupAnchor: [0, -28]
    }})
}}).bindTooltip("{pedido}")
  .addTo(window.mapa);

window.marcadores["{pedido}"] = mk_{pedido};
window.estadoMarcadores["{estado}"].push("{pedido}");
"""

markers_js += """
});
</script>
"""

mapa.get_root().html.add_child(folium.Element(markers_js))

# ============================================================
# 6. PANEL LATERAL + MODAL
# ============================================================
panel_html = Template("""
{% macro html(this, kwargs) %}

<style>
#panelANS{
    position: fixed;
    right:20px; top:20px;
    width:240px;
    background:white;
    padding:15px;
    border-radius:12px;
    box-shadow:0 0 12px rgba(0,0,0,0.3);
    z-index:999999;
    font-family:Arial;
}
.filtroBtn{
    width:100%; padding:8px;
    border-radius:6px;
    margin-top:6px;
    cursor:pointer;
    font-weight:bold;
    text-align:center;
}
.subtitulo{
    font-size:14px; margin-top:12px;
    font-weight:bold;
}
</style>

<div id="panelANS">

<div id="modalError"
     style="display:none; position:fixed;
     top:50%; left:50%; transform:translate(-50%, -50%);
     background:white; padding:20px;
     border-radius:10px; width:250px;
     box-shadow:0 0 20px rgba(0,0,0,0.3);
     z-index:9999999; text-align:center;">
</div>

<b style="font-size:18px;">📊 Control ANS</b><br><br>

<b>Buscar pedido:</b><br>
<input id="buscarPedido" style="width:100%;padding:6px;"><br>
<button onclick="buscarPedido()" class="filtroBtn" style="background:#e0e0e0;">Buscar</button>
<button onclick="limpiarBusqueda()" class="filtroBtn" style="background:#cccccc;">Limpiar</button>

<hr>

<b class="subtitulo">Filtros:</b>
<div onclick="filtrarEstado('A TIEMPO')" class="filtroBtn" style="background:#00C853;color:white;">A TIEMPO</div>
<div onclick="filtrarEstado('ALERTA')" class="filtroBtn" style="background:#FFD600;">ALERTA</div>
<div onclick="filtrarEstado('ALERTA_0 DIAS')" class="filtroBtn" style="background:#FF8F00;">ALERTA 0 DÍAS</div>
<div onclick="filtrarEstado('VENCIDO')" class="filtroBtn" style="background:#D50000;color:white;">VENCIDO</div>
<div onclick="filtrarEstado('SIN FECHA')" class="filtroBtn" style="background:#6a1b9a;color:white;">SIN FECHA</div>
<div onclick="mostrarTodos()" class="filtroBtn" style="background:#bbdefb;">MOSTRAR TODOS</div>

<hr>

<b class="subtitulo">🗺️ Capas del Mapa</b>
<div onclick="setCapa('sat')" class="filtroBtn" style="background:#c5e1a5;">Satélite</div>
<div onclick="setCapa('calles')" class="filtroBtn" style="background:#aed581;">Vista Urbana</div>

</div>

<script>
// ======= CAPAS GOOGLE =======
const capas = {
    sat: "https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
    calles: "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}"
};

window.currentLayer = null;

window.setCapa = function(tipo){
    if(window.currentLayer){ window.mapa.removeLayer(window.currentLayer); }
    window.currentLayer = L.tileLayer(capas[tipo], {maxZoom: 20}).addTo(window.mapa);
    window.mapa.invalidateSize(true);
};

// ======= MODAL =======
function mostrarModal(msg){
    let m = document.getElementById("modalError");
    m.style.display = "block";
    m.innerHTML = `
        <b style="color:#b71c1c;">❗ ${msg}</b><br><br>
        <button onclick="cerrarModal()" 
                style="background:#b71c1c;color:white;padding:6px 12px;
                       border:none;border-radius:6px;cursor:pointer;">
            Cerrar
        </button>
    `;
    setTimeout(()=>{ cerrarModal(); }, 2500);
}

function cerrarModal(){
    document.getElementById("modalError").style.display = "none";
}

// ======= FILTROS Y BÚSQUEDA =======
setTimeout(function(){

    function refrescar(){ setTimeout(()=> window.mapa.invalidateSize(true), 30); }

    window.ocultarTodos = ()=>{ 
        Object.values(window.marcadores).forEach(m=>m.setOpacity(0)); 
        refrescar(); 
    };

    window.mostrarTodos = ()=>{
        Object.values(window.marcadores).forEach(m=>m.setOpacity(1));
        window.mapa.setView([6.24, -75.57], 13);
        refrescar();
    };

    window.filtrarEstado = (estado)=>{
        window.ocultarTodos();
        window.estadoMarcadores[estado].forEach(p=> window.marcadores[p].setOpacity(1));
        refrescar();
    };

    window.buscarPedido = ()=>{
        let p = document.getElementById("buscarPedido").value.trim();
        if(!p) return;

        if(window.marcadores[p]){
            window.ocultarTodos();
            let mk = window.marcadores[p];
            mk.setOpacity(1);
            window.mapa.setView(mk.getLatLng(), 18);
            mk.openPopup();
            refrescar();

            setTimeout(()=>{ window.mapa.setZoom(13); refrescar(); }, 600);
        } else {
            mostrarModal("Pedido no encontrado");
        }
    };

    window.limpiarBusqueda = ()=>{
        document.getElementById("buscarPedido").value = "";
        window.mostrarTodos();
    };

}, 400);

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

print("Mapa ANS generado correctamente.")
webbrowser.open(str(ruta_salida))
