import sys
import os

BASE = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(BASE, 'MainApp_PlaneacionHTML.html')

# Lee el archivo completo
with open(HTML_PATH, 'r', encoding='utf-8') as f:
    lineas = f.readlines()

print(f"Total líneas leídas: {len(lineas)}")

# Módulos: (nombre_archivo, inicio_1based, fin_1based)
modulos = [
    ("Mod_Dashboard.html",          1351, 1701),
    ("Mod_Metricas.html",           1702, 2924),
    ("Mod_Validador.html",          2925, 3844),
    ("Mod_Tracking.html",           3845, 3943),
    ("Mod_NuevoPedido.html",        3944, 4777),
    ("Mod_PedidoINT.html",          4778, 4982),
    ("Mod_Generador.html",          4983, 5759),
    ("Mod_Planificador.html",       5760, 13507),
    ("Mod_Programador.html",        14431, 15322),
    ("Mod_ProgInteractivo.html",    15323, 16901),
    ("Mod_Wip.html",                16902, 17230),
    ("Mod_Captura.html",            17231, 18765),
    ("Mod_EditorProd.html",         18766, 19643),
    ("Mod_NuevaRuta.html",          19644, 20107),
    ("Mod_EditorRutas.html",        20108, 20604),
    ("Mod_TableroSupervisor.html",  20605, 21949),
    ("Mod_GestionDocs.html",        21950, 22143),
    ("Mod_Usuarios.html",           22144, 22536),
    ("Mod_TrackingModule.html",     22537, 23310),
    ("Mod_MrpMp.html",              23311, 24268),
    ("Mod_SalidaAlambres.html",     24269, 24478),
    ("Mod_MisRegistros.html",       24479, 25044),
    ("Mod_PedidosPendientes.html",  25047, 25302),
]

# PASO A: Extraer cada módulo a su propio archivo
for nombre, inicio, fin in modulos:
    # Convertir a índices base-0
    i0 = inicio - 1
    i1 = fin        # slice [i0:i1] = líneas inicio..fin inclusive
    bloque = lineas[i0:i1]
    ruta = os.path.join(BASE, nombre)
    with open(ruta, 'w', encoding='utf-8') as f:
        f.writelines(bloque)
    print(f"  Creado {nombre}: {len(bloque)} líneas ({inicio}–{fin})")

# PASO B: Reemplazar los bloques en el HTML principal con etiquetas include
# Reemplazos en índices base-0: [inicio, fin) → una sola línea con el tag
# Ordenar de mayor a menor para no desplazar índices
reemplazos = [
    (1350, 1701, "<?!= include('Mod_Dashboard'); ?>\n"),
    (1701, 2924, "<?!= include('Mod_Metricas'); ?>\n"),
    (2924, 3844, "<?!= include('Mod_Validador'); ?>\n"),
    (3844, 3943, "<?!= include('Mod_Tracking'); ?>\n"),
    (3943, 4777, "<?!= include('Mod_NuevoPedido'); ?>\n"),
    (4777, 4982, "<?!= include('Mod_PedidoINT'); ?>\n"),
    (4982, 5759, "<?!= include('Mod_Generador'); ?>\n"),
    (5759, 13507, "<?!= include('Mod_Planificador'); ?>\n"),
    (13507, 15322, "<?!= include('Mod_Programador'); ?>\n"),
    (15322, 16901, "<?!= include('Mod_ProgInteractivo'); ?>\n"),
    (16901, 17230, "<?!= include('Mod_Wip'); ?>\n"),
    (17230, 18765, "<?!= include('Mod_Captura'); ?>\n"),
    (18765, 19643, "<?!= include('Mod_EditorProd'); ?>\n"),
    (19643, 20107, "<?!= include('Mod_NuevaRuta'); ?>\n"),
    (20107, 20604, "<?!= include('Mod_EditorRutas'); ?>\n"),
    (20604, 21949, "<?!= include('Mod_TableroSupervisor'); ?>\n"),
    (21949, 22143, "<?!= include('Mod_GestionDocs'); ?>\n"),
    (22143, 22536, "<?!= include('Mod_Usuarios'); ?>\n"),
    (22536, 23310, "<?!= include('Mod_TrackingModule'); ?>\n"),
    (23310, 24268, "<?!= include('Mod_MrpMp'); ?>\n"),
    (24268, 24478, "<?!= include('Mod_SalidaAlambres'); ?>\n"),
    (24478, 25044, "<?!= include('Mod_MisRegistros'); ?>\n"),
    (25046, 25302, "<?!= include('Mod_PedidosPendientes'); ?>\n"),
]

for inicio, fin, tag in sorted(reemplazos, key=lambda x: -x[0]):
    lineas[inicio:fin] = [tag]

with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.writelines(lineas)

print(f"\nMainApp_PlaneacionHTML.html actualizado.")
print(f"Líneas resultado: {len(lineas)}")
