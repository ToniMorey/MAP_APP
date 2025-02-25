import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webview
import os
import customtkinter as cctk
import pandas as pd
import folium
import tempfile
import time
import threading
from geopy.geocoders import Nominatim
import numpy as np
from geopy.extra.rate_limiter import RateLimiter
import json
from folium import GeoJson
import re
from shapely.geometry import Point
import geopandas as gpd
from folium import plugins
from branca.colormap import linear
from folium import LayerControl

#GEOJSON poligons cat
with open("GEOData/catalunya.geojson", 'r') as file:
    CAT_DATA = json.load(file)

# Plantilla
def descargar_plantilla():
    plantilla = pd.DataFrame({
        "BvD": [], 
        "Nom empresa": [], #Need
        "Sector": [], #Need
        "Adreça": [], #1st
        "Latitud": [], #2nd
        "Longitud": [] #2nd
    })
    plantilla.to_excel('Plantilla.xlsx', index=False)
    messagebox.showinfo("Descarregar Plantilla", "La plantilla s'ha descarregat com a 'Plantilla.xlsx'")

#Variables globals que jo crec que sobren
mapa = None
temp_file_path = None
tiempo_inicio = None

#Cargar archivo:
        #Empezar ventana con temporizador al cargar el archivo
        #Cargar el mapa
        #Cerrar el temporizador cuando la ventana esté cargada
def cargar_archivo():
    global mapa, temp_file_path, tiempo_inicio

    archivo_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if archivo_path:
        #Setup finestra
        ventana_temporizador = cctk.CTkToplevel(root)
        ventana_temporizador.geometry("300x100")
        ventana_temporizador.title("Càrrega")
        root.after(100, lambda: ventana_temporizador.lift()) #lift to apareier davant de la finestra principal

        #Això d'aquí crec que no s'utilitza
        tiempo_label = cctk.StringVar(value="Carregant...")
        etiqueta_tiempo = cctk.CTkLabel(ventana_temporizador, textvariable=tiempo_label)
        etiqueta_tiempo.pack(pady=20)
        
        tiempo_inicio = time.time()  

        
        def actualizar_tiempo():
            while tiempo_inicio:  
                elapsed_time = time.time() - tiempo_inicio
                tiempo_label.set(f"Temps transcorregut: {elapsed_time:.2f} segons...")
                time.sleep(0.1)

        def cerrar_temporizador():
            tiempo_label.set("Completat!.")
            ventana_temporizador.destroy()


        #Càrrega del mapa:
            #Execució de la funció MAPmaker
            #Mapa en un temp
            #Mostrar mapa en una finestra
            #Finalitzar el temporitzador
        def progreso_carga():
            try:
                global mapa
                mapa = MAPmaker(archivo_path, icon_size, icon_color) 

                def mostrar_mapa():
                    global tiempo_inicio
                    tiempo_total = time.time() - tiempo_inicio
                    tiempo_inicio = None  
                    messagebox.showinfo(
                        "Mapa Generat",
                        f"Mapa carregat\nEmpreses carregades: {carregades}\nErrors de càrrega: {errors}\nTemps d'execució: {tiempo_total:.2f} segons."
                    )

                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
                    temp_file_path = temp_file.name
                    mapa.save(temp_file_path)
                    temp_file.close()

                    webview.create_window("Mapa", temp_file_path)
                    webview.start()
                    actualizar_boton_errores()

                    cerrar_temporizador()

                root.after(0, mostrar_mapa)

            except Exception as e: 
                tiempo_inicio = None  
                cerrar_temporizador()
                messagebox.showinfo("Error", f"Error al carregar l'arxiu: {e}")

        threading.Thread(target=progreso_carga).start()

        threading.Thread(target=actualizar_tiempo, daemon=True).start()

def cargar_archivo_zones():
    global mapa, temp_file_path, tiempo_inicio

    archivo_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if archivo_path:
        #Setup finestra
        ventana_temporizador = cctk.CTkToplevel(root)
        ventana_temporizador.geometry("300x100")
        ventana_temporizador.title("Càrrega")
        root.after(100, lambda: ventana_temporizador.lift()) #lift to apareier davant de la finestra principal

        #Això d'aquí crec que no s'utilitza
        tiempo_label = cctk.StringVar(value="Carregant...")
        etiqueta_tiempo = cctk.CTkLabel(ventana_temporizador, textvariable=tiempo_label)
        etiqueta_tiempo.pack(pady=20)
        
        tiempo_inicio = time.time()  

        
        def actualizar_tiempo():
            while tiempo_inicio:  
                elapsed_time = time.time() - tiempo_inicio
                tiempo_label.set(f"Temps transcorregut: {elapsed_time:.2f} segons...")
                time.sleep(0.1)

        def cerrar_temporizador():
            tiempo_label.set("Completat!.")
            ventana_temporizador.destroy()


        #Càrrega del mapa:
            #Execució de la funció MAPmaker
            #Mapa en un temp
            #Mostrar mapa en una finestra
            #Finalitzar el temporitzador
        def progreso_carga():
            try:
                global mapa
                mapa = MAPmaker_zones(archivo_path, tipo_select, zona_color) 

                def mostrar_mapa():
                    global tiempo_inicio
                    tiempo_total = time.time() - tiempo_inicio
                    tiempo_inicio = None  
                    messagebox.showinfo(
                        "Mapa Generat",
                        f"Mapa carregat\nEmpreses carregades: {carregades}\nErrors de càrrega: {errors}\nTemps d'execució: {tiempo_total:.2f} segons."
                    )

                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
                    temp_file_path = temp_file.name
                    mapa.save(temp_file_path)
                    temp_file.close()

                    webview.create_window("Mapa", temp_file_path)
                    webview.start()
                    actualizar_boton_errores()

                    cerrar_temporizador()

                root.after(0, mostrar_mapa)

            except Exception as e: 
                tiempo_inicio = None  
                cerrar_temporizador()
                messagebox.showinfo("Error", f"Error al carregar l'arxiu: {e}")

        threading.Thread(target=progreso_carga).start()

        threading.Thread(target=actualizar_tiempo, daemon=True).start()    

def descargar_mapa():
    if mapa is not None:
        mapa.save("mapa.html")
        messagebox.showinfo("Descarregar Mapa", "El mapa s'ha descarregat com a: 'mapa.html'")
    else:
        messagebox.showwarning("Error", "Primer s'ha de carregar un arxiu.")



#Llistes de selecció usuari
colores_iconos = [
    'red', 'blue', 'green', 'purple', 'orange', 'darkred', 'lightred', 'beige',
    'darkblue', 'darkgreen', 'cadetblue', 'darkpurple', 'white', 'pink',
    'lightblue', 'lightgreen', 'gray', 'black', 'lightgray', 'Multicolor'
]

tamanyos = [str(size) for size in [10,15,20,25,30,35,40]]

fondos = {"Clar":"CartoDB positron",
          "Carrer":"OpenStreetMap",
          "Fosc":"CartoDB dark_matter"}

colores_zonas = {"Blaus":"Blues",
                 "Blau - Verd":"BuGn",
                 "Blau - Morat":"BuPu",
                 "Verd - Blau":"GnBu",
                 "Verds":"Greens",
                 "Grisos":"Greys",
                 "Taronges":"Oranges",
                 "Taronja - Vermell":"OrRd",
                 "Morat - Blau":"PuBu",
                 "Morat - Blau - Verd": "PuBuGn",
                 "Morat - Vermell":"PuRd",
                 "Morats":"Purples",
                 "Vermell - Morat":"RdPu",
                 "Vermells":"Reds",
                 "Groc - Verd":"YlGn",
                 "Groc - Verd - Blau":"YlGnBu",
                 "Groc - Taronja - Marró":"YlOrBr",
                 "Groc - Taronja - Vermell":"YlOrRd"}

#Si les coordinates venene en format DMS
def dms_to_decimal(degrees, minutes, seconds, direction):
    decimal = degrees + minutes / 60 + seconds / 3600
    if direction in ['S', 'W']:
        decimal *= -1
    return decimal

def process_dms(coordinate):
    if pd.isna(coordinate):
        return None
    try:
        dms_parts = coordinate.replace('°', ' ').replace("'", ' ').replace('"', ' ').split()
        degrees = int(dms_parts[0])
        minutes = int(dms_parts[1])
        seconds = float(dms_parts[2])
        direction = dms_parts[3]
        return dms_to_decimal(degrees, minutes, seconds, direction)
    except (IndexError, ValueError):
        return None
    
#Si només tenim adreça
def lat(adress):
    geolocator = Nominatim(user_agent='ACCIO')
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=True)
    try:
        ubicacion = geocode(adress)
        if ubicacion:
            return ubicacion.latitude
        else:
            return None
    except Exception as e:
        return None
    
def long(adress):
    geolocator = Nominatim(user_agent='ACCIO')
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=True)
    try:
        ubicacion = geocode(adress)
        if ubicacion:
            return ubicacion.longitude
        else:
            return None
    except Exception as e:
        return None

#Logica: T
    #1st -> adreça 
        #Si adreça null o error
        #2nd -> Coordinates
        #Si coordinates en decimal OK  TEST! AQUESTA PART NO SE SI FUNCIONA
            #Si coordinates en DMS
            #Coordinates to decimal

def calc_coord(row):
    try:
        if pd.notna(row['Adreça']) and row['Adreça'].strip():
            latitud = lat(row['Adreça'])
            longitud = long(row['Adreça'])
            return pd.Series({'Latitud_decimal': latitud, 'Longitud_decimal': longitud, 'Error': False})
    except:
        pass  

    
    if 'Latitud' in row and 'Longitud' in row and pd.notna(row['Latitud']) and pd.notna(row['Longitud']):
        
        if isinstance(row['Latitud'], (int, float)) and isinstance(row['Longitud'], (int, float)):
            latitud = row['Latitud']
            longitud = row['Longitud']

        else:
            latitud = process_dms(row['Latitud'])
            longitud = process_dms(row['Longitud'])
        
        return pd.Series({'Latitud_decimal': latitud, 'Longitud_decimal': longitud, 'Error': False})
    else:
        return pd.Series({'Latitud_decimal': np.nan, 'Longitud_decimal': np.nan, 'Error': True})


    

#Funció del mapa general:
    #Arreglar archiu que es carrega:
        #Reemplazar punts per comes perque sino dona error (1)
        #Crear les columnes de coordenades (2)
        #Si hi ha algun NA (algún tipo d'error que ha sortit adalt) l'adjuntam a un df per poderlo descarregar més endevant (3)
    #ICONS -> PENDIENTE DE CANVI PER CLIENT
    #Crear ICONS amb HTML per poder canviar el tamany i el color (el client el volia diferent) CHATGPT perque de HTML ni idea
    #GEOJSON poligonal de cat per delimitar el territori
def MAPmaker(archivo, icon_size, icon_color):
    global carregades, errors, errores_df, df
    catalonia_lat = 41.391184810285964
    catalonia_long = 2.1682949088915375
    mapa = folium.Map(location=[catalonia_lat, catalonia_long], zoom_start=8.5)

    fondo = fondos[fondo_combobox.get()]
    folium.TileLayer(fondo).add_to(mapa)
    

    df = pd.read_excel(archivo)
    if 'Latitud' in df.columns and 'Longitud' in df.columns:

        df["Latitud"] = df["Latitud"].str.replace(",",".") #(1)
        df["Longitud"] = df["Longitud"].str.replace(",",".")

    coordenadas = df.apply(calc_coord, axis = 1) #(2)
    df['Latitud_decimal'] = coordenadas['Latitud_decimal']
    df['Longitud_decimal'] = coordenadas['Longitud_decimal']

    df['Errors'] = df['Latitud_decimal'].apply(lambda i: True if pd.isna(i) else False) #(3)

    errores_df = df[df['Errors']]
    
    carregades = 0
    errors = 0
    for i in df['Errors']:
        if i == False:
            carregades+=1
        elif i == True:
            errors+=1
    
    icons_ = {
    # Industria Alimentaria - VERD
    "Alimentació": ("apple-alt", "#4CAF50", "Industria Alimentaria"),
    "Vins, caves i begudes": ("wine-bottle", "#8FBC8F", "Industria Alimentaria"), 
    "Foodtech": ("utensils", "#66CDAA", "Industria Alimentaria"),

    # Tecnología e Innovación - BLAU
    "TIC i electrònica": ("microchip", "#4682B4", "Tecnología e Innovación"), 
    "AI & Big Data": ("robot", "#5A9BD3", "Tecnología e Innovación"),
    "IoT & Sensors": ("satellite-dish", "#6CA6CD", "Tecnología e Innovación"),
    "Cloud & Edge Computing": ("cloud", "#7EB6E4", "Tecnología e Innovación"),
    "Supercomputing": ("server", "#6495ED", "Tecnología e Innovación"),
    "Immersive technologies & digital entertainment": ("vr-cardboard", "#4169E1", "Tecnología e Innovación"),
    "Photonics & Quantum Sciences": ("lightbulb", "#1E90FF", "Tecnología e Innovación"),
    "Simulation & Digital Twins": ("desktop", "#5B92E5", "Tecnología e Innovación"),
    "DLT/Blockchain": ("link", "#6A5ACD", "Tecnología e Innovación"),
    "Cybersecurity": ("shield-alt", "#4682B4", "Tecnología e Innovación"),
    "New Space": ("satellite", "#5F9EA0", "Tecnología e Innovación"), 
    "R+D": ("flask", "#87CEEB", "Tecnología e Innovación"),

    # Finanzas y Servicios Empresariales - Taronja
    "Consultoria i serveis empresarials": ("briefcase", "#FF8C00", "Finanzas y Servicios Empresariales"),
    "Finances i assegurances": ("university", "#FFA07A", "Finanzas y Servicios Empresariales"),
    "Fintech & Insurtech": ("money-check-alt", "#FF7F50", "Finanzas y Servicios Empresariales"),
    "Digital Assets": ("coins", "#FF6347", "Finanzas y Servicios Empresariales"),
    "E-commerce": ("shopping-cart", "#FF4500", "Finanzas y Servicios Empresariales"),

    # Movilidad y Transporte - Vermell
    "Automoció": ("car", "#FF4500", "Movilidad y Transporte"), 
    "Altres indústries del transport": ("bus", "#DC143C", "Movilidad y Transporte"),
    "Electric Vehicle & Micromobility": ("bolt", "#FF6347", "Movilidad y Transporte"),
    "Mobilitat ferroviària": ("train", "#FF7F7F", "Movilidad y Transporte"),
    "Serveis de mobilitat": ("biking", "#FFA07A", "Movilidad y Transporte"),
    "Connected & Autonomous Vehicle": ("car", "#FF6A6A", "Movilidad y Transporte"),
    "Motocicleta i mobilitat lleugera": ("motorcycle", "#E34234", "Movilidad y Transporte"),

    # Cultura, Ocio y Deportes - Groc
    "Esports": ("running", "#FFD700", "Cultura, Ocio y Deportes"), 
    "Turisme i oci": ("umbrella-beach", "#FFFACD", "Cultura, Ocio y Deportes"), 
    "Cultura": ("theater-masks", "#F0E68C", "Cultura, Ocio y Deportes"),
    "Educació i formació i serveis editorials": ("book", "#FFDAB9", "Cultura, Ocio y Deportes"),
    "Videojocs": ("gamepad", "#FFE4B5", "Cultura, Ocio y Deportes"),

    # Sector Industrial - Gris obscur
    "Packaging": ("box-open", "#A9A9A9", "Sector Industrial"), 
    "Química i plàstics": ("flask", "#808080", "Sector Industrial"),
    "Infraestructures i construcció": ("hard-hat", "#696969", "Sector Industrial"), 
    "Transformació de metall i altres materials": ("industry", "#778899", "Sector Industrial"),
    "Frontier materials": ("cubes", "#708090", "Sector Industrial"),
    "Additive manufacturing": ("cubes", "#B0C4DE", "Sector Industrial"),
    "Matèries primeres": ("cubes", "#C0C0C0", "Sector Industrial"),
    "Maquinària i béns d’equipament": ("cogs", "#D3D3D3", "Sector Industrial"), 
    "Altres manufactures de disseny": ("paint-brush", "#DCDCDC", "Sector Industrial"),
    
    # Moda y Estilo de Vida - Rosa
    "Tèxtil i moda": ("t-shirt", "#FFC0CB", "Moda y Estilo de Vida"), 
    "Cosmètica": ("air-freshener", "#FFB6C1", "Moda y Estilo de Vida"),
    "Hàbitat": ("couch", "#FF69B4", "Moda y Estilo de Vida"),

    # Salud y Biotecnología - Marró
    "Biotecnologia": ("dna", "#8B4513", "Salud y Biotecnología"), 
    "Salut i serveis sanitaris": ("user-md", "#A0522D", "Salud y Biotecnología"), 
    "Vaccines & New biological design": ("syringe", "#D2691E", "Salud y Biotecnología"),
    "Personalized medicine": ("user-md", "#CD853F", "Salud y Biotecnología"),
    "Medical devices": ("heartbeat", "#BC8F8F", "Salud y Biotecnología"),
    "Omic Sciences": ("chart-line", "#DEB887", "Salud y Biotecnología"),
    "POCT (Point of Care Testing)": ("vials", "#F4A460", "Salud y Biotecnología"),
    "Catalysis & biocatalysis": ("flask", "#FF7F50", "Salud y Biotecnología"),
    "Indústria farmacèutica": ("pills", "#A0522D", "Salud y Biotecnología"),

    # Energía y Sostenibilidad - Verd Fosc
    "Energia": ("bolt", "#228B22", "Energía y Sostenibilidad"),
    "Clean Energy": ("solar-panel", "#32CD32", "Energía y Sostenibilidad"),
    "Decarbonization": ("wind", "#2E8B57", "Energía y Sostenibilidad"),
    "Green chemistry": ("leaf", "#3CB371", "Energía y Sostenibilidad"),
    "Hydrogen": ("burn", "#006400", "Energía y Sostenibilidad"),
    "Water cycle technologies": ("tint", "#8FBC8F", "Energía y Sostenibilidad"),
    "Smart grid & Distributed networks": ("plug", "#556B2F", "Energía y Sostenibilidad"),
    "Sustainable materials": ("seedling", "#6B8E23", "Energía y Sostenibilidad"),
    "Aigua": ("tint", "#66CDAA", "Energía y Sostenibilidad"),
    "Electricitat": ("charging-station", "#9ACD32", "Energía y Sostenibilidad"),
    
    # Agricultura y Recursos Naturales - turquesa
    "Agricultura": ("seedling", "#006400", "Agricultura y Recursos Naturales"),
    "Agritech": ("leaf", "#66CDAA", "Agricultura y Recursos Naturales"),
    "Bioeconomy": ("seedling", "#20B2AA", "Agricultura y Recursos Naturales"),
    "NBS (Nature Based Solutions)": ("tree", "#3CB371", "Agricultura y Recursos Naturales"),

    # Residuos y Reciclaje - Vermell obscur
    "Residus": ("recycle", "#2E8B57", "Residuos y Reciclaje"),
    "Recycling & Recovery": ("recycle", "#98FB98", "Residuos y Reciclaje"),
    "Urban Mining": ("recycle", "#4682B4", "Residuos y Reciclaje"),
    "Chemical recycling": ("sync-alt", "#8A2BE2"),

    # Administración Pública - Púrpura
    "Administració pública": ("landmark", "#4169E1", "Administración Pública"),
    "Smart Building": ("building", "#B22222"),
    "Smart City": ("city", "#191970"),

    # Robòtica i automatització - Blau fosc
    "Automation": ("cogs", "#FF6347"),
    "Robotics & Collaborative Robotics": ("robot", "#5F9EA0"),
    "Micro & Nano electronics": ("microchip", "#00008B"),
    "Semiconductors": ("memory", "#4682B4"),

    # Sin Tecnología Asignada - Gris
    "No technology assigned": ("question-circle", "#696969", "Sin Tecnología Asignada"),
}


    sectores = df['Sector'].unique()
    grupos_sector = {sector: folium.FeatureGroup(name=sector) for sector in sectores}


    for index, row in df.iterrows():
        lat = row['Latitud_decimal']
        long = row['Longitud_decimal']
        icon_size = int(icon_combobox.get())  
        icon_color = color_combobox.get()
        if icon_color == 'Multicolor':
            icon_color = icons_.get(row['Sector'], (None, '#D0E3F3'))[1]
        else:
            icon_color = color_combobox.get()

        if pd.notna(lat) and pd.notna(long):
            try:
                icon_sector = icons_.get(row['Sector'],('circle', None))[0]
                tooltip_html = f"Nom empresa: {row['Nom empresa']}<br>Sector: {row['Sector']}"
                
                icon_html = f"""
                <div style="background-color: {icon_color}; 
                            width: {icon_size}px; height: {icon_size}px;  
                            border-radius: 50% 50% 50% 0;
                            transform: rotate(-45deg);
                            display: flex; 
                            justify-content: center; 
                            align-items: center;
                            position: absolute; 
                            left: -{icon_size / 2}px; 
                            top: -{icon_size}px;
                            ">
                    <i class="fa fa-{icon_sector}" 
                    style="color: white; 
                            font-size: {icon_size / 2}px; 
                            transform: rotate(45deg);"></i>
                </div>
            """
                # box-shadow: -2px -4px 4px rgba(255, 6, 6, 0.5); -- Afegir sombra si cal. 
                marcador = folium.Marker(
                    location=[lat, long],
                    tooltip=tooltip_html,
                    popup=row['Nom empresa'],
                    icon=folium.DivIcon(html=icon_html)
                )
                marcador.add_to(grupos_sector[row['Sector']])

                GeoJson(
                    CAT_DATA,
                    name="cat",
                    style_function=lambda feature: {  
                        "color": "#661a2e",  
                        "weight": 2,
                        "fillColor":"white",
                        "fillOpacity":0.001  
                    }
                ). add_to(mapa)

            except Exception as e:
                print(f"Error al agregar marcador en la fila {index}: {e}")
        else:
            print(f"Coordenadas inválidas en la fila {index}: Latitud={lat}, Longitud={long}")

    for sector, grupo in grupos_sector.items():
        grupo.add_to(mapa)

    LayerControl().add_to(mapa)

    return mapa

def MAPmaker_zones(archivo, tipo_select, zona_color):
    global carregades, errors, errores_df, df
    catalonia_lat = 41.391184810285964
    catalonia_long = 2.1682949088915375
    mapa = folium.Map(location=[catalonia_lat, catalonia_long], zoom_start=8.5)

    df = pd.read_excel(archivo)
    if 'Latitud' in df.columns and 'Longitud' in df.columns:

        df["Latitud"] = df["Latitud"].str.replace(",",".") #(1)
        df["Longitud"] = df["Longitud"].str.replace(",",".")

    coordenadas = df.apply(calc_coord, axis = 1) #(2)
    df['Latitud_decimal'] = coordenadas['Latitud_decimal']
    df['Longitud_decimal'] = coordenadas['Longitud_decimal']

    df['Errors'] = df['Latitud_decimal'].apply(lambda i: True if pd.isna(i) else False) #(3)

    errores_df = df[df['Errors']]
    
    carregades = 0
    errors = 0
    for i in df['Errors']:
        if i == False:
            carregades+=1
        elif i == True:
            errors+=1

    tipo_select = princ.get()
    zona_color = colores_zonas[color_combobox_zona.get()]
    

    if tipo_select == "Provincia":
        poligons = gpd.read_file("GEOData/province_map.geojson")
    elif tipo_select == "Comarca":
        poligons = gpd.read_file("GEOData/county_map.geojson")
    else:
        poligons = gpd.read_file("GEOData/municipality_map.geojson")

    geometry = [Point(xy) for xy in zip(df['Longitud_decimal'], df['Latitud_decimal'])]
    points_gdf = gpd.GeoDataFrame(df, geometry=geometry)

   
    points_gdf.set_crs("EPSG:4326", inplace=True)  
    poligons.to_crs("EPSG:4326", inplace=True)

    
    points_in_polygons = gpd.sjoin(points_gdf, poligons, how="left", predicate="within")


    if tipo_select == "Provincia":
    
        dot_counts = points_in_polygons['nomprov'].value_counts().reset_index()
        dot_counts.columns = ['nomprov', 'point_count']

        
        all_dots = pd.DataFrame(poligons['nomprov'])
        all_dots['point_count'] = all_dots['nomprov'].map(dot_counts.set_index('nomprov')['point_count']).fillna(0)
        all_dots['point_count_falso'] = np.log1p(all_dots['point_count'])*10
        llave = 'feature.properties.nomprov'

    elif tipo_select == "Comarca":
    
        dot_counts = points_in_polygons['nomcomar'].value_counts().reset_index()
        dot_counts.columns = ['nomcomar', 'point_count']

        
        all_dots = pd.DataFrame(poligons['nomcomar'])
        all_dots['point_count'] = all_dots['nomcomar'].map(dot_counts.set_index('nomcomar')['point_count']).fillna(0)
        all_dots['point_count_falso'] = all_dots['point_count'].clip(upper=300)
        #all_dots['point_count_falso'] = np.log1p(all_dots['point_count_falso'])*10
        llave = 'feature.properties.nomcomar'
        
    else: 
    
        dot_counts = points_in_polygons['nom_muni'].value_counts().reset_index()
        dot_counts.columns = ['nom_muni', 'point_count']

        
        all_dots = pd.DataFrame(poligons['nom_muni'])
        all_dots['point_count'] = all_dots['nom_muni'].map(dot_counts.set_index('nom_muni')['point_count']).fillna(0)
        all_dots['point_count_falso'] = np.log1p(all_dots['point_count'])*10
        llave = 'feature.properties.nom_muni'

    
    all_dots

    colormap = getattr(linear, f"{zona_color}_09").scale(0, all_dots['point_count_falso'].max()) 

    folium.TileLayer("CartoDB positron").add_to(mapa)

    GeoJson(
                    CAT_DATA,
                    name="cat",
                    style_function=lambda feature: {  
                        "color": "#661a2e",  
                        "weight": 2,
                        "fillColor":"white",
                        "fillOpacity":0.001  
                    }
                ). add_to(mapa)

    if tipo_select == "Provincia": 
        for _, row in all_dots.iterrows():
            if row['point_count'] > 0: 
                municipio = poligons[poligons['nomprov'] == row['nomprov']]
                folium.GeoJson(
                    municipio,
                    tooltip=f"Provincia: {row['nomprov']} <br> Nombre empreses: {int(row['point_count'])}",
                    style_function=lambda x, count=row['point_count']: {
                        'fillColor': colormap(count), 
                        'color': 'black', 
                        'weight': 0.5,  
                        'fillOpacity': 1
                    }
                ).add_to(mapa)

    elif tipo_select == "Comarca": 
        for _, row in all_dots.iterrows():
            if row['point_count'] > 0:  
                municipio = poligons[poligons['nomcomar'] == row['nomcomar']]
                folium.GeoJson(
                    municipio,
                    tooltip=f"Comarca: {row['nomcomar']} <br> Nombre empreses: {int(row['point_count'])}",
                    style_function=lambda x, count=row['point_count']: {
                        'fillColor': colormap(count),  
                        'color': 'black', 
                        'weight': 0.5,  
                        'fillOpacity': 1
                    }
                ).add_to(mapa)


    else: 
        for _, row in all_dots.iterrows():
            if row['point_count'] > 0: 
                municipio = poligons[poligons['nom_muni'] == row['nom_muni']]
                folium.GeoJson(
                    municipio,
                    tooltip=f"Municipi: {row['nom_muni']} <br> Nombre empreses: {int(row['point_count'])}",
                    style_function=lambda x, count=row['point_count']: {
                        'fillColor': colormap(count),  
                        'color': 'black', 
                        'weight': 0.5, 
                        'fillOpacity': 1
                    }
                ).add_to(mapa)
    
    return mapa
    

def actualizar_boton_errores():
    global boton_descargar_errores
    if errors > 0:
        boton_descargar_errores.grid(row=6, column=1, padx=20, pady=10)
        boton_descargar_errores.configure(fg_color="darkred") 
        boton_descargar_errores_zona.pack(pady = 5)
        boton_descargar_errores_zona.configure(fg_color="darkred") 
    else:
        boton_descargar_errores.pack_forget()
        boton_descargar_errores_zona.pack_forget()


def descargar_errores_df():
    if errores_df is not None:
        errores_df.drop(columns=['Latitud_decimal', 'Longitud_decimal', 'Errors'], inplace = True)
        errores_df.to_excel('Errors.xlsx')
        messagebox.showinfo("Errors", "S'ha descarregat el fitxer de errors com a 'Errors.xlsx'")
    else:
        messagebox.showwarning("Error", "No hi ha cap error de càrrega.")

#TKINTER MAIN WINDOW (CUSTOMTKINTER)
root = cctk.CTk()
cctk.set_appearance_mode("light")
cctk.set_default_color_theme("blue")

root.title("MAP-Maker")
root.geometry("700x500")

root.iconbitmap("icons/10168128.ico")

tabview = cctk.CTkTabview(root)
tabview.pack(fill="both", expand=True)

tabview.add("Mapes de punts")
tab_markers = tabview.tab("Mapes de punts")

tabview.add("Mapes de zones")
tab_zona = tabview.tab("Mapes de zones")
        #MARKERS
#Plantilla
boton_plantilla = cctk.CTkButton(tab_markers, text="Descarrega plantilla", command=descargar_plantilla)
boton_plantilla.grid(row=0, column=1, padx=10, pady=10)

#Carregar arxiu
boton_cargar_archivo = cctk.CTkButton(tab_markers, text="Carrega l'arxiu i previsualitza el mapa", command=cargar_archivo)
boton_cargar_archivo.grid(row=1, column=1, padx=10, pady=10)

#Icones tamany

label_icon = cctk.CTkLabel(tab_markers, text="Selecciona el tamany de la icona:")
label_icon.grid(row=2, column=0, padx=5, pady=5)
icon_combobox = cctk.CTkComboBox(tab_markers, values=tamanyos, state="readonly")
icon_combobox.grid(row=3, column=0, padx=5, pady=5)
icon_combobox.set(tamanyos[2])


#Icones colors
label_color = cctk.CTkLabel(tab_markers, text="Selecciona el color de la icona:")
label_color.grid(row=2, column=1, padx=5, pady=5)
color_combobox = cctk.CTkComboBox(tab_markers, values=colores_iconos, state="readonly")
color_combobox.grid(row=3, column=1, padx=5, pady=5)
color_combobox.set(colores_iconos[0])

#Fons mapa
fondo_color = cctk.CTkLabel(tab_markers, text="Selecciona el fons:")
fondo_color.grid(row=2, column=2, padx=5, pady=5)
fondo_combobox = cctk.CTkComboBox(tab_markers, values=list(fondos.keys()), state="readonly")
fondo_combobox.grid(row=3, column=2, padx=5, pady=5)
fondo_combobox.set(list(fondos.keys())[0])


#Descarregar mapa
boton_descargar_mapa = cctk.CTkButton(tab_markers, text="Descarrega el mapa", command=descargar_mapa)
boton_descargar_mapa.grid(row=5, column=1, padx=20, pady=20)

#Errors
boton_descargar_errores = cctk.CTkButton(
    tab_markers, 
    text="Descarrega els errors", 
    command=descargar_errores_df,
    fg_color="darkred"  # Color inicial
)


        #ZONES
boton_plantilla = cctk.CTkButton(tab_zona, text="Descarrega plantilla", command=descargar_plantilla)
boton_plantilla.pack(pady=5)#plantilla

boton_cargar_archivo = cctk.CTkButton(tab_zona, text="Carrega l'arxiu i previsualitza el mapa", command=cargar_archivo_zones)
boton_cargar_archivo.pack(pady=15) #Carregar archiu

#seleccionar si com - prov - muni

tipo_icon = cctk.CTkLabel(tab_zona, text="Selecciona tipus de mapa:")
tipo_icon.pack()
princ = cctk.StringVar(value="Muinicipis")
munis_selec = cctk.CTkRadioButton(tab_zona, text ="Municipi", variable=princ, value= "Municipi")
munis_selec.pack(pady=5)

provin_selec= cctk.CTkRadioButton(tab_zona, text ="Provincia", variable=princ, value= "Provincia")
provin_selec.pack(pady=5)

comarc_selec = cctk.CTkRadioButton(tab_zona, text ="Comarca", variable=princ, value= "Comarca")
comarc_selec.pack(pady=5)


#colors
label_color_zona = cctk.CTkLabel(tab_zona, text="Selecciona el color de la icona:")
label_color_zona.pack()
color_combobox_zona = cctk.CTkComboBox(tab_zona, values=list(colores_zonas.keys()), state="readonly")
color_combobox_zona.pack(pady=5)
color_combobox_zona.set(list(colores_zonas.keys())[13])

boton_descargar_mapa_zones = cctk.CTkButton(tab_zona, text="Descarrega el mapa", command=descargar_mapa)
boton_descargar_mapa_zones.pack(pady=5)

boton_descargar_errores_zona = cctk.CTkButton(
    tab_zona, 
    text="Descarrega els errors", 
    command=descargar_errores_df,
    fg_color="darkred"  # Color inicial
)

#Això crec que també sobra pero no ho vull treure
icon_size = int(icon_combobox.get())  
icon_color = color_combobox.get()
fondo = fondos[fondo_combobox.get()]
tipo_select = princ.get()
zona_color = colores_zonas[color_combobox_zona.get()]

root.mainloop()
