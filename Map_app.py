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

#GEOJSON poligons cat
with open("catalunya.geojson", 'r') as file:
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
    'lightblue', 'lightgreen', 'gray', 'black', 'lightgray'
]

tamanyos = [str(size) for size in [10,15,20,25,30,35,40]]

fondos = {"Clar":"CartoDB positron",
          "Carrer":"OpenStreetMap",
          "Fosc":"CartoDB dark_matter"}

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
        #Si coordinates en decimal i formato text!!! OK
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
        try:
            if re.match(r'^\d+\.\d+$', str(row['Latitud'])) and re.match(r'^\d+\.\d+$', str(row['Longitud'])):
                latitud = float(row['Latitud'])
                longitud = float(row['Longitud'])
            else:
                latitud = process_dms(row['Latitud'])
                longitud = process_dms(row['Longitud'])
            
            return pd.Series({'Latitud_decimal': latitud, 'Longitud_decimal': longitud, 'Error': False})
        except Exception as e:
            return pd.Series({'Latitud_decimal': np.nan, 'Longitud_decimal': np.nan, 'Error': True})
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
        "Alimentació": "apple-alt", "TIC i electrònica": "microchip", "Consultoria i serveis empresarials": "briefcase",
        "Automoció": "car", "Esports": "running", "Packaging": "box-open", "Química i plàstics": "flask",
        "Infraestructures i construcció": "hard-hat", "Altres indústries del transport": "truck",
        "Turisme i oci": "umbrella-beach", "Tèxtil i moda": "t-shirt", "Maquinària i béns d’equipament": "cogs",
        "Logística de mercaderies": "truck-loading", "Hàbitat": "couch", "Energia": "bolt",
        "Finances i assegurances": "university", "Salut animal": "paw", "Videojocs": "gamepad",
        "Biotecnologia": "dna", "Salut i serveis sanitaris": "user-md", "Cosmètica": "air-freshener",
        "Vins, caves i begudes": "wine-bottle", "Indústria farmacèutica": "pills",
        "Matèries primeres": "cubes", "Transformació de metall i altres materials": "industry",
        "Electricitat": "charging-station", "Audiovisual": "video", "Educació i formació i serveis editorials": "book",
        "Administració pública": "landmark", "Agricultura": "seedling", "Aigua": "tint",
        "Altres manufactures de disseny": "paint-brush", "Altres serveis": "servicestack", "Cultura": "theater-masks",
        "Mobilitat ferroviària": "train", "Motocicleta i mobilitat lleugera": "motorcycle",
        "R+D": "flask", "Residus": "recycle", "Serveis de mobilitat": "biking", "Additive manufacturing": "cubes",
        "Agritech": "leaf", "AI & Big Data": "robot","Automation": "cogs","Batteries & Storage": "battery-full","Bioeconomy": "seedling","Biotechnology": "vial","Blue Economy": "water",
        "Catalysis & biocatalysis": "flask","Chemical recycling": "sync-alt","Clean Energy": "solar-panel","Cloud & Edge Computing": "cloud","Connected & Autonomous Vehicle": "car",
        "Connectivity": "wifi","Cybersecurity": "shield-alt","Decarbonization": "wind","Digital Assets": "coins","Digital Health": "heartbeat","DLT/Blockchain": "link",
        "Drones": "helicopter","E-commerce": "shopping-cart","Electric Vehicle & Micromobility": "bolt","Energy Harvesting": "solar-panel","Fintech & Insurtech": "money-check-alt",
        "Foodtech": "utensils","Frontier materials": "cubes","Green chemistry": "leaf","Hydrogen": "burn","Immersive technologies & digital entertainment": "vr-cardboard","IoT & Sensors": "satellite-dish",
        "Medical devices": "heartbeat","Micro & Nano electronics": "microchip","NBS (Nature Based Solutions)": "tree","New Space": "satellite","No technology assigned": "question-circle",
        "Omic Sciences": "chart-line","Personalized medicine": "user-md","Photonics & Quantum Sciences": "lightbulb","POCT (Point of Care Testing)": "vials","Recycling & Recovery": "recycle",
        "Robotics & Collaborative Robotics": "robot","Semiconductors": "memory","Simulation & Digital Twins": "desktop","Smart Building": "building","Smart City": "city",
        "Smart grid & Distributed networks": "plug","Supercomputing": "server","Sustainable materials": "seedling","Urban Mining": "recycle","Vaccines & New biological design": "syringe",
        "Water cycle technologies": "tint"}

    for index, row in df.iterrows():
        lat = row['Latitud_decimal']
        long = row['Longitud_decimal']
        icon_size = int(icon_combobox.get())  
        icon_color = color_combobox.get()

        if pd.notna(lat) and pd.notna(long):
            try:
                icon_sector = icons_.get(row['Sector'], 'circle')
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
                                top: -{icon_size}px;">
                        <i class="fa fa-{icon_sector}" 
                        style="color: white; 
                                font-size: {icon_size / 2}px; 
                                transform: rotate(45deg);"></i>
                    </div>
                """
                folium.Marker(
                    location=[lat, long],
                    tooltip=tooltip_html,
                    popup=row['Nom empresa'],
                    icon=folium.DivIcon(html=icon_html)
                ).add_to(mapa)

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


    return mapa

def actualizar_boton_errores():
    global boton_descargar_errores
    if errors > 0:
        boton_descargar_errores.pack(pady=5)  
        boton_descargar_errores.configure(fg_color="darkred") 
    else:
        boton_descargar_errores.pack_forget()


def descargar_errores_df():
    if errores_df is not None:
        errores_df.to_excel('Errors.xlsx')
        messagebox.showinfo("Errors", "S'ha descarregat el fitxer de errors com a 'Errors.xlsx'")
    else:
        messagebox.showwarning("Error", "No hi ha cap error de càrrega.")

#TKINTER MAIN WINDOW (CUSTOMTKINTER)
root = cctk.CTk()
cctk.set_appearance_mode("light")
cctk.set_default_color_theme("blue")

root.title("MAP-Maker")
root.geometry("600x400")

root.iconbitmap("icons/10168128.ico")

#Plantilla
boton_plantilla = cctk.CTkButton(root, text="Descarrega plantilla", command=descargar_plantilla)
boton_plantilla.pack(pady=5)

#Carregar arxiu
boton_cargar_archivo = cctk.CTkButton(root, text="Carrega l'arxiu i previsualitza el mapa", command=cargar_archivo)
boton_cargar_archivo.pack(pady=5)

#Icones tamany
label_icon = tk.Label(root, text="Selecciona el tamany de la icona:")
label_icon.pack()
icon_combobox = cctk.CTkComboBox(root, values=tamanyos, state="readonly")
icon_combobox.pack(pady=5)
icon_combobox.set(tamanyos[2])


#Icones colors
label_color = tk.Label(root, text="Selecciona el color de la icona:")
label_color.pack()
color_combobox = cctk.CTkComboBox(root, values=colores_iconos, state="readonly")
color_combobox.pack(pady=5)
color_combobox.set(colores_iconos[0])

#Fons mapa
fondo_color = tk.Label(root, text="Selecciona el fons:")
fondo_color.pack()
fondo_combobox = cctk.CTkComboBox(root, values=list(fondos.keys()), state="readonly")
fondo_combobox.pack(pady=5)
fondo_combobox.set(list(fondos.keys())[0])


#Descarregar mapa
boton_descargar_mapa = cctk.CTkButton(root, text="Descarrega el mapa", command=descargar_mapa)
boton_descargar_mapa.pack(pady=5)

#Errors
boton_descargar_errores = cctk.CTkButton(
    root, 
    text="Descarrega els errors", 
    command=descargar_errores_df,
    fg_color="darkred"  # Color inicial
)

#Això crec que també sobra pero no ho vull treure
icon_size = int(icon_combobox.get())  
icon_color = color_combobox.get()
fondo = fondos[fondo_combobox.get()]

root.mainloop()
