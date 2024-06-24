## Librerías
import os
import pandas as pd
import openpyxl
from openpyxl import Workbook
import networkx as nx
import matplotlib.pyplot as plt
import itertools
import random
from IPython.display import display, HTML
from tabulate import tabulate
import warnings

# Ignorar las advertencias de tipo DeprecationWarning
warnings.filterwarnings("ignore", category=DeprecationWarning)

## Gestionar datos

### Clase Persona

class Persona():
    ruta_archivo = "personas.xlsx"  # Ruta al archivo de datos

    def __init__(self, nombre, deporte=None, musica=None, danza=None, instrumento=None, club=None, hobbie=None):
        # Asigna el nombre y los gustos opcionales a una persona
        self.nombre = nombre
        self.deporte = deporte
        self.musica = musica
        self.danza = danza
        self.instrumento = instrumento
        self.club = club
        self.hobbie = hobbie

    def __str__(self):
        # Devuelve el nombre de la persona en formato de texto
        return f"{self.nombre}"

    def get_gustos_str(self):
        # Devuelve los gustos de la persona como una cadena de texto
        return f"{self.deporte}, {self.musica}, {self.danza}, {self.instrumento}, {self.club}, {self.hobbie}"

    def tiene_afinidad(self, persona2):
        # Comprueba si la persona actual comparte al menos dos gustos con otra persona
        cont = 0
        for gusto in self.__dict__.keys():
            if gusto == 'nombre':
                continue
            if (getattr(self, gusto) == getattr(persona2, gusto)):
                cont += 1
            if cont == 2:
                return True
        return False

    def guardar(self):
        # Guarda los datos de la persona en el Excel
        datos_persona = list(self.__dict__.values())

        if not os.path.isfile(self.ruta_archivo):
            workbook = Workbook()  # Crea un nuevo archivo de Excel si no hubiera uno
            hoja = workbook.active
            hoja.append(list(self.__dict__.keys()))  # Agrega los nombres de los atributos como la primera fila
        else:
            workbook = openpyxl.load_workbook(self.ruta_archivo)  # Abre el archivo existente
            hoja = workbook.active

        hoja.append(datos_persona)  # Agrega los datos de la persona como una nueva fila
        workbook.save(self.ruta_archivo)  # Guarda los cambios en el archivo
        print(f"Se ha guardado a {self.nombre} en el archivo {self.ruta_archivo}")


### Generar y cargar datos de forma aleatoria
# El siguiente código genera y carga datos de forma aleatoria utilizando una lista de nombres predefinida:

nombres = [
    "Juan", "María", "Pedro", "Ana", "Luis", "Laura", "Carlos", "Sofía", "Diego", "Marta",
    "José", "Lucía", "Andrés", "Paula", "Miguel", "Valentina", "Javier", "Camila", "Pablo", "Elena",
    "Fernando", "Isabella", "Ricardo", "Julia", "Gabriel", "Victoria", "Alejandro", "Gabriela", "Santiago", "Daniela",
    "Raúl", "Natalia", "Roberto", "Olivia", "Esteban", "Jimena", "Francisco", "Antonella", "Daniel", "Carolina",
    "Gonzalo", "Ángela", "Arturo", "Clara", "Emilio", "Adriana", "Hugo", "Florencia", "Enrique", "Beatriz",
    "Lorenzo", "Constanza", "Renato", "Daniella", "Rodrigo", "Paulina", "Sebastián", "Agustina", "Josué", "Martina",
    "Oscar", "Catalina", "Adrián", "Rosario", "Simón", "Juliana", "Mateo", "Renata", "Rafael", "Isabel",
    "Víctor", "Carmen", "Diego", "Mariana", "Emmanuel", "Abril", "Ignacio", "Emily", "Álvaro", "Melissa",
    "Ramón", "Valeria", "César", "Amanda", "Eduardo", "Luciana", "Nicolás", "Danielle", "Alberto", "Alejandra",
    "Tomás", "Brenda", "Luis", "Gloria", "Ulises", "Jennifer", "Jorge", "Celeste", "Raúl", "Patricia"
]

def generar_datos_aleatorios(n, nombres_list):
    deportes = ["futbol", "basquet", "voley", "natacion", "karate", "otros"]
    musicas = ["salsa", "rock", "bachata", "regaetton", "merengue", "technocumbia", "folklorica", "otros"]
    danzas = ["salsa", "rock", "bachata", "regaetton", "merengue", "technocumbia", "folklorica", "otros"]
    instrumentos = ["guitarra", "bateria", "piano", "saxo", "no toca", "otros"]
    clubes = ["x", "y", "z", "d", "p"]
    hobbies = ["cine", "visitar museos", "viajar", "oratoria", "videojuegos", "conciertos", "otros"]

    personas = []
    for i in range(1, n+1):
        nombre = random.choice(nombres_list)
        deporte = random.choice(deportes)
        musica = random.choice(musicas)
        danza = random.choice(danzas)
        instrumento = random.choice(instrumentos)
        club = random.choice(clubes)
        hobbie = random.choice(hobbies)

        persona = Persona(nombre, deporte, musica, danza, instrumento, club, hobbie)
        persona.guardar()  # Guarda cada persona en el archivo excel
        personas.append(persona)

### Leer datos del excel

def get_personas_desde_excel(ruta_archivo):
    if not os.path.exists(ruta_archivo):  # si el archivo no existe, lo creamos
        df = pd.DataFrame(columns=['nombre', 'deporte', 'musica', 'danza', 'instrumento', 'club', 'hobbie'])
        df.to_excel(ruta_archivo, index=False)
    df = pd.read_excel(ruta_archivo)  # ahora seguro el archivo existe
    personas = []
    
    for _, row in df.iterrows():
        persona = Persona(row['nombre'], row['deporte'], row['musica'], row['danza'], row['instrumento'], row['club'], row['hobbie'])
        personas.append(persona)
        
    return personas

### Borrar datos del excel

def borrar_datos_excel(ruta_archivo):
    if os.path.exists(ruta_archivo):
        df = pd.DataFrame(columns=['nombre', 'deporte', 'musica', 'danza', 'instrumento', 'club', 'hobbie'])
        df.to_excel(ruta_archivo, index=False)
    else:
        print(f'El archivo {ruta_archivo} no existe.')

## Crear Grafo de personas

### Crear grafo

def crear_grafo(personas):
    G = nx.Graph()  # Crea un objeto de grafo vacío
    G.add_nodes_from(personas)  # Agrega los nodos al grafo

    # Revisa todas las combinaciones posibles de personas en el grafo
    for persona1, persona2 in itertools.combinations(personas, 2):
        if persona1.tiene_afinidad(persona2):  # Verifica si hay afinidad entre las personas
            G.add_edge(persona1, persona2)  # Agrega una arista entre dos personas si hay afinidad

    return G

### Graficar grafo

def graficar_grafo_no_dirigido_simple(G, nodos_especiales=[]):
    ax = plt.gca()
    pos = nx.spring_layout(G, k=2)  # Obtiene la posición de los nodos utilizando el algoritmo de spring layout
    opciones = {
        'pos': pos,
        'edge_color': '#808080',  # Color de las aristas
        'arrows': True,  # Flechas en las aristas
        'ax': ax,
        'connectionstyle': "arc3,rad=0",
        'edgelist': G.edges(),
    }

    nx.draw_networkx_edges(G, **opciones)  # Dibuja las aristas del grafo
    nx.draw_networkx_labels(G, pos, font_size=8,
                            font_color='black', font_weight='bold')  # Dibuja las etiquetas de los nodos

    # Separa los nodos en nodos especiales y nodos regulares, y los dibuja con diferentes colores
    nodos_especiales = set(nodos_especiales)
    nodos_regulares = [nodo for nodo in G.nodes if nodo not in nodos_especiales]
    nx.draw_networkx_nodes(G, pos, nodelist=nodos_especiales, node_color='red', node_size=900, alpha=1)  # Nodos especiales en rojo
    nx.draw_networkx_nodes(G, pos, nodelist=nodos_regulares, node_color='#0080FF', node_size=900, alpha=1)  # Nodos regulares con color original
    
    # Obtiene las aristas que conectan dos nodos especiales y las aristas regulares
    aristas_especiales = [(u, v) for (u, v) in G.edges if u in nodos_especiales and v in nodos_especiales]
    aristas_regulares = [(u, v) for (u, v) in G.edges if (u, v) not in aristas_especiales and (v, u) not in aristas_especiales]
    
    # Dibuja las aristas especiales con un color diferente
    nx.draw_networkx_edges(G, pos, edgelist=aristas_especiales, edge_color='red', arrows=True, ax=ax)
    nx.draw_networkx_edges(G, pos, edgelist=aristas_regulares, edge_color='#808080', arrows=True, ax=ax)

    etiquetas_aristas = nx.get_edge_attributes(G, "weight")
    nx.draw_networkx_edge_labels(G, pos, etiquetas_aristas)  # Dibuja las etiquetas de las aristas

    plt.axis('off')
    plt.show()  # Muestra el gráfico

## Encontrar equipos

### DFS (Depth First Search) Limitado

def dfs_limitado(grafico, inicio, max_nodos):
    pila = [inicio]  # Inicializa una pila con el nodo inicial
    visitados = []  # Lista para almacenar los nodos visitados

    while pila and len(visitados) < max_nodos:  # Mientras haya nodos en la pila y no se haya alcanzado el límite máximo de nodos visitados
        nodo = pila.pop()  # Extrae un nodo de la pila
        if nodo not in visitados:  # Si el nodo no ha sido visitado
            visitados.append(nodo)  # Agrega el nodo a la lista de visitados
            pila.extend(n for n in grafico.neighbors(nodo) if n not in visitados)  # Agrega a la pila los vecinos no visitados del nodo actual

    return visitados  # Cuando termina devuelve la lista de nodos visitados

### Encontrar subgrafos conexos de minimos 5 nodos y maximo 7 nodos

def encontrar_subgrafos_conexos(graph, min_nodes=5, max_nodes=7, print_process=True):
    subgrafos = []
    graph_copy = graph.copy()

    while len(graph_copy.nodes) >= min_nodes:

        # Seleccionar un nodo aleatorio para empezar
        start_node = random.choice(list(graph_copy.nodes))

        # Realizar un DFS para encontrar un subgrafo conexo
        subgraph_nodes = dfs_limitado(graph_copy, start_node, max_nodes)

        # Si el subgrafo tiene suficientes nodos, añadirlo a la lista de subgrafos
        if len(subgraph_nodes) >= min_nodes:
            subgrafos.append(graph.subgraph(subgraph_nodes))
            
            if (print_process):
                # Imprimir grafo con los nodos removidos en rojo
                print('Nodo Inicial:', start_node)
                # print('Nodos retirados:', [str(node) for node in subgraph_nodes])
                graficar_grafo_no_dirigido_simple(graph_copy, subgraph_nodes)

        # Quitar los nodos del subgrafo de la copia del grafo original
        graph_copy.remove_nodes_from(subgraph_nodes)

    # plot_graph(graph_copy)
    return subgrafos

### Imprimir equipos

def imprimir_equipos(subgrafos, imprimir_tabla=True, imprimir_grafo=True):
    print(f'Se encontraron {len(subgrafos)} equipos:')  # Imprime la cantidad de equipos encontrados
    for i, subgrafo in enumerate(subgrafos):
        if imprimir_tabla:
            print(f'Equipo {i+1}:')  # Imprime el número de equipo
            tabla = []
            for persona in subgrafo.nodes:
                tabla.append([persona.nombre, persona.deporte, persona.musica, persona.danza, persona.instrumento, persona.club, persona.hobbie])

            headers = ["Nombre", "Deporte", "Música", "Danza", "Instrumento", "Club", "Hobbie"]
            df = pd.DataFrame(tabla, columns=headers)
            display(HTML(df.to_html(index=False, classes='table-bordered table-striped')))  # Imprime una tabla con los datos de las personas en el equipo
        else:
            print(f'Equipo {i+1}: {[str(persona) for persona in subgrafo.nodes]}')  # Imprime los nombres de las personas en el equipo

        if imprimir_grafo:
            graficar_grafo_no_dirigido_simple(subgrafo)  # Grafica el subgrafo del equipo


def agregar_persona():
    print("Por favor ingresa los datos de la persona:")
    nombre = input("Nombre: ")
    deporte = input("Deporte favorito: ")
    musica = input("Música favorita: ")
    danza = input("Danza de preferencia: ")
    instrumento = input("Instrumento que toca: ")
    club = input("Club favorito: ")
    hobbie = input("Hobbie favorito: ")

    persona = Persona(nombre, deporte, musica, danza, instrumento, club, hobbie)
    persona.guardar()

    print("La persona ha sido agregada con éxito.")


def limpiar_consola():
    if os.name == 'nt':  # Para Windows
        os.system('cls')
    else:  # Para Unix/Linux
        os.system('clear')

def menu_principal():
    print("\n*********** Generador de equipos de trabajo ***********")
    print("1. Generar datos aleatorios")
    print("2. Agregar una persona")
    print("3. Ver datos actuales")
    print("4. Borrar todos los datos")
    print("5. Crear equipos")
    print("0. Salir")

while True:
    limpiar_consola()
    menu_principal()
    opcion = input("Seleccione una opción del menú: ")

    if opcion == '1':
        limpiar_consola()
        n = int(input("Ingrese el número de personas a generar: "))
        generar_datos_aleatorios(n, nombres)
        print(f"\n{n} personas han sido generadas y guardadas exitosamente.")
        input()
    elif opcion == '2':
        limpiar_consola()
        agregar_persona()
        input()
    elif opcion == '3':
        limpiar_consola()
        personas = get_personas_desde_excel(Persona.ruta_archivo)
        print("Datos actuales:")
        tabla = [["Nombre", "Deporte", "Musica", "Danza", "Instrumento", "Club", "Hobbie"]]
        for persona in personas:
            gustos = persona.get_gustos_str().split(", ")
            fila = [persona.nombre] + gustos
            tabla.append(fila)
        print(tabulate(tabla, headers="firstrow", tablefmt='fancy_grid'))
        input()
    elif opcion == '4':
        limpiar_consola()
        confirmacion = input("¿Estás seguro de que deseas borrar todos los datos? (S/N) ")
        if confirmacion.lower() == 's':
            borrar_datos_excel(Persona.ruta_archivo)
            print("Todos los datos han sido borrados exitosamente.")
        input()
    elif opcion == '5':
        limpiar_consola()
        personas = get_personas_desde_excel(Persona.ruta_archivo)
        G = crear_grafo(personas)
        subgrafos = encontrar_subgrafos_conexos(G, min_nodes=5, max_nodes=7, print_process=False)
        for i, subgrafo in enumerate(subgrafos):
            tabla_equipo = [["Nombre", "Deporte", "Musica", "Danza", "Instrumento", "Club", "Hobbie"]]
            for persona in subgrafo.nodes:
                gustos = persona.get_gustos_str().split(", ")
                fila = [persona.nombre] + gustos
                tabla_equipo.append(fila)
            print(f'\nEquipo {i+1}:')
            print(tabulate(tabla_equipo, headers="firstrow", tablefmt='fancy_grid'))
        if len(subgrafos) == 0:
            print("No se encontraron equipos.")
        input()
    elif opcion == '0':
        print("Saliendo del programa.")
        break
    else:
        print("Opción no válida. Por favor, seleccione una opción válida.")
