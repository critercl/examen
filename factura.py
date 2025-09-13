try:
    import sqlite3
    import os
    from rich.table import Table
    from rich.prompt import Prompt, IntPrompt, Confirm
    from rich.panel import Panel
    from rich.console import Console
    from colorama import init, Fore, Back, Style
    from datetime import date
    import xlsxwriter

except ImportError as e:
    print("\n")
    print("-"*80)
    print("🛠️ Estimado Profesor JOSE MAUREIRA para poder correr esta aplicacion debe tener instaladas la siguientes Librerias 🛠️\n")
    print(f" 🐍 Libreria rich, forma de instalacion: {e}")
    print(f"\n el detalle del error es : {e}")
    print("-"*80)
    print("\n\n")
    exit(1)

class Factura:
    # Metodo Construtor, Define el enombre de la bae de datos
    def __init__(self, base='facturas.db'):
        self.base=base
        self.conn=None
    #metodo  Conectar a la base de datos 
    def conectar(self):
        try:
            self.conn=sqlite3.connect(self.base)
            return True
        except sqlite3.Error as e:
            self.conn=None
            return e
    # Metodo Cerrar Base Datos
    def cerrar(self):
        self.conn.close()
        self.conn=None
    # Crear Tabla de Facturador
    def CreaTabla(self):
        cursor=self.conn.cursor()
        try:
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS "CabeceraFactura" (
                    "numeroFactura"	INTEGER,
                    "nombreCliente"	TEXT NOT NULL,
                    "rutCliente"	TEXT NOT NULL,
                    "fechaEmision"	DATE NOT NULL,
                    "FormaPago"	INTEGER NOT NULL,
                    "iva"	INTEGER NOT NULL,
                    "TotalBruto"	INTEGER NOT NULL,
                    "totalCompra"	INTEGER NOT NULL,
                    PRIMARY KEY("numeroFactura")
                );
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS "DetalleFactura" (
                    "idDetalle"	INTEGER,
                    "numeroFactura"	INTEGER NOT NULL,
                    "nombreProducto"	TEXT NOT NULL,
                    "cantidad"	INTEGER NOT NULL,
                    "precioUnitario"	INTEGER NOT NULL,
                    "totalItem"	INTEGER NOT NULL,
                    PRIMARY KEY("idDetalle" AUTOINCREMENT)
                );
            ''')
            return True
        except sqlite3 as e:
            return e
    # Insertar Cabecera de Factura
    def InsertarCabecera(self,numero,nombre,rut,fecha,pago,iva,bruto,total):
        cursor=self.conn.cursor()
        query='''INSERT INTO "CabeceraFactura" ("numeroFactura", "nombreCliente", "rutCliente", "fechaEmision","FormaPago","iva","TotalBruto","totalCompra") VALUES (?, ?, ?, ?,?, ?, ?, ?)'''
        fecha=fecha.isoformat()
        cursor.execute(query,(numero,nombre,rut,fecha,pago,iva,bruto,total))
        self.conn.commit()
    # Insertar Detalle de Factura
    def InsertarDetalle(self,numero,producto,cantidad,precio,total):
        cursor=self.conn.cursor()
        query='''INSERT INTO "DetalleFactura" ("numeroFactura", "nombreCliente", "rutCliente", "fechaEmision","FormaPago","iva","TotalBruto","totalCompra") VALUES (?, ?, ?, ?,?, ?, ?, ?)'''
        fecha=fecha.isoformat()
        cursor.execute(query,(numero,producto,cantidad,precio,total))
        self.conn.commit()
def menu():
    while True:
        console.print(Panel("[bold blue]📦 Sistema de Gestión de Facturas - REQ04 [/bold blue]", expand=False))
        opciones = [
            "1. Registrar nueva factura",
            "2. Exportar facturas a Excel (por rango de fechas)",
            "3. Ver todas las facturas",
            "4. Salir"
        ]
        for opcion in opciones:
            console.print(f"[cyan]{opcion}[/cyan]")
        eleccion = Prompt.ask("\n👉 Seleccione una opción", choices=["1", "2", "3", "4"], default="4")
        match eleccion:
            case "1":
                registrarFactura()
            case "4":
                break

def registrarFactura():
    producto=[]
    cantidad=[]
    valor=[]
    console.print("[bold magenta] Registro de Facturas [/bold magenta] ")
    numero=IntPrompt.ask("Numero de Factura: ")
    rut=Prompt.ask("Rut Cliente")
    nombre=Prompt.ask("Nombre Cliente: ")
    fecha=Prompt.ask("Fecha Factura (YYY-MM-DD) ",default=date.today().isoformat())
    pago=Prompt.ask("Forma de Pago 1: Efectivo, 2: Tarjeta, 3: Transferencia", choices=["1","2","3"], default="1")
    pos=0
    while True:
        aux_prod=Prompt.ask("Nombre de Producto :")
        aux_cant=IntPrompt.ask("Cantidad a Comprar :")
        aux_valor=IntPrompt.ask("Valor Producto")
        producto.append(aux_prod)
        cantidad.append(aux_cant)
        pos+=1

        





def Limpiar():
    if os.name=="nt":
        os.system("cls")
    else:
        os.system('clear')
# Formato Numerico Chileno
def formato_chileno(numero):
    #formatea número chileno sin decimales
    return f"{int(numero):,}".replace(",", ".") 

console = Console()
Limpiar()
#Crea Base Datos
gestor=Factura("facturas.db")
#Conecta a la Base
result=gestor.conectar()
if result==True:
    print(Fore.GREEN+"✅ Se conecto La base Correctamente")
else:
    print(Fore.RED+f"🔴 Error de Conexion a la Base {result}") 

#Crea Tabla
result=gestor.CreaTabla()
if result==True:
    print(Fore.GREEN+"✅ Se crea Tablas y base de Datos REQ01")
else:
    print(Fore.RED+f"🔴 Error al crear base y tablas {result}") 
#Creacion de Clase
print(Fore.GREEN+"✅ Se crea Creacion de Clase Factura - REQ02 ")
print(Fore.GREEN+"✅ Se crea Metodo insertar Cabecera y Detalle - REQ03 ")

menu()
