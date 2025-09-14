try:
    import sqlite3
    import os
    from rich.table import Table
    from rich.prompt import Prompt, IntPrompt, Confirm
    from rich.panel import Panel
    from rich.console import Console
    from colorama import init, Fore, Back, Style
    from datetime import date, datetime
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
        cursor.execute(query,(numero,nombre,rut,fecha,pago,iva,bruto,total))
        self.conn.commit()
    # Buscar Nunmero de Factura
    def BuscvarNroFactura(self,numero):
        cursor=self.conn.cursor()
        query="select count(*)  from CabeceraFactura where numeroFactura=?"
        cursor.execute(query,(numero,))
        return  cursor.fetchone()[0]
    def ExportaraExcel(self,inicio,termino):
        FormasPagos=["EFECTIVO","TARJETA","TRANSFERENCIA"]
        query="select * from CabeceraFactura Where fechaEmision>=? and fechaEmision<=?"
        cursor=self.conn.cursor()
        cursor.execute(query,(inicio,termino))
        result=cursor.fetchall()
        archivo=f"Exportacion_Factura_desde_{inicio}_hasta_{termino}.xlsx"
        console.print("[bold green]📊 se Crea Archivo Excel "+ archivo+"[/bold green]")
        excel = xlsxwriter.Workbook(archivo)
        libro=excel.add_worksheet(f"Facturas_{inicio}_{termino}")
        #### Formatos para el Excel ###
        formato_titulo = excel.add_format({'bold': True,'font_size': 16,'bg_color': '#D3D3D3', 'align': 'center','underline': True})
        formato_encabezado = excel.add_format({'bold': True,'font_size': 12,'bg_color': '#D3D3D3','border': 1,'align': 'center'})
        formato_numero = excel.add_format({'num_format': '#,##0.00'})

        libro.merge_range('A1:P1', 'Listado de Facturas Desde '+str(inicio)+" AL "+str(termino),formato_titulo)
        fila=1
        for row in result:
            libro.write(fila,0,"Numero de Factura",formato_encabezado)
            libro.write(fila,1,row[0],formato_encabezado)
            libro.write(fila,2,"Cliente",formato_encabezado)
            libro.write(fila,3,row[1],formato_encabezado)
            libro.write(fila,4,"RUT",formato_encabezado)
            libro.write(fila,5,row[2],formato_encabezado)
            libro.write(fila,6,"Fecha",formato_encabezado)
            libro.write(fila,7,row[3],formato_encabezado)
            libro.write(fila,8,"Forma Pago",formato_encabezado)
            libro.write(fila,9,FormasPagos[row[4]-1],formato_encabezado)
            libro.write(fila,10,"Total Neto",formato_encabezado)
            libro.write(fila,11,row[6],formato_encabezado)
            libro.write(fila,12,"Iva",formato_encabezado)
            libro.write(fila,13,row[5],formato_encabezado)
            libro.write(fila,14,"Total a Pagar",formato_encabezado)
            libro.write(fila,15,row[7],formato_encabezado)            
            query="SELECT * FROM DetalleFactura WHERE numeroFactura=?"
            cursor.execute(query,(row[0],))
            result_posicion=cursor.fetchall()
            fila+=2
            libro.merge_range(f'A{fila}:P{fila}', f'DETALLE FACTURA N° {row[0]}',formato_titulo)
            fila+=1
            libro.merge_range(f'A{fila}:C{fila}', 'POSICION',formato_titulo)
            libro.merge_range(f'D{fila}:F{fila}', 'PRODUCTO',formato_titulo)
            libro.merge_range(f'G{fila}:I{fila}', 'CANTIDAD',formato_titulo)
            libro.merge_range(f'J{fila}:L{fila}', 'PRECIO',formato_titulo)
            libro.merge_range(f'M{fila}:P{fila}', 'TOTAL POSICION',formato_titulo)
            fila+=1
            num=1
            for pos in result_posicion:
                libro.merge_range(f'A{fila}:C{fila}', num,formato_encabezado)
                libro.merge_range(f'D{fila}:F{fila}', pos[2],formato_encabezado)
                libro.merge_range(f'G{fila}:I{fila}', pos[3],formato_encabezado)
                libro.merge_range(f'J{fila}:L{fila}', pos[4],formato_encabezado)
                libro.merge_range(f'M{fila}:P{fila}', pos[5],formato_encabezado)
                fila+=1
                num+=1

        excel.close()
    # Insertar Detalle de Factura
    def InsertarDetalle(self,numero,producto,cantidad,precio,total):
        cursor=self.conn.cursor()
        query='''INSERT INTO "DetalleFactura" ("numeroFactura", "nombreProducto", "cantidad", "precioUnitario","totalItem") VALUES (?, ?, ?, ?, ?)'''
        cursor.execute(query,(numero,producto,cantidad,precio,total))
        self.conn.commit()
def menu():
    while True:
        console.print(Panel("[bold blue]📦 Sistema de Gestión de Facturas - REQ04 [/bold blue]", expand=False))
        opciones = [
            "1. Registrar nueva factura",
            "2. Exportar facturas a Excel (por rango de fechas)",
            "3. Salir"
        ]
        for opcion in opciones:
            console.print(f"[cyan]{opcion}[/cyan]")
        eleccion = Prompt.ask("\n👉 Seleccione una opción", choices=["1", "2", "3"], default="3")
        match eleccion:
            case "1":
                registrarFactura()
            case "2":
                console.print(Panel("[bold blue]📤 EXPORTAR FACTURAS A EXCEL [/bold blue]", expand=False))
                while True:
                    inicio=Prompt.ask("📅 Fecha Inicio (YYYY-MM-DD) ",default=date.today().isoformat())
                    try:
                        inicio = datetime.strptime(inicio, "%Y-%m-%d").date()
                        inicio=str(inicio)
                        break
                    except ValueError:
                        console.print(f"[bold red] Error en el formato de Fecha (YYYY-MM-DD) Ejemplo: {date.today().isoformat()} [/bold red] ")
                while True:
                    termino=Prompt.ask("📅 Fecha Termino (YYYY-MM-DD) ",default=date.today().isoformat())
                    try:
                        termino = datetime.strptime(termino, "%Y-%m-%d").date()
                        termino=str(termino)
                        break
                    except ValueError:
                        console.print(f"[bold red] Error en el formato de Fecha (YYYY-MM-DD) Ejemplo: {date.today().isoformat()} [/bold red] ")
                result=gestor.ExportaraExcel(inicio,termino)
            case "3":
                break

def registrarFactura():
    producto=[]
    cantidad=[]
    valor=[]
    totalItem=[]
    iva=0
    bruto=0
    console.print("[bold magenta] Registro de Facturas [/bold magenta] ")
    numero=IntPrompt.ask("🔢 Numero de Factura ")
    result=gestor.BuscvarNroFactura(numero)
    if result==1:
        console.print("❌ La Factura que esta intentando de ingresar ya existe en la base ❌", style="bold red")
        console.print("Presione cualquier tecla para continuar...", style="bold cyan")
        console.input()
        return False
    rut=Prompt.ask("🆔 Rut Cliente")
    nombre=Prompt.ask("👤 Nombre Cliente ")
    while True:
        fecha=Prompt.ask("📅 Fecha Factura (YYYY-MM-DD) ",default=date.today().isoformat())
        try:
            fecha_str = datetime.strptime(fecha, "%Y-%m-%d").date()
            break
        except ValueError:
                console.print(f"[bold red] Error en el formato de Fecha (YYYY-MM-DD) Ejemplo: {date.today().isoformat()} [/bold red] ")
    pago=Prompt.ask("Forma de Pago 1: Efectivo, 2: Tarjeta, 3: Transferencia", choices=["1","2","3"], default="1")
    pos=0
    while True:
        aux_prod=Prompt.ask("Nombre de Producto ")
        aux_cant=IntPrompt.ask("Cantidad a Comprar ")
        aux_valor=IntPrompt.ask("Valor Producto ")
        
        bruto+=(aux_cant*aux_valor)
        aux_total=round((aux_cant*aux_valor)*1.19,0)
        iva+=round((aux_cant*aux_valor)*0.19,0)
        producto.append(aux_prod)
        cantidad.append(aux_cant)
        valor.append(aux_valor)
        totalItem.append(aux_total)
        console.print(f"Neto {(aux_cant*aux_valor)}")
        console.print(f"Iva {round((aux_cant*aux_valor)*0.19,0)}")
        console.print(f"Total {totalItem}")
        
        op=Prompt.ask("Desea Continuar Ingresando ", choices=["S","N"])
        if op=="N":
            break
        print(producto[pos])
        pos+=1
    gestor.InsertarCabecera(numero,nombre,rut,fecha,pago,iva,bruto,(bruto+iva))
    for pos in range(len(producto)):
        gestor.InsertarDetalle(numero,producto[pos],cantidad[pos],valor[pos],round(totalItem[pos],0)) 

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


# Acceso al menu principal
menu()
