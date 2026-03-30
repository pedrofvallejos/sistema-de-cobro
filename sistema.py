from datetime import datetime, timedelta
import os
import platform
import shutil  # <-- Para hacer la copia de seguridad
from tkinter import messagebox
import customtkinter as ctk
from dateutil.relativedelta import relativedelta
import pandas as pd

# Librerías para generar el PDF (Requiere: pip install reportlab)
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

ARCHIVO_EXCEL = "clientes_db.xlsx"

def format_moneda_visual(valor):
    try:
        return f"$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "$ 0,00"

def limpiar_monto(texto):
    return texto.replace(".", "").replace(",", ".")

# --- NUEVA FUNCIÓN: COPIA DE SEGURIDAD AUTOMÁTICA ---
def hacer_backup():
    if os.path.exists(ARCHIVO_EXCEL):
        carpeta_backups = "Backups"
        if not os.path.exists(carpeta_backups):
            os.makedirs(carpeta_backups)
        
        ahora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_backup = f"Backup_Clientes_{ahora}.xlsx"
        ruta_backup = os.path.join(carpeta_backups, nombre_backup)
        
        try:
            shutil.copy2(ARCHIVO_EXCEL, ruta_backup)
        except Exception as e:
            print(f"No se pudo realizar el backup: {e}")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Ejecutamos el backup al arrancar
        hacer_backup()
        
        if not os.path.exists(ARCHIVO_EXCEL):
            columnas = ["DNI", "Nombre", "Domicilio", "Telefono", "Producto", "Precio Total", 
                        "Cuotas Totales", "Valor Cuota", "Cuotas Pagas", "Saldo", 
                        "Frecuencia", "Fecha Carga"]
            pd.DataFrame(columns=columnas).to_excel(ARCHIVO_EXCEL, index=False)
        
        try:
            self.df_datos = pd.read_excel(ARCHIVO_EXCEL)
            self.df_datos["DNI"] = self.df_datos["DNI"].astype(str)
        except:
            self.df_datos = pd.DataFrame()

        self.title("Francisco - Sistema v9.0 (stuar´s Edition)")
        self.geometry("1400 thuận 850")
        ctk.set_appearance_mode("dark")

        self.grid_columnconfigure(0, weight=6)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)

        # --- PANEL IZQUIERDO (Con pestañas para Morosos) ---
        self.frame_lista = ctk.CTkFrame(self, fg_color="#161616")
        self.frame_lista.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.frame_busq = ctk.CTkFrame(self.frame_lista, fg_color="transparent")
        self.frame_busq.pack(fill="x", padx=30, pady=20)
        
        self.entry_busqueda = ctk.CTkEntry(self.frame_busq, placeholder_text="Buscar cliente o DNI...", height=40)
        self.entry_busqueda.pack(side="left", fill="x", expand=True)
        self.entry_busqueda.bind("<KeyRelease>", lambda e: self.actualizar_lista())
        
        # Pestañas para dividir Clientes y Morosos
        self.tab_listas = ctk.CTkTabview(self.frame_lista, fg_color="transparent")
        self.tab_listas.pack(padx=20, pady=5, fill="both", expand=True)
        self.tab_listas.add("Todos los Clientes")
        self.tab_listas.add("🚨 Clientes en Mora")
        
        self.scroll_clientes = ctk.CTkScrollableFrame(self.tab_listas.tab("Todos los Clientes"), fg_color="transparent")
        self.scroll_clientes.pack(fill="both", expand=True)
        
        self.scroll_morosos = ctk.CTkScrollableFrame(self.tab_listas.tab("🚨 Clientes en Mora"), fg_color="transparent")
        self.scroll_morosos.pack(fill="both", expand=True)

        # --- PANEL DERECHO ---
        self.frame_derecho = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_derecho.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.frame_derecho.grid_columnconfigure(0, weight=1)
        self.frame_derecho.grid_rowconfigure(1, weight=1) # El detalle se expande
        
        # 1. NUEVO: PANEL DE CAJA DIARIA / RESUMEN (Arriba a la derecha)
        self.frame_resumen = ctk.CTkFrame(self.frame_derecho, fg_color="#1a1a1a", border_width=1, border_color="#2e7d32")
        self.frame_resumen.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        
        self.lbl_capital = ctk.CTkLabel(self.frame_resumen, text="Capital Activo: $ 0,00", font=("Segoe UI", 16, "bold"), text_color="#2e7d32")
        self.lbl_capital.pack(pady=(15, 5))
        
        self.lbl_cant_clientes = ctk.CTkLabel(self.frame_resumen, text="Clientes Activos: 0", font=("Segoe UI", 13), text_color="#aaaaaa")
        self.lbl_cant_clientes.pack(pady=(0, 15))
        
        # 2. PANEL DE DETALLE (Centro derecha)
        self.frame_detalle = ctk.CTkFrame(self.frame_derecho, fg_color="#1a1a1a", border_width=1, border_color="#1f538d")
        self.frame_detalle.grid(row=1, column=0, sticky="nsew")
        
        self.lbl_info = ctk.CTkLabel(self.frame_detalle, text="Seleccione un cliente", font=("Segoe UI", 16), justify="left")
        self.lbl_info.pack(pady=30, padx=30, fill="both", expand=True)
        
        self.btn_pagar = ctk.CTkButton(self.frame_detalle, text="💰 REGISTRAR COBRO", state="disabled", height=55, font=("Segoe UI", 16, "bold"), command=self.registrar_pago)
        self.btn_pagar.pack(pady=5, fill="x", padx=40)
        
        self.btn_imprimir = ctk.CTkButton(self.frame_detalle, text="🖨️ IMPRIMIR BOLETA", state="disabled", fg_color="#c0392b", hover_color="#e74c3c", height=55, font=("Segoe UI", 16, "bold"), command=self.generar_remito_pdf)
        self.btn_imprimir.pack(pady=5, fill="x", padx=40)
        
        self.btn_nuevo = ctk.CTkButton(self.frame_derecho, text="+ NUEVO CLIENTE", fg_color="#2e7d32", height=50, command=self.ventana_agregar)
        self.btn_nuevo.grid(row=2, column=0, sticky="ew", pady=(15, 0))

        self.dni_sel = None
        self.actualizar_lista()

    def calcular_mora(self, fila):
        """Calcula fechas de vencimiento y atraso (Con el tope de cuotas)"""
        hoy = datetime.now().date()
        try:
            f_carga = pd.to_datetime(fila['Fecha Carga']).date()
            frec = fila['Frecuencia']
            pagas = int(fila['Cuotas Pagas'])
            totales = int(fila['Cuotas Totales'])
            
            if pagas >= totales:
                return f_carga, hoy, 0, 0
            
            if frec == "Semanal": 
                proximo_venc = f_carga + timedelta(weeks=pagas)
                dias_ciclo = 7
            elif frec == "Quincenal": 
                proximo_venc = f_carga + timedelta(days=pagas * 15)
                dias_ciclo = 15
            else: 
                proximo_venc = f_carga + relativedelta(months=pagas)
                dias_ciclo = 30
            
            dias_atraso = (hoy - proximo_venc).days
            
            if dias_atraso >= 0:
                cuotas_debe = (dias_atraso // dias_ciclo) + 1
                cuotas_restantes = totales - pagas
                if cuotas_debe > cuotas_restantes:
                    cuotas_debe = cuotas_restantes
            else:
                cuotas_debe = 0
            
            return f_carga, proximo_venc, max(0, dias_atraso), cuotas_debe
        except:
            return hoy, hoy, 0, 0

    def actualizar_lista(self):
        # Limpiamos los dos scrolls
        for w in self.scroll_clientes.winfo_children(): w.destroy()
        for w in self.scroll_morosos.winfo_children(): w.destroy()
            
        if self.df_datos.empty: return
        busq = self.entry_busqueda.get().lower()
        
        capital_total = 0.0
        clientes_activos = 0
        lista_morosos = []
        
        for _, c in self.df_datos.iterrows():
            f_carga, f_venc, atraso, debe = self.calcular_mora(c)
            saldo = float(c['Saldo'])
            
            # Sumamos al resumen si el cliente aún debe
            if saldo > 0:
                capital_total += saldo
                clientes_activos += 1
            
            # Filtro de búsqueda
            if busq and (busq not in str(c['Nombre']).lower() and busq not in str(c['DNI'])): 
                continue
            
            # Definimos colores y estados
            if saldo <= 0:
                color, estado = "#454545", "FINALIZADO"
            elif atraso > 0:
                color, estado = "#b22222", f"MORA: {atraso}d (Debe {debe})"
                # Guardamos a los morosos para ordenarlos después
                lista_morosos.append((c, atraso, debe, color, estado))
            else:
                color, estado = "#2e7d32", "AL DÍA"
                
            txt = f"{c['Nombre']} | {estado} | Saldo: {format_moneda_visual(saldo)}"
            ctk.CTkButton(self.scroll_clientes, text=txt, fg_color=color, anchor="w", 
                          command=lambda d=c['DNI']: self.ver_detalle(d)).pack(pady=2, fill="x", padx=10)

        # Cargar la pestaña de morosos ordenada por mayor atraso
        lista_morosos.sort(key=lambda x: x[1], reverse=True)
        for c, atraso, debe, color, estado in lista_morosos:
            txt = f"{c['Nombre']} | {estado} | Saldo: {format_moneda_visual(c['Saldo'])}"
            ctk.CTkButton(self.scroll_morosos, text=txt, fg_color=color, anchor="w", 
                          command=lambda d=c['DNI']: self.ver_detalle(d)).pack(pady=2, fill="x", padx=10)
            
        # Actualizamos el panel de resumen
        self.lbl_capital.configure(text=f"Capital en la Calle: {format_moneda_visual(capital_total)}")
        self.lbl_cant_clientes.configure(text=f"Clientes Activos: {clientes_activos}")

    def ver_detalle(self, dni):
        self.dni_sel = str(dni)
        c = self.df_datos[self.df_datos["DNI"] == self.dni_sel].iloc[0]
        f_ini, f_venc, atraso, debe = self.calcular_mora(c)
        
        texto = (f"👤 CLIENTE: {c['Nombre']}\n"
                 f"🆔 DNI: {c['DNI']}  |  📞 TEL: {c['Telefono']}\n"
                 f"🏠 DOMICILIO: {c['Domicilio']}\n"
                 f"📦 PRODUCTO: {c['Producto']}\n"
                 f"──────────────────────────────\n"
                 f"📅 FECHA DE COMPRA: {f_ini.strftime('%d/%m/%Y')}\n"
                 f"📊 CUOTAS PAGAS: {int(c['Cuotas Pagas'])} de {int(c['Cuotas Totales'])}\n"
                 f"💵 VALOR CUOTA: {format_moneda_visual(c['Valor Cuota'])}\n"
                 f"💰 SALDO RESTANTE: {format_moneda_visual(c['Saldo'])}\n"
                 f"──────────────────────────────\n"
                 f"🔔 PRÓXIMO COBRO: {f_venc.strftime('%d/%m/%Y')}\n"
                 f"⚠️ ATRASO ACTUAL: {atraso} días\n"
                 f"📉 CUOTAS EN MORA: {debe if atraso > 0 else 0}")
        
        self.lbl_info.configure(text=texto)
        self.btn_pagar.configure(state="normal" if float(c['Saldo']) > 0 else "disabled")
        self.btn_imprimir.configure(state="normal")

    def registrar_pago(self):
        idx = self.df_datos.index[self.df_datos["DNI"] == self.dni_sel].tolist()[0]
        v_cuota = float(self.df_datos.at[idx, "Valor Cuota"])
        if messagebox.askyesno("Cobro", f"¿Registrar pago de {format_moneda_visual(v_cuota)}?"):
            self.df_datos.at[idx, "Saldo"] -= v_cuota
            self.df_datos.at[idx, "Cuotas Pagas"] += 1
            self.df_datos.to_excel(ARCHIVO_EXCEL, index=False)
            self.actualizar_lista(); self.ver_detalle(self.dni_sel)

    def generar_remito_pdf(self):
        if not self.dni_sel: return
        c = self.df_datos[self.df_datos["DNI"] == self.dni_sel].iloc[0]
        
        carpeta_boletas = "Boletas"
        if not os.path.exists(carpeta_boletas): os.makedirs(carpeta_boletas)
        
        nombre_limpio = "".join(x for x in str(c['Nombre']) if x.isalnum() or x in "._- ")
        nombre_archivo = f"Remito_{c['DNI']}_{nombre_limpio.replace(' ', '_')}.pdf"
        ruta_completa = os.path.join(carpeta_boletas, nombre_archivo)
        
        doc = SimpleDocTemplate(ruta_completa, pagesize=letter,
                                rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
        story = []
        styles = getSampleStyleSheet()
        style_normal = styles['Normal']
        
        style_empresa = ParagraphStyle('Empresa', parent=style_normal, fontName='Helvetica-Bold', fontSize=24, leading=28, textColor=colors.HexColor("#1f538d"))
        style_subtitulo = ParagraphStyle('Sub', parent=style_normal, fontName='Helvetica-Bold', fontSize=11, leading=14)
        style_garantia = ParagraphStyle('Garantia', parent=style_normal, fontName='Helvetica-Oblique', fontSize=9, leading=12, alignment=4)
        
        datos_cabecera = []
        if os.path.exists("logo.png"):
            from reportlab.platypus import Image
            img_logo = Image("logo.png", 2.5*cm, 2.5*cm)
            datos_cabecera.append([img_logo, Paragraph("Stuar´s<br/><font size=10 color='#555555'>Sistema de Créditos y Cobranzas</font>", style_empresa)])
        else:
            datos_cabecera.append([Paragraph("Stuar´s", style_empresa), Paragraph("<b>REMITO DE COMPRA</b><br/>Documento no válido como factura", ParagraphStyle('Der', parent=style_normal, alignment=2))])
            
        tabla_cabecera = Table(datos_cabecera, colWidths=[5*cm, 12.5*cm])
        tabla_cabecera.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        story.append(tabla_cabecera)
        
        story.append(Spacer(1, 0.5*cm))
        story.append(Paragraph("<b>X</b>", ParagraphStyle('Centro', parent=style_normal, alignment=1, fontSize=18)))
        story.append(Spacer(1, 0.5*cm))
        
        f_carga = pd.to_datetime(c['Fecha Carga'])
        datos_cliente = [
            [Paragraph("<b>Cliente:</b>", style_subtitulo), Paragraph(str(c['Nombre']), style_normal), Paragraph("<b>Fecha:</b>", style_subtitulo), Paragraph(f_carga.strftime('%d/%m/%Y'), style_normal)],
            [Paragraph("<b>DNI:</b>", style_subtitulo), Paragraph(str(c['DNI']), style_normal), Paragraph("<b>Teléfono:</b>", style_subtitulo), Paragraph(str(c['Telefono']), style_normal)],
            [Paragraph("<b>Domicilio:</b>", style_subtitulo), Paragraph(str(c['Domicilio']), style_normal), Paragraph("<b>Frecuencia:</b>", style_subtitulo), Paragraph(str(c['Frecuencia']), style_normal)]
        ]
        
        tabla_cliente = Table(datos_cliente, colWidths=[2.5*cm, 6.5*cm, 2.5*cm, 6*cm])
        tabla_cliente.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#f4f4f4")),
            ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#f4f4f4")),
            ('PADDING', (0,0), (-1,-1), 6),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(tabla_cliente)
        story.append(Spacer(1, 0.6*cm))
        
        story.append(Paragraph("DETALLE DEL PRODUCTO", style_subtitulo))
        story.append(Spacer(1, 0.2*cm))
        
        datos_producto = [
            ["Producto / Descripción", "Precio Total", "Cuotas", "Valor Cuota"],
            [str(c['Producto']), format_moneda_visual(c['Precio Total']), str(int(c['Cuotas Totales'])), format_moneda_visual(c['Valor Cuota'])]
        ]
        
        tabla_producto = Table(datos_producto, colWidths=[8.5*cm, 3*cm, 3*cm, 3*cm])
        tabla_producto.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f538d")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('PADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(tabla_producto)
        story.append(Spacer(1, 0.8*cm))
        
        story.append(Paragraph("CALENDARIO DE VENCIMIENTOS", style_subtitulo))
        story.append(Spacer(1, 0.2*cm))
        
        datos_vencimientos = [["N° Cuota", "Fecha de Vencimiento", "Monto", "Estado"]]
        frec = c['Frecuencia']
        for i in range(1, int(c['Cuotas Totales']) + 1):
            if i == 1: f_venc = f_carga
            else:
                if frec == "Semanal": f_venc = f_carga + timedelta(weeks=i-1)
                elif frec == "Quincenal": f_venc = f_carga + timedelta(days=(i-1) * 15)
                else: f_venc = f_carga + relativedelta(months=i-1)
            
            estado = "PAGADA" if i <= int(c['Cuotas Pagas']) else "PENDIENTE"
            datos_vencimientos.append([f"Cuota {i}", f_venc.strftime('%d/%m/%Y'), format_moneda_visual(c['Valor Cuota']), estado])
            
        tabla_vencimientos = Table(datos_vencimientos, colWidths=[4*cm, 5.5*cm, 4*cm, 4*cm])
        tabla_vencimientos.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#555555")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(tabla_vencimientos)
        story.append(Spacer(1, 1*cm))
        
        clausula = ("<b>CLÁUSULA DE GARANTÍA:</b> Stuar´s garantiza el correcto funcionamiento del producto "
                    "por desperfectos de fabricación por un período de 30 días a partir de la fecha de entrega. "
                    "La garantía quedará anulada si el producto presenta golpes, roturas, enmiendas, o uso indebido. "
                    "La falta de pago de dos o más cuotas consecutivas otorgará el derecho a la empresa de "
                    "proceder al retiro del bien adquirido sin derecho a reclamo de sumas abonadas previamente.")
        story.append(Paragraph(clausula, style_garantia))
        story.append(Spacer(1, 2*cm))
        
        datos_firmas = [["..................................................", ".................................................."],
                        ["Firma del Cliente", "Firma Autorizada Stuar´s"]]
        tabla_firmas = Table(datos_firmas, colWidths=[8.75*cm, 8.75*cm])
        tabla_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'), ('FONTSIZE', (0,1), (-1,1), 10)]))
        story.append(tabla_firmas)
        
        doc.build(story)
        
        # Abrir el PDF de golpe en pantalla
        try:
            sistema = platform.system()
            if sistema == "Windows": os.startfile(ruta_completa)
            elif sistema == "Darwin": os.system(f"open '{ruta_completa}'")
            else: os.system(f"xdg-open '{ruta_completa}'")
        except: pass
            
        messagebox.showinfo("Éxito", f"Se ha generado la boleta en la carpeta Boletas:\n{nombre_archivo}")

    def ventana_agregar(self):
        vent = ctk.CTkToplevel(self)
        vent.geometry("500x850"); vent.title("Nuevo Crédito"); vent.attributes("-topmost", True)
        sf = ctk.CTkScrollableFrame(vent, fg_color="transparent")
        sf.pack(fill="both", expand=True, padx=10, pady=10)

        entradas = {}
        for campo in ["DNI", "Nombre", "Domicilio", "Telefono", "Producto", "Precio", "Cuotas"]:
            ctk.CTkLabel(sf, text=f"{campo} *").pack(pady=(5,0))
            e = ctk.CTkEntry(sf, width=300)
            e.pack(pady=2)
            entradas[campo] = e

        ctk.CTkLabel(sf, text="Frecuencia *").pack()
        frec_v = ctk.StringVar(value="Semanal")
        ctk.CTkComboBox(sf, values=["Semanal", "Quincenal", "Mensual"], variable=frec_v, width=300).pack()

        lbl_v = ctk.CTkLabel(sf, text="CUOTA: $ 0,00", font=("", 20, "bold"), text_color="#2e7d32")
        lbl_v.pack(pady=20)

        def formatear_precio_evento(e):
            entrada = entradas["Precio"].get()
            solo_numeros = "".join(filter(lambda x: x.isdigit() or x == ",", entrada))
            if solo_numeros:
                if "," in solo_numeros:
                    p = solo_numeros.split(",")
                    entero = "{:,}".format(int(p[0] or 0)).replace(",", ".")
                    decimal = p[1][:2]
                    entradas["Precio"].delete(0, "end"); entradas["Precio"].insert(0, f"{entero},{decimal}")
                else:
                    entradas["Precio"].delete(0, "end"); entradas["Precio"].insert(0, "{:,}".format(int(solo_numeros)).replace(",", "."))
            recalcular()

        def recalcular():
            try:
                p = float(limpiar_monto(entradas["Precio"].get()))
                c = int(entradas["Cuotas"].get())
                lbl_v.configure(text=f"CUOTA: {format_moneda_visual(p/c)}")
            except: lbl_v.configure(text="CUOTA: $ ---")

        entradas["Precio"].bind("<KeyRelease>", formatear_precio_evento)
        entradas["Cuotas"].bind("<KeyRelease>", lambda e: recalcular())

        def guardar():
            for c in entradas:
                if not entradas[c].get().strip(): return messagebox.showwarning("Falta", f"Completar {c}")
            try:
                p = float(limpiar_monto(entradas["Precio"].get()))
                c = int(entradas["Cuotas"].get())
                v_c = p/c
                nueva_fila = {
                    "DNI": entradas["DNI"].get(), "Nombre": entradas["Nombre"].get().upper(),
                    "Domicilio": entradas["Domicilio"].get().upper(), "Telefono": entradas["Telefono"].get(),
                    "Producto": entradas["Producto"].get().upper(), "Precio Total": p,
                    "Cuotas Totales": c, "Valor Cuota": v_c, "Cuotas Pagas": 1, "Saldo": p - v_c,
                    "Frecuencia": frec_v.get(), "Fecha Carga": datetime.now().strftime("%Y-%m-%d")
                }
                self.df_datos = pd.concat([self.df_datos, pd.DataFrame([nueva_fila])], ignore_index=True)
                self.df_datos.to_excel(ARCHIVO_EXCEL, index=False)
                self.actualizar_lista(); vent.destroy()
                messagebox.showinfo("Éxito", "Venta guardada. Se descontó la 1° cuota.")
            except: messagebox.showerror("Error", "Revisar montos.")

        ctk.CTkButton(sf, text="GUARDAR VENTA", fg_color="#2e7d32", height=50, command=guardar).pack(pady=20)

if __name__ == "__main__":
    app = App()
    app.mainloop()