from datetime import datetime, timedelta
import os
import platform
import shutil
import webbrowser
from urllib.parse import quote
from tkinter import messagebox
import customtkinter as ctk
from dateutil.relativedelta import relativedelta
import pandas as pd

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, Image # <-- Agregado Image aquí
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

ARCHIVO_EXCEL = "clientes_db.xlsx"

# --- CONFIGURACIÓN DE TU NEGOCIO ---
NOMBRE_NEGOCIO = "STUAR'S"
RUBRO_NEGOCIO = "ELECTRODOMESTICOS, INSUMOS Y ACCESORIOS"
DIRECCION_NEGOCIO = "Av. Luis Vernet 3247"
TELEFONO_NEGOCIO = "11 3028-1518"

# --- CONFIGURACIÓN DE MORA ---
DIAS_MORA_PARA_PUNITORIO = 30 # Equivale a un mes

def format_moneda_visual(valor):
    try:
        return f"$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "$ 0,00"

def limpiar_monto(texto):
    return texto.replace(".", "").replace(",", ".")

def hacer_backup():
    if os.path.exists(ARCHIVO_EXCEL):
        carpeta_backups = "Backups"
        if not os.path.exists(carpeta_backups): os.makedirs(carpeta_backups)
        ahora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_backup = f"Backup_Clientes_{ahora}.xlsx"
        ruta_backup = os.path.join(carpeta_backups, nombre_backup)
        try: shutil.copy2(ARCHIVO_EXCEL, ruta_backup)
        except: pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        hacer_backup()
        
        if not os.path.exists(ARCHIVO_EXCEL):
            columnas = ["DNI", "Nombre", "Domicilio", "Telefono", "Producto", "Precio Total", 
                        "Cuotas Totales", "Valor Cuota", "Cuotas Pagas", "Saldo", 
                        "Frecuencia", "Fecha Carga"]
            pd.DataFrame(columns=columnas).to_excel(ARCHIVO_EXCEL, index=False)
        
        try:
            self.df_datos = pd.read_excel(ARCHIVO_EXCEL)
            self.df_datos["DNI"] = self.df_datos["DNI"].astype(str)
            self.df_datos["Telefono"] = self.df_datos["Telefono"].astype(str).str.replace(".0", "", regex=False)
        except:
            self.df_datos = pd.DataFrame()

        self.title(f"Francisco - Sistema {NOMBRE_NEGOCIO} v14.0")
        self.geometry("1400x850")
        ctk.set_appearance_mode("dark")

        self.grid_columnconfigure(0, weight=6)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)

        # --- PANEL IZQUIERDO ---
        self.frame_lista = ctk.CTkFrame(self, fg_color="#161616")
        self.frame_lista.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.frame_busq = ctk.CTkFrame(self.frame_lista, fg_color="transparent")
        self.frame_busq.pack(fill="x", padx=30, pady=20)
        self.entry_busqueda = ctk.CTkEntry(self.frame_busq, placeholder_text="Buscar cliente o DNI...", height=40)
        self.entry_busqueda.pack(side="left", fill="x", expand=True)
        self.entry_busqueda.bind("<KeyRelease>", lambda e: self.actualizar_lista())
        
        self.tab_listas = ctk.CTkTabview(self.frame_lista, fg_color="transparent")
        self.tab_listas.pack(padx=20, pady=5, fill="both", expand=True)
        self.tab_listas.add("Todos los Clientes")
        self.tab_listas.add("🚨 Clientes en Mora")
        self.tab_listas.add("🧮 Calculadora de Mora")
        
        self.scroll_clientes = ctk.CTkScrollableFrame(self.tab_listas.tab("Todos los Clientes"), fg_color="transparent")
        self.scroll_clientes.pack(fill="both", expand=True)
        self.scroll_morosos = ctk.CTkScrollableFrame(self.tab_listas.tab("🚨 Clientes en Mora"), fg_color="transparent")
        self.scroll_morosos.pack(fill="both", expand=True)
        
        self.frame_calc = ctk.CTkFrame(self.tab_listas.tab("🧮 Calculadora de Mora"), fg_color="transparent")
        self.frame_calc.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.lbl_calc_cliente = ctk.CTkLabel(self.frame_calc, text="Seleccione un cliente moroso para calcular", font=("Segoe UI", 16, "bold"), text_color="#aaaaaa")
        self.lbl_calc_cliente.pack(pady=10)
        
        self.lbl_calc_detalle = ctk.CTkLabel(self.frame_calc, text="Cuota pura: $ 0,00\nDías de atraso: 0", font=("Segoe UI", 14), justify="left")
        self.lbl_calc_detalle.pack(pady=10)
        
        ctk.CTkLabel(self.frame_calc, text="Punitorio a cobrar ($):", font=("Segoe UI", 13)).pack(pady=(10,0))
        self.entry_punitorio = ctk.CTkEntry(self.frame_calc, placeholder_text="Ej: 2500", width=200, height=35)
        self.entry_punitorio.pack(pady=5)
        self.entry_punitorio.bind("<KeyRelease>", lambda e: self.recalcular_total_mora())
        
        self.lbl_total_mora = ctk.CTkLabel(self.frame_calc, text="TOTAL A COBRAR: $ 0,00", font=("Segoe UI", 18, "bold"), text_color="#e67e22")
        self.lbl_total_mora.pack(pady=15)
        
        self.btn_guardar_punitivo = ctk.CTkButton(self.frame_calc, text="💾 APLICAR COBRO CON PUNITORIO", fg_color="#d35400", hover_color="#e67e22", height=40, state="disabled", command=self.registrar_pago_con_punitorio)
        self.btn_guardar_punitivo.pack(pady=10)

        # --- PANEL DERECHO ---
        self.frame_derecho = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_derecho.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.frame_derecho.grid_columnconfigure(0, weight=1)
        self.frame_derecho.grid_rowconfigure(1, weight=1)

        self.frame_resumen = ctk.CTkFrame(self.frame_derecho, fg_color="#1a1a1a", border_width=1, border_color="#2e7d32")
        self.frame_resumen.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        self.lbl_capital = ctk.CTkLabel(self.frame_resumen, text="Capital Activo: $ 0,00", font=("Segoe UI", 16, "bold"), text_color="#2e7d32")
        self.lbl_capital.pack(pady=(15, 5))
        self.lbl_cant_clientes = ctk.CTkLabel(self.frame_resumen, text="Clientes Activos: 0", font=("Segoe UI", 13), text_color="#aaaaaa")
        self.lbl_cant_clientes.pack(pady=(0, 15))
        
        self.frame_detalle = ctk.CTkFrame(self.frame_derecho, fg_color="#1a1a1a", border_width=1, border_color="#1f538d")
        self.frame_detalle.grid(row=1, column=0, sticky="nsew")
        self.lbl_info = ctk.CTkLabel(self.frame_detalle, text="Seleccione un cliente", font=("Segoe UI", 14), justify="left")
        self.lbl_info.pack(pady=15, padx=30, fill="both", expand=True)
        
        self.btn_pagar = ctk.CTkButton(self.frame_detalle, text="💰 REGISTRAR COBRO", state="disabled", height=45, font=("Segoe UI", 14, "bold"), command=self.registrar_pago)
        self.btn_pagar.pack(pady=3, fill="x", padx=40)

        self.btn_whatsapp = ctk.CTkButton(self.frame_detalle, text="💬 ENVIAR COMPROBANTE DE PAGO", state="disabled", fg_color="#25D366", hover_color="#128C7E", text_color="black", height=45, font=("Segoe UI", 13, "bold"), command=self.enviar_whatsapp)
        self.btn_whatsapp.pack(pady=3, fill="x", padx=40)
        
        self.btn_aviso_vencimiento = ctk.CTkButton(self.frame_detalle, text="🔔 ENVIAR RECORDATORIO DE VENCIMIENTO", state="disabled", fg_color="#f39c12", hover_color="#e67e22", text_color="black", height=45, font=("Segoe UI", 12, "bold"), command=self.enviar_aviso_vencimiento)
        self.btn_aviso_vencimiento.pack(pady=3, fill="x", padx=40)

        self.btn_aviso_mora = ctk.CTkButton(self.frame_detalle, text="⚠️ ENVIAR RECLAMO DE MORA", state="disabled", fg_color="#d35400", hover_color="#e67e22", text_color="white", height=45, font=("Segoe UI", 12, "bold"), command=self.enviar_aviso_mora)
        self.btn_aviso_mora.pack(pady=3, fill="x", padx=40)
        
        self.btn_imprimir = ctk.CTkButton(self.frame_detalle, text="🖨️ IMPRIMIR BOLETA PDF", state="disabled", fg_color="#c0392b", hover_color="#e74c3c", height=45, font=("Segoe UI", 12, "bold"), command=self.generar_remito_pdf)
        self.btn_imprimir.pack(pady=(3, 10), fill="x", padx=40)
        
        self.btn_nuevo = ctk.CTkButton(self.frame_derecho, text="+ NUEVO CLIENTE", fg_color="#2e7d32", height=50, command=self.ventana_agregar)
        self.btn_nuevo.grid(row=2, column=0, sticky="ew", pady=(15, 0))

        self.index_sel = None
        self.cuota_pura_sel = 0.0
        self.actualizar_lista()

    def calcular_mora(self, fila):
        hoy = datetime.now().date()
       # hoy = hoy + timedelta(days=45)  # <-- ¡SIMULADOR DE TIEMPO! Sumale los días que quieras
        try:
            f_carga = pd.to_datetime(fila['Fecha Carga']).date()
            frec = fila['Frecuencia']; pagas = int(fila['Cuotas Pagas']); totales = int(fila['Cuotas Totales'])
            if pagas >= totales: return f_carga, hoy, 0, 0
            if frec == "Semanal": proximo_venc = f_carga + timedelta(weeks=pagas); dias_ciclo = 7
            elif frec == "Quincenal": proximo_venc = f_carga + timedelta(days=pagas * 15); dias_ciclo = 15
            else: proximo_venc = f_carga + relativedelta(months=pagas); dias_ciclo = 30
            dias_atraso = (hoy - proximo_venc).days
            cuotas_debe = (dias_atraso // dias_ciclo) + 1 if dias_atraso >= 0 else 0
            cuotas_restantes = totales - pagas
            if cuotas_debe > cuotas_restantes: cuotas_debe = cuotas_restantes
            return f_carga, proximo_venc, max(0, dias_atraso), cuotas_debe
        except: return hoy, hoy, 0, 0

    def actualizar_lista(self):
        for w in self.scroll_clientes.winfo_children(): w.destroy()
        for w in self.scroll_morosos.winfo_children(): w.destroy()
        if self.df_datos.empty: return
        busq = self.entry_busqueda.get().lower()
        cap_total = 0.0; cl_act = 0; morosos = []
        
        for idx, c in self.df_datos.iterrows():
            f_ini, f_venc, atraso, debe = self.calcular_mora(c)
            saldo = float(c['Saldo'])
            if saldo > 0: cap_total += saldo; cl_act += 1
            if busq and (busq not in str(c['Nombre']).lower() and busq not in str(c['DNI'])): continue
            if saldo <= 0: color, est = "#454545", "FINALIZADO"
            elif atraso > 0: color, est = "#b22222", f"MORA: {atraso}d (Debe {debe})"; morosos.append((idx, c, atraso, est, color))
            else: color, est = "#2e7d32", "AL DÍA"
            
            ctk.CTkButton(self.scroll_clientes, text=f"{c['Nombre']} | {c['Producto']} | Saldo: {format_moneda_visual(saldo)}", fg_color=color, anchor="w", command=lambda i=idx: self.ver_detalle(i)).pack(pady=2, fill="x", padx=10)
        
        morosos.sort(key=lambda x: x[2], reverse=True)
        for idx, c, atr, est, col in morosos:
            ctk.CTkButton(self.scroll_morosos, text=f"{c['Nombre']} | {c['Producto']} | Saldo: {format_moneda_visual(c['Saldo'])}", fg_color=col, anchor="w", command=lambda i=idx: self.ver_detalle(i)).pack(pady=2, fill="x", padx=10)
        
        self.lbl_capital.configure(text=f"Capital en la Calle: {format_moneda_visual(cap_total)}")
        self.lbl_cant_clientes.configure(text=f"Clientes Activos: {cl_act}")

    def ver_detalle(self, index):
        self.index_sel = index
        c = self.df_datos.loc[index]
        f_ini, f_venc, atraso, debe = self.calcular_mora(c)
        texto = (f"👤 CLIENTE: {c['Nombre']}\n🆔 DNI: {c['DNI']} | 📞 TEL: {c['Telefono']}\n🏠 DOM: {c['Domicilio']}\n📦 PROD: {c['Producto']}\n"
                 f"──────────────────────────────\n📅 COMPRA: {f_ini.strftime('%d/%m/%Y')}\n📊 CUOTAS: {int(c['Cuotas Pagas'])} de {int(c['Cuotas Totales'])}\n"
                 f"💵 VALOR CUOTA: {format_moneda_visual(c['Valor Cuota'])}\n💰 SALDO RESTANTE: {format_moneda_visual(c['Saldo'])}\n"
                 f"──────────────────────────────\n🔔 PRÓX. VENC: {f_venc.strftime('%d/%m/%Y')}\n⚠️ ATRASO: {atraso} días\n📉 DEBE ACTUALMENTE: {debe} cuotas")
        self.lbl_info.configure(text=texto)
        
        saldo = float(c['Saldo'])
        self.btn_pagar.configure(state="normal" if saldo > 0 else "disabled")
        self.btn_imprimir.configure(state="normal")
        self.btn_whatsapp.configure(state="normal")
        self.btn_aviso_vencimiento.configure(state="normal" if saldo > 0 else "disabled")
        self.btn_aviso_mora.configure(state="normal" if atraso > 0 and saldo > 0 else "disabled")
        
        if atraso > 0 and saldo > 0:
            self.cuota_pura_sel = float(c['Valor Cuota'])
            self.lbl_calc_cliente.configure(text=f"Liquidando a: {c['Nombre']}", text_color="#e67e22")
            self.lbl_calc_detalle.configure(text=f"Cuota pura: {format_moneda_visual(self.cuota_pura_sel)}\nDías de atraso: {atraso}\nCuotas vencidas: {debe}")
            self.btn_guardar_punitivo.configure(state="normal")
            self.recalcular_total_mora()
        else:
            self.cuota_pura_sel = 0.0
            self.lbl_calc_cliente.configure(text="Seleccione un cliente moroso para calcular", text_color="#aaaaaa")
            self.lbl_calc_detalle.configure(text="Cuota pura: $ 0,00\nDías de atraso: 0")
            self.btn_guardar_punitivo.configure(state="disabled")
            self.lbl_total_mora.configure(text="TOTAL A COBRAR: $ 0,00")
            self.entry_punitorio.delete(0, "end")

    def recalcular_total_mora(self):
        try:
            texto_punitivo = self.entry_punitorio.get().strip()
            punitorio = float(texto_punitivo) if texto_punitivo else 0.0
            total = self.cuota_pura_sel + punitorio
            self.lbl_total_mora.configure(text=f"TOTAL A COBRAR: {format_moneda_visual(total)}")
        except:
            self.lbl_total_mora.configure(text="TOTAL A COBRAR: $ ---")

    def registrar_pago(self):
        if self.index_sel is None: return
        v_cuota = float(self.df_datos.at[self.index_sel, "Valor Cuota"])
        if messagebox.askyesno("Cobro", f"¿Registrar pago de cuota pura de {format_moneda_visual(v_cuota)}?"):
            self.df_datos.at[self.index_sel, "Saldo"] -= v_cuota
            self.df_datos.at[self.index_sel, "Cuotas Pagas"] += 1
            self.df_datos.to_excel(ARCHIVO_EXCEL, index=False)
            self.actualizar_lista(); self.ver_detalle(self.index_sel)
            if messagebox.askyesno("WhatsApp", "¿Desea enviar el comprobante por WhatsApp ahora?"):
                self.enviar_whatsapp()

    def registrar_pago_con_punitorio(self):
        if self.index_sel is None: return
        try:
            texto_punitivo = self.entry_punitorio.get().strip()
            punitorio = float(texto_punitivo) if texto_punitivo else 0.0
        except:
            messagebox.showerror("Error", "El monto del punitorio no es válido."); return
            
        v_cuota = float(self.df_datos.at[self.index_sel, "Valor Cuota"])
        total = v_cuota + punitorio
        
        if messagebox.askyesno("Cobro con Mora", f"¿Registrar pago total de {format_moneda_visual(total)}?\n(Cuota: {format_moneda_visual(v_cuota)} + Punitorio: {format_moneda_visual(punitorio)})"):
            self.df_datos.at[self.index_sel, "Saldo"] -= v_cuota
            self.df_datos.at[self.index_sel, "Cuotas Pagas"] += 1
            self.df_datos.to_excel(ARCHIVO_EXCEL, index=False)
            self.actualizar_lista(); self.ver_detalle(self.index_sel)
            
            messagebox.showinfo("Éxito", f"Se registró el cobro de la cuota.\nGanancia extra por punitorio: {format_moneda_visual(punitorio)}")
            
            if messagebox.askyesno("WhatsApp", "¿Desea enviar el comprobante por WhatsApp ahora?"):
                self.enviar_whatsapp_con_punitorio(punitorio, total)

    def obtener_telefono_formateado(self, c):
        tel = str(c['Telefono']).replace(" ", "").replace("-", "").replace(".0", "")
        if not tel.startswith("54") and len(tel) > 0: tel = "54" + tel
        return tel

    def enviar_whatsapp(self):
        if self.index_sel is None: return
        c = self.df_datos.loc[self.index_sel]
        tel = self.obtener_telefono_formateado(c)
        if len(tel) < 10: 
            messagebox.showwarning("WhatsApp", "El número de teléfono parece inválido."); return

        mensaje = (f"Hola *{c['Nombre']}*!\n\n"
                   f"Registramos tu pago de *{format_moneda_visual(c['Valor Cuota'])}* en *{NOMBRE_NEGOCIO}*.\n"
                   f"PRODUCTO: {c['Producto']}\n"
                   f"CUOTAS PAGAS: {int(c['Cuotas Pagas'])} de {int(c['Cuotas Totales'])}\n"
                   f"SALDO RESTANTE: *{format_moneda_visual(c['Saldo'])}*\n\n"
                   f"Muchas gracias por tu pago!")
        
        url = f"https://wa.me/{tel}?text={quote(mensaje)}"
        webbrowser.open(url)

    def enviar_whatsapp_con_punitorio(self, punitorio, total):
        if self.index_sel is None: return
        c = self.df_datos.loc[self.index_sel]
        tel = self.obtener_telefono_formateado(c)
        if len(tel) < 10: return
        
        mensaje = (f"Hola *{c['Nombre']}*!\n\n"
                   f"Registramos tu pago por un total de *{format_moneda_visual(total)}* en *{NOMBRE_NEGOCIO}*.\n"
                   f"Detalle: Cuota pura ({format_moneda_visual(c['Valor Cuota'])}) + Recargo por mora ({format_moneda_visual(punitorio)})\n"
                   f"PRODUCTO: {c['Producto']}\n"
                   f"CUOTAS PAGAS: {int(c['Cuotas Pagas'])} de {int(c['Cuotas Totales'])}\n"
                   f"SALDO RESTANTE DE LA COMPRA: *{format_moneda_visual(c['Saldo'])}*\n\n"
                   f"Muchas gracias por tu pago!")
        
        url = f"https://wa.me/{tel}?text={quote(mensaje)}"
        webbrowser.open(url)

    def enviar_aviso_vencimiento(self):
        if self.index_sel is None: return
        c = self.df_datos.loc[self.index_sel]
        f_ini, f_venc, atraso, debe = self.calcular_mora(c)
        tel = self.obtener_telefono_formateado(c)
        
        if len(tel) < 10: 
            messagebox.showwarning("WhatsApp", "El número de teléfono parece inválido."); return

        mensaje = (f"Hola *{c['Nombre']}*!\n\n"
                   f"Te recordamos de *{NOMBRE_NEGOCIO}* que se aproxima el vencimiento de tu cuota.\n"
                   f"FECHA DE VENCIMIENTO: *{f_venc.strftime('%d/%m/%Y')}*\n"
                   f"VALOR DE LA CUOTA: *{format_moneda_visual(c['Valor Cuota'])}*\n"
                   f"PRODUCTO: {c['Producto']}\n\n"
                   f"Que tengas un excelente día!")
        
        url = f"https://wa.me/{tel}?text={quote(mensaje)}"
        webbrowser.open(url)

    def enviar_aviso_mora(self):
        if self.index_sel is None: return
        c = self.df_datos.loc[self.index_sel]
        f_ini, f_venc, atraso, debe = self.calcular_mora(c)
        tel = self.obtener_telefono_formateado(c)
        
        if len(tel) < 10: 
            messagebox.showwarning("WhatsApp", "El número de teléfono parece inválido."); return

        mensaje = (f"Hola *{c['Nombre']}*! ATENCION\n\n"
                   f"Nos comunicamos de *{NOMBRE_NEGOCIO}* porque registramos un atraso en tu cuenta.\n"
                   f"DIAS DE ATRASO: *{atraso}*\n"
                   f"CUOTAS ADEUDADAS: *{debe}*\n\n"
                   f"Por favor, ponete en contacto con nosotros a la brevedad para conocer el monto a abonar con los recargos correspondientes y regularizar tu situación. Muchas gracias!")
        
        url = f"https://wa.me/{tel}?text={quote(mensaje)}"
        webbrowser.open(url)

    def generar_remito_pdf(self):
        if self.index_sel is None: return
        c = self.df_datos.loc[self.index_sel]
        f_ini, f_venc, atraso, debe = self.calcular_mora(c)
        
        carpeta_boletas = "Boletas"
        if not os.path.exists(carpeta_boletas): os.makedirs(carpeta_boletas)
        
        nombre_limpio = "".join(x for x in str(c['Nombre']) if x.isalnum() or x in "._- ")
        nombre_archivo = f"Remito_{c['DNI']}_{nombre_limpio.replace(' ', '_')}.pdf"
        ruta_completa = os.path.join(carpeta_boletas, nombre_archivo)
        
        doc = SimpleDocTemplate(ruta_completa, pagesize=letter,
                                rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=2*cm, bottomMargin=2*cm)
        story = []
        styles = getSampleStyleSheet()
        style_normal = styles['Normal']
        
        # --- NUEVOS ESTILOS PARA ADAPTARSE AL CARTEL ---
        style_contacto = ParagraphStyle('Contacto', parent=style_normal, fontName='Helvetica', fontSize=10, leading=14, alignment=2)
        style_subtitulo = ParagraphStyle('Sub', parent=style_normal, fontName='Helvetica-Bold', fontSize=11, leading=14)
        style_garantia = ParagraphStyle('Garantia', parent=style_normal, fontName='Helvetica-Oblique', fontSize=9, leading=12, alignment=4)
        style_mora = ParagraphStyle('Mora', parent=style_normal, fontName='Helvetica-Oblique', fontSize=9, leading=11, alignment=4)

        # 🚨 SECCIÓN LOGO: Intenta cargar la imagen en un costado
        ruta_logo = "logo.png"
        bloque_izq = []
        
        if os.path.exists(ruta_logo):
            # Carga la imagen guardada. Ajustamos el tamaño proporcionalmente.
            img_logo = Image(ruta_logo, width=7.5*cm, height=3.5*cm)
            img_logo.hAlign = 'LEFT'
            bloque_izq.append(img_logo)
        else:
            # Si no existe la imagen todavía, deja el espacio o usa texto de respaldo
            bloque_izq.append(Paragraph(f"<b>{NOMBRE_NEGOCIO}</b>", ParagraphStyle('Bck', fontName='Helvetica-Bold', fontSize=28)))
            bloque_izq.append(Spacer(1, 0.2*cm))
            bloque_izq.append(Paragraph(f"{RUBRO_NEGOCIO}", ParagraphStyle('BckRub', fontName='Helvetica-Bold', fontSize=9)))

        # Bloque Derecho: Dirección + Teléfono
        texto_der = Paragraph(f"<b>{DIRECCION_NEGOCIO}</b><br/>Cel: {TELEFONO_NEGOCIO}<br/>Fecha: {datetime.now().strftime('%d/%m/%Y')}", style_contacto)

        datos_cabecera = [[bloque_izq, texto_der]]
        tabla_cabecera = Table(datos_cabecera, colWidths=[11.5*cm, 6.5*cm])
        tabla_cabecera.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), # Se cambia a middle para que la imagen y el texto queden parejos
            ('ALIGN', (1,0), (1,0), 'RIGHT')
        ]))
        
        story.append(tabla_cabecera)
        story.append(Spacer(1, 0.8*cm))
        
        story.append(Paragraph("<b>X</b>", ParagraphStyle('Centro', parent=style_normal, alignment=1, fontSize=18)))
        story.append(Spacer(1, 0.5*cm))
        
        f_carga = pd.to_datetime(c['Fecha Carga'])
        datos_cliente = [
            [Paragraph("<b>Cliente:</b>", style_subtitulo), Paragraph(str(c['Nombre']).upper(), style_normal), Paragraph("<b>DNI:</b>", style_subtitulo), Paragraph(str(c['DNI']), style_normal)],
            [Paragraph("<b>Domicilio:</b>", style_subtitulo), Paragraph(str(c['Domicilio']).upper(), style_normal), Paragraph("<b>Teléfono:</b>", style_subtitulo), Paragraph(str(c['Telefono']), style_normal)]
        ]
        
        tabla_cliente = Table(datos_cliente, colWidths=[2.5*cm, 6.5*cm, 2.5*cm, 6.5*cm])
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
            [str(c['Producto']).upper(), format_moneda_visual(c['Precio Total']), str(int(c['Cuotas Totales'])), format_moneda_visual(c['Valor Cuota'])]
        ]
        
        tabla_producto = Table(datos_producto, colWidths=[9*cm, 3*cm, 3*cm, 3*cm])
        tabla_producto.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#555555")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,0), 'CENTER'),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('PADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(tabla_producto)
        story.append(Spacer(1, 0.8*cm))
        
        story.append(Paragraph("PLAN DE PAGOS Y VENCIMIENTOS", style_subtitulo))
        story.append(Spacer(1, 0.2*cm))
        
        datos_vencimientos = [["N° Cuota", "Fecha de Vencimiento", "Monto", "Estado"]]
        frec = c['Frecuencia']
        for i in range(1, int(c['Cuotas Totales']) + 1):
            if i == 1: f_venc_c = f_carga
            else:
                if frec == "Semanal": f_venc_c = f_carga + timedelta(weeks=i-1)
                elif frec == "Quincenal": f_venc_c = f_carga + timedelta(days=(i-1) * 15)
                else: f_venc_c = f_carga + relativedelta(months=i-1)
            
            estado = "PAGADA" if i <= int(c['Cuotas Pagas']) else "PENDIENTE"
            datos_vencimientos.append([f"Cuota {i}", f_venc_c.strftime('%d/%m/%Y'), format_moneda_visual(c['Valor Cuota']), estado])
            
        tabla_vencimientos = Table(datos_vencimientos, colWidths=[4.5*cm, 5*cm, 4.5*cm, 4*cm])
        tabla_vencimientos.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#777777")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('PADDING', (0,0), (-1,-1), 4),
            *[('BACKGROUND', (0, i+1), (-1, i+1), colors.HexColor("#e2f0d9")) for i in range(int(c['Cuotas Pagas']))],
        ]))
        story.append(tabla_vencimientos)
        story.append(Spacer(1, 0.6*cm))
        
        clausula = ("<b>CLÁUSULA DE GARANTÍA:</b> stuar´s garantiza el correcto funcionamiento del producto "
                    "por desperfectos de fabricación por un período de 30 días a partir de la fecha de entrega. "
                    "La garantía quedará anulada si el producto presenta golpes, roturas, enmiendas, o uso indebido. "
                    f"Transcurridos los {DIAS_MORA_PARA_PUNITORIO} días de atraso, se generarán recargos punitorios.")
        story.append(Paragraph(clausula, ParagraphStyle('G', parent=style_garantia, fontSize=8, leading=10)))
        
        story.append(Spacer(1, 0.4*cm)) 
        
        texto_mora_imagen = ("Se entiende que en caso de incumplimiento con el pago acordado, "
                             "el <b>GARANTE</b> debe asumir inmediatamente el cargo de dichos pagos adeudados "
                             "con sus recargos por mora o bien devolver el producto hasta la cancelación total "
                             "de dicho artículo.")
        
        story.append(Paragraph(texto_mora_imagen, style_mora))
        story.append(Spacer(1, 2.2*cm))
        
        datos_firmas = [["..................................................", ".................................................."],
                        ["Firma y Aclaración del Cliente", "Firma Autorizada Stuar´s"]]
        tabla_firmas = Table(datos_firmas, colWidths=[9*cm, 9*cm])
        tabla_firmas.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,1), (-1,1), 10)
        ]))
        story.append(tabla_firmas)
        
        doc.build(story)
        
        try:
            if platform.system() == "Windows": os.startfile(ruta_completa)
            else: os.system(f"open '{ruta_completa}'")
        except: pass
            
        messagebox.showinfo("Éxito", f"Se ha generado la boleta en la carpeta Boletas:\n{nombre_archivo}")

    def ventana_agregar(self):
        vent = ctk.CTkToplevel(self)
        vent.geometry("500x850")
        vent.title("Nuevo Crédito")
        vent.attributes("-topmost", True)
        
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
        
        def formatear_monto_en_vivo(event):
            texto = entradas["Precio"].get()
            solo_numeros = "".join([c for c in texto if c.isdigit()])
            if solo_numeros:
                numero = int(solo_numeros)
                formateado = f"{numero:,}".replace(",", ".")
                entradas["Precio"].delete(0, "end")
                entradas["Precio"].insert(0, formateado)
            recalcular()

        def recalcular(*args):
            try: 
                p = float(limpiar_monto(entradas["Precio"].get()))
                c = int(entradas["Cuotas"].get())
                v_c = p / c
                lbl_v.configure(text=f"CUOTA: {format_moneda_visual(v_c)}")
            except: 
                lbl_v.configure(text="CUOTA: $ ---")
                
        entradas["Precio"].bind("<KeyRelease>", formatear_monto_en_vivo)
        entradas["Cuotas"].bind("<KeyRelease>", recalcular)
        
        def guardar():
            for c in entradas:
                if not entradas[c].get().strip(): 
                    return messagebox.showwarning("Falta", f"Completar {c}")
            
            try:
                p = float(limpiar_monto(entradas["Precio"].get()))
                c = int(entradas["Cuotas"].get())
                v_c = p / c
            except ValueError:
                return messagebox.showerror("Error", "El precio o las cuotas no son números válidos.")
            
            nueva_fila = {
                "DNI": str(entradas["DNI"].get()), 
                "Nombre": entradas["Nombre"].get().upper(), 
                "Domicilio": entradas["Domicilio"].get().upper(), 
                "Telefono": str(entradas["Telefono"].get()), 
                "Producto": entradas["Producto"].get().upper(), 
                "Precio Total": p, 
                "Cuotas Totales": c, 
                "Valor Cuota": v_c, 
                "Cuotas Pagas": 1, 
                "Saldo": p - v_c, 
                "Frecuencia": frec_v.get(), 
                "Fecha Carga": datetime.now().strftime("%Y-%m-%d")
            }
            
            self.df_datos = pd.concat([self.df_datos, pd.DataFrame([nueva_fila])], ignore_index=True)
            self.df_datos.to_excel(ARCHIVO_EXCEL, index=False)
            
            self.actualizar_lista()
            vent.destroy()
            messagebox.showinfo("Éxito", "Venta guardada.")
            
        ctk.CTkButton(sf, text="GUARDAR VENTA", fg_color="#2e7d32", height=50, command=guardar).pack(pady=20)

if __name__ == "__main__":
    app = App()
    app.mainloop()