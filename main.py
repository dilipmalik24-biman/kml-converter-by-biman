# --- Pydroid 3 Multi-Layer GUI Converter with Fuzzy Photo Matching & Base64 HD Engine ---
import os
import sys
import zipfile
import math
import base64
import csv
import re
from html import escape
from io import BytesIO

os.environ["KIVY_NO_ARGS"] = "1"
os.environ["KIVY_LOG_MODE"] = "MIXED"

from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.label import MDLabel
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.textfield import MDTextField
from kivymd.uix.card import MDCard
from kivymd.uix.filemanager import MDFileManager
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock
from kivy.utils import platform
import openpyxl

try:
    from PIL import Image
except ImportError:
    Image = None

def utm_to_latlon(easting, northing, zone):
    a = 6378137.0         
    f = 1.0 / 298.257223563 
    k0 = 0.9996           
    b = a * (1.0 - f)
    e2 = (a**2 - b**2) / (a**2)
    ep2 = (a**2 - b**2) / (b**2)
    lon0 = ((zone * 6.0) - 183.0) * (math.pi / 180.0)
    x = easting - 500000.0
    y = northing
    M = y / k0
    mu = M / (a * (1.0 - e2 / 4.0 - 3.0 * e2**2 / 64.0 - 5.0 * e2**3 / 256.0))
    e1 = (1.0 - math.sqrt(1.0 - e2)) / (1.0 + math.sqrt(1.0 - e2))
    phi1 = (mu + (3.0 * e1 / 2.0 - 27.0 * e1**3 / 32.0) * math.sin(2.0 * mu)
            + (21.0 * e1**2 / 16.0 - 55.0 * e1**4 / 32.0) * math.sin(4.0 * mu)
            + (151.0 * e1**3 / 96.0) * math.sin(6.0 * mu))
    sin_phi1 = math.sin(phi1)
    cos_phi1 = math.cos(phi1)
    tan_phi1 = math.tan(phi1)
    N1 = a / math.sqrt(1.0 - e2 * sin_phi1**2)
    R1 = a * (1.0 - e2) / (1.0 - e2 * math.sin(phi1)**2)**1.5
    D = x / (N1 * k0)
    lat = phi1 - (N1 * tan_phi1 / R1) * (D**2 / 2.0 - (5.0 + 3.0 * tan_phi1**2 + 10.0 * (ep2 * cos_phi1**2) - 4.0 * (ep2 * cos_phi1**2)**2 - 9.0 * ep2) * D**4 / 24.0 + (61.0 + 90.0 * tan_phi1**2 + 298.0 * (ep2 * cos_phi1**2) + 45.0 * tan_phi1**4 - 252.0 * ep2 - 3.0 * (ep2 * cos_phi1**2)**2) * D**6 / 720.0)
    lon = lon0 + (D - (1.0 + 2.0 * tan_phi1**2 + (ep2 * cos_phi1**2)) * D**3 / 6.0 + (5.0 - 2.0 * (ep2 * cos_phi1**2) + 28.0 * tan_phi1**2 - 3.0 * (ep2 * cos_phi1**2)**2 + 8.0 * ep2 + 24.0 * tan_phi1**4) * D**5 / 120.0) / cos_phi1
    return math.degrees(lat), math.degrees(lon)

def sanitize_and_clean_excel(file_path):
    """Excel Filter XML Stripper to prevent openpyxl crashes"""
    try:
        cleaned_zip_buffer = BytesIO()
        with zipfile.ZipFile(file_path, 'r') as source_zip:
            with zipfile.ZipFile(cleaned_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as target_zip:
                for item in source_zip.infolist():
                    data = source_zip.read(item.filename)
                    if "xl/worksheets/sheet" in item.filename.lower():
                        xml_content = data.decode('utf-8', errors='ignore')
                        xml_content = re.sub(r'<autoFilter[^>]*>.*?</autoFilter>', '', xml_content, flags=re.DOTALL)
                        xml_content = re.sub(r'<autoFilter[^>]*/>', '', xml_content)
                        data = xml_content.encode('utf-8')
                    target_zip.writestr(item, data)
        cleaned_zip_buffer.seek(0)
        return cleaned_zip_buffer
    except Exception:
        return file_path

def find_photo_file(base_dir, raw_filename):
    """Fuzzy matching to automatically find photos with/without extension"""
    if not raw_filename or str(raw_filename).lower() == "nan":
        return None

    raw_filename = str(raw_filename).strip()
    
    # 1. Direct Exact Match
    target_path = os.path.join(base_dir, raw_filename)
    if os.path.exists(target_path) and os.path.isfile(target_path):
        return target_path

    # 2. Fuzzy Match inside directory
    try:
        folder_files = os.listdir(base_dir)
        raw_base_name = os.path.splitext(raw_filename)[0].lower()
        
        for file in folder_files:
            file_base_name = os.path.splitext(file)[0].lower()
            if file_base_name == raw_base_name or file.lower().startswith(raw_base_name):
                matched_path = os.path.join(base_dir, file)
                if os.path.isfile(matched_path) and not file.endswith(('.xlsx', '.csv', '.kmz', '.kml', '.py')):
                    return matched_path
    except Exception:
        pass

    return None

class GISMasterApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.selected_file = None
        self.mode = "WGS84"
        self.sel_x = self.sel_y = self.sel_label = self.sel_photo = None
        self.icon_url = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"
        self.i_color, self.l_color = "ff0000ff", "ffffffff"
        self.menus = {}
        self.columns = []
        self.data_rows = []
        
        self.kml_layers_repository = []
        self.layer_counter = 0

        self.internal_file_manager = MDFileManager(
            exit_manager=self.exit_internal_manager,
            select_path=self.on_file_selected
        )

    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Teal"
        screen = MDScreen()
        scroll = ScrollView()
        layout = MDBoxLayout(orientation="vertical", padding=15, spacing=12, size_hint_y=None)
        layout.bind(minimum_height=layout.setter("height"))

        header = MDCard(size_hint=(1, None), height="85dp", radius=[12,], md_bg_color=(0, 0.4, 0.4, 1), padding=12)
        header.add_widget(MDLabel(text="GIS MULTI-LAYER CONVERTER\nFuzzy Extension & Base64 HD Engine", halign="center", theme_text_color="Custom", text_color=(1, 1, 1, 1), font_style="H6"))
        layout.add_widget(header)

        self.btn_file = MDRaisedButton(text="1. LOAD DATA FILE (CSV/XLSX)", size_hint=(1, None), on_release=self.trigger_file_selection)
        self.btn_mode = MDRaisedButton(text="2. INPUT CRS SYSTEM", size_hint=(1, None), md_bg_color=(0.1, 0.4, 0.6, 1))
        self.btn_x = MDRaisedButton(text="SELECT X (EASTING / LON)", size_hint=(1, None), disabled=True)
        self.btn_y = MDRaisedButton(text="SELECT Y (NORTHING / LAT)", size_hint=(1, None), disabled=True)
        self.btn_label = MDRaisedButton(text="SELECT POINT LABEL", size_hint=(1, None), disabled=True)
        self.btn_photo = MDRaisedButton(text="SELECT PHOTO COLUMN", size_hint=(1, None), disabled=True)
        self.btn_icon = MDRaisedButton(text="SELECT ICON SHAPE", size_hint=(1, None), md_bg_color=(0.4, 0.2, 0.6, 1))
        self.btn_icon_col = MDRaisedButton(text="SELECT ICON COLOR", size_hint=(1, None), md_bg_color=(1, 0, 0, 1))
        self.btn_label_col = MDRaisedButton(text="SELECT LABEL COLOR", size_hint=(1, None), md_bg_color=(1, 1, 1, 1), text_color=(0, 0, 0, 1))

        self.i_scale = MDTextField(hint_text="Icon Display Scale (e.g. 1.0)", text="1.0", size_hint=(1, None))
        self.l_scale = MDTextField(hint_text="Label Display Scale (e.g. 0.9)", text="0.9", size_hint=(1, None))
        
        self.btn_add_layer = MDRaisedButton(text="➕ ADD THIS AS A LAYER", md_bg_color=(0.5, 0.4, 0.1, 1), size_hint=(1, None), on_release=self.pack_current_layer)
        
        self.out_name = MDTextField(hint_text="Output File Name (e.g. Total_PSS_Grid)", size_hint=(1, None))
        self.btn_gen = MDRaisedButton(text="🚀 GENERATE MULTI-LAYER HD KMZ", md_bg_color=(0, 0.6, 0.3, 1), size_hint=(1, None), on_release=self.compile_all_layers)

        for w in [self.btn_file, self.btn_mode, self.btn_x, self.btn_y, self.btn_label, self.btn_photo, self.btn_icon, self.btn_icon_col, self.btn_label_col, self.i_scale, self.l_scale, self.btn_add_layer, self.out_name, self.btn_gen]:
            layout.add_widget(w)

        self.status_card = MDCard(size_hint=(1, None), height="150dp", padding=15, radius=[12,], md_bg_color=(0.12, 0.12, 0.12, 1))
        self.status = MDLabel(text="STATUS: NO LAYERS ADDED YET\n\nFile load karein, variables map karein aur 'Add This As A Layer' dabayein.", halign="center", theme_text_color="Primary", font_style="Body1")
        self.status_card.add_widget(self.status)
        layout.add_widget(self.status_card)

        self.setup_menus()
        scroll.add_widget(layout)
        screen.add_widget(scroll)
        return screen

    def setup_menus(self):
        m_list = [("WGS84 (Lat/Long)", "WGS84"), ("UTM 42N", 42), ("UTM 43N", 43)]
        self.menus["mode"] = MDDropdownMenu(caller=self.btn_mode, items=[{"text": x[0], "on_release": lambda v=x: self.set_mode(v)} for x in m_list], width_mult=4)
        self.btn_mode.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["mode"].open(), 0.05))

        s_list = [("Circle", "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"), ("Square", "http://maps.google.com/mapfiles/kml/shapes/placemark_square.png"), ("Star", "http://maps.google.com/mapfiles/kml/shapes/star.png"), ("Triangle", "http://maps.google.com/mapfiles/kml/shapes/triangle.png"), ("Diamond", "http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png")]
        self.menus["icon"] = MDDropdownMenu(caller=self.btn_icon, items=[{"text": x[0], "on_release": lambda v=x: self.set_icon(v)} for x in s_list], width_mult=4)
        self.btn_icon.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["icon"].open(), 0.05))

        c_list = [("Red", "ff0000ff", (1, 0, 0, 1)), ("Green", "ff00ff00", (0, 1, 0, 1)), ("Blue", "ffff0000", (0, 0, 1, 1)), ("Yellow", "ff00ffff", (1, 1, 0, 1)), ("White", "ffffffff", (1, 1, 1, 1))]
        self.menus["i_col"] = MDDropdownMenu(caller=self.btn_icon_col, items=[{"text": x[0], "on_release": lambda v=x: self.set_color("i", v)} for x in c_list], width_mult=4)
        self.menus["l_col"] = MDDropdownMenu(caller=self.btn_label_col, items=[{"text": x[0], "on_release": lambda v=x: self.set_color("l", v)} for x in c_list], width_mult=4)
        self.btn_icon_col.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["i_col"].open(), 0.05))
        self.btn_label_col.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["l_col"].open(), 0.05))

    def set_mode(self, m):
        self.mode = m[1]
        self.btn_mode.text = f"Mode: {m[0]}"
        self.menus["mode"].dismiss()

    def set_icon(self, s):
        self.icon_url = s[1]
        self.btn_icon.text = f"Shape: {s[0]}"
        self.menus["icon"].dismiss()

    def set_color(self, t, c):
        if t == "i":
            self.i_color = c[1]
            self.btn_icon_col.md_bg_color = c[2]
            self.menus["i_col"].dismiss()
        else:
            self.l_color = c[1]
            self.btn_label_col.md_bg_color = c[2]
            self.menus["l_col"].dismiss()

    def trigger_file_selection(self, *args):
        initial_path = "/storage/emulated/0"
        self.internal_file_manager.show(initial_path)

    def exit_internal_manager(self, *args):
        self.internal_file_manager.close()

    def on_file_selected(self, path):
        self.exit_internal_manager()
        if not path or not os.path.isfile(path):
            return
            
        self.selected_file = path
        self.columns = []
        self.data_rows = []

        if path.endswith(".xlsx"):
            excel_src = sanitize_and_clean_excel(path)
            wb = openpyxl.load_workbook(excel_src, data_only=True)
            sheet = wb.active
            for r_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                if r_idx == 0:
                    self.columns = [str(c).strip() for c in row if c is not None]
                else:
                    if any(x is not None for x in row):
                        self.data_rows.append(list(row))
        else:
            with open(path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                self.columns = [c.strip() for c in next(reader)]
                for row in reader:
                    if row: self.data_rows.append(row)

        self.menus["x"] = MDDropdownMenu(caller=self.btn_x, items=[{"text": c, "on_release": lambda x=c: self.set_val("x", x)} for c in self.columns], width_mult=4)
        self.menus["y"] = MDDropdownMenu(caller=self.btn_y, items=[{"text": c, "on_release": lambda x=c: self.set_val("y", x)} for c in self.columns], width_mult=4)
        self.menus["label"] = MDDropdownMenu(caller=self.btn_label, items=[{"text": c, "on_release": lambda x=c: self.set_val("l", x)} for c in self.columns], width_mult=4)
        self.menus["photo"] = MDDropdownMenu(caller=self.btn_photo, items=[{"text": c, "on_release": lambda x=c: self.set_val("p", x)} for c in self.columns], width_mult=4)

        self.btn_x.disabled = self.btn_y.disabled = self.btn_label.disabled = self.btn_photo.disabled = False
        self.btn_x.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["x"].open(), 0.05))
        self.btn_y.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["y"].open(), 0.05))
        self.btn_label.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["label"].open(), 0.05))
        self.btn_photo.bind(on_release=lambda x: Clock.schedule_once(lambda dt: self.menus["photo"].open(), 0.05))

        # Intelligent auto-mapping defaults
        for col in self.columns:
            cl = col.lower().strip()
            if 'lon' in cl or 'easting' in cl or cl == 'x': self.set_val("x", col)
            if 'lat' in cl or 'northing' in cl or cl == 'y': self.set_val("y", col)
            if 'name' in cl or 'label' in cl or 'unique' in cl or 'id' in cl or 'pole_no' in cl: self.set_val("l", col)
            if 'local photo' in cl or 'photo' in cl or 'image' in cl or 'pic' in cl: self.set_val("p", col)

        self.btn_file.text = f"LOADED: {os.path.basename(path)}"

    def set_val(self, t, v):
        if t == "x":
            self.sel_x = v
            self.btn_x.text = f"X: {v}"
            self.menus["x"].dismiss()
        elif t == "y":
            self.sel_y = v
            self.btn_y.text = f"Y: {v}"
            self.menus["y"].dismiss()
        elif t == "l":
            self.sel_label = v
            self.btn_label.text = f"Label: {v}"
            self.menus["label"].dismiss()
        elif t == "p":
            self.sel_photo = v
            self.btn_photo.text = f"Photo Col: {v}"
            self.menus["photo"].dismiss()

    def pack_current_layer(self, *args):
        if not self.selected_file or not self.sel_x or not self.sel_y or not self.sel_label or not self.sel_photo:
            self.status.text = "LAYER ERROR:\nPehle file chunhein aur variables map karein!"
            return

        if self.sel_x not in self.columns or self.sel_y not in self.columns or self.sel_label not in self.columns or self.sel_photo not in self.columns:
            self.status.text = "MAPPING ERROR:\nChuna hua column is file me nahi mila!"
            return

        layer_name = os.path.splitext(os.path.basename(self.selected_file))[0]
        self.layer_counter += 1
        style_id = f"style_layer_{self.layer_counter}"
        
        style_xml = f"""    <Style id="{style_id}">
      <IconStyle>
        <color>{self.i_color}</color>
        <scale>{float(self.i_scale.text)}</scale>
        <Icon><href>{self.icon_url}</href></Icon>
      </IconStyle>
      <LabelStyle>
        <color>{self.l_color}</color>
        <scale>{float(self.l_scale.text)}</scale>
      </LabelStyle>
    </Style>"""

        x_idx = self.columns.index(self.sel_x)
        y_idx = self.columns.index(self.sel_y)
        l_idx = self.columns.index(self.sel_label)
        p_idx = self.columns.index(self.sel_photo)
        base_dir = os.path.dirname(self.selected_file)

        placemarks_xml = ""
        points_count = 0

        for row in self.data_rows:
            try:
                while len(row) < len(self.columns): row.append("")
                val_x = str(row[x_idx]).strip()
                val_y = str(row[y_idx]).strip()
                val_lbl = str(row[l_idx]).strip()

                if not val_x or not val_y or val_x.lower() == "nan" or val_y.lower() == "nan":
                    continue

                vx, vy = float(val_x), float(val_y)
                lon, lat = (vx, vy) if self.mode == "WGS84" else utm_to_latlon(vx, vy, self.mode)

                desc = (
                    '<div style="font-family:\'Segoe UI\',Arial; width:300px; background:#fdfdfd; '
                    'padding:0; border-radius:8px; overflow:hidden; border:1px solid #ccc;">'
                )
                desc += f'<div style="background:#005A5B; color:white; padding:10px; font-weight:bold;">Asset ID: {val_lbl}</div>'
                desc += '<table style="width:100%; border-collapse:collapse; font-size:12px;">'

                img_html = ""
                for idx, col in enumerate(self.columns):
                    val = str(row[idx]).strip()
                    if not val or val.lower() == "nan" or idx == x_idx or idx == y_idx: continue
                    
                    if idx == p_idx:
                        # FUZZY AUTO MATCH LOGIC
                        resolved_img_path = find_photo_file(base_dir, val)
                        
                        if resolved_img_path:
                            if Image is not None:
                                img = Image.open(resolved_img_path)
                                img.thumbnail((900, 900))
                                buf = BytesIO()
                                img.save(buf, format='JPEG', quality=85, optimize=True)
                                b_data = buf.getvalue()
                            else:
                                with open(resolved_img_path, "rb") as f: b_data = f.read()
                            
                            b64_str = base64.b64encode(b_data).decode('utf-8')
                            img_html = f'<br/><div style="text-align:center; padding:10px;"><img src="data:image/jpeg;base64,{b64_str}" width="100%"/></div>'
                        else:
                            img_html = f'<br/><i style="color:red; padding:10px; display:block; text-align:center;">[Local File Not Found: {val}]</i>'
                    else:
                        desc += f'<tr><td style="padding:6px; font-weight:bold; background:#f2f2f2; border:1px solid #ddd;">{col}</td><td style="padding:6px; border:1px solid #ddd;">{escape(val)}</td></tr>'

                full_desc = desc + "</table>" + img_html + "</div>"
                
                placemarks_xml += f"""      <Placemark>
        <name>{escape(val_lbl)}</name>
        <styleUrl>#{style_id}</styleUrl>
        <description><![CDATA[{full_desc}]]></description>
        <Point><coordinates>{lon},{lat},0</coordinates></Point>
      </Placemark>
"""
                points_count += 1
            except Exception:
                continue

        layer_xml = f"""  <Folder>
    <name>{escape(layer_name)} ({points_count} Points)</name>
{style_xml}
{placemarks_xml}  </Folder>
"""
        self.kml_layers_repository.append(layer_xml)
        
        self.status.text = f"SUCCESS: Layer '{layer_name}' added!\nTotal layers in memory: {len(self.kml_layers_repository)}\n\nAb aap doosri file load kar sakte hain."
        
        # Reset current context for next layer load
        self.selected_file = None
        self.sel_x = self.sel_y = self.sel_label = self.sel_photo = None
        self.columns = []
        self.data_rows = []
        
        self.btn_x.disabled = self.btn_y.disabled = self.btn_label.disabled = self.btn_photo.disabled = True
        self.btn_x.text = "SELECT X (EASTING / LON)"
        self.btn_y.text = "SELECT Y (NORTHING / LAT)"
        self.btn_label.text = "SELECT POINT LABEL"
        self.btn_photo.text = "SELECT PHOTO COLUMN"
        self.btn_file.text = "1. LOAD NEXT DATA FILE"

    def compile_all_layers(self, *args):
        if not self.kml_layers_repository:
            self.status.text = "COMPILE ERROR:\nPehle kam se kam ek layer add karein!"
            return

        out_filename = self.out_name.text.strip() or "Multi_Layer_PSS_Output"
        if not out_filename.endswith(".kmz"): out_filename += ".kmz"

        if platform == 'android':
            from android.storage import primary_external_storage_path
            save_path = os.path.join(primary_external_storage_path(), "Download", out_filename)
        else:
            save_path = out_filename

        final_xml_payload = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>{escape(out_filename.replace('.kmz',''))}</name>
{"".join(self.kml_layers_repository)}  </Document>
</kml>"""

        with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as kmz:
            kmz.writestr("doc.kml", final_xml_payload)

        self.status.text = (
            f"CONVERSION COMPLETED SUCCESSFULLY!\n"
            f"----------------------------------------\n"
            f"Total Layers Packaged: {len(self.kml_layers_repository)}\n"
            f"Saved to: Android Download Directory\n"
            f"File: {out_filename}"
        )
        
        self.kml_layers_repository = []
        self.layer_counter = 0
        self.btn_gen.disabled = False

if __name__ == "__main__":
    GISMasterApp().run()
