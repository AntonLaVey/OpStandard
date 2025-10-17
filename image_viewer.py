def load_file(self, path, page, cache_key):
        try:
            if path.lower().endswith(".xlsx"):
                sheet_type = page.lower() if page != "Image" else "front"
                sheet = self.excel_converter.find_sheet(path, sheet_type)
                if not sheet:
                    self.root.after(0, lambda: self.image_label.config(image="", text=f"Sheet not found: {page}"))
                    return
                png_path = self.excel_converter.convert_excel_to_png(path, sheet)
                if not png_path:
                    self.root.after(0, lambda: self.image_label.config(image="", text=f"Conversion failed"))
                    return
                img = Image.open(png_path)
            else:
                img = Image.open(path)
            
            w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight() - self.control_bar_collapsed_height
            img.thumbnail((min(w, MAX_IMAGE_DIMENSION), min(h, MAX_IMAGE_DIMENSION)), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.image_cache.put(cache_key, photo)
            self.root.after(0, lambda: (self.image_label.config(image=photo, text=""),
                                       setattr(self.image_label, 'image', photo)))
        except Exception as e:
            logger.error(f"Load error: {e}")
            self.root.after(0, lambda: self.image_label.config(image="", text=f"Error: {os.path.basename(path)}"))
    
    def precache_model_aggressive(self, model_path):
        """Aggressively precache current model"""
        logger.info(f"Aggressive precache: {os.path.basename(model_path)}")
        if not model_path or not os.path.exists(model_path):
            return
        try:
            for f in sorted(os.listdir(model_path)):
                if f.lower().endswith(".xlsx"):
                    file_path = os.path.join(model_path, f)
                    for sheet_type in ["front", "back"]:
                        sheet = self.excel_converter.find_sheet(file_path, sheet_type)
                        if sheet:
                            logger.info(f"Aggressive cache: {os.path.basename(file_path)} - {sheet_type}")
                            self.excel_converter.convert_excel_to_png(file_path, sheet)
        except Exception as e:
            logger.error(f"Aggressive precache error: {e}")
    
    def precache_dept(self, dept):
        """Slowly precache entire department in background"""
        logger.info(f"Background precaching department: {dept}")
        dept_path = os.path.join(NETWORK_BASE_PATH, dept)
        try:
            models = sorted([d for d in os.listdir(dept_path) if os.path.isdir(os.path.join(dept_path, d))])
            for model in models:
                if self.stop_precache.is_set():
                    return
                model_path = os.path.join(dept_path, model)
                logger.info(f"Background precaching model: {model}")
                try:
                    for f in sorted(os.listdir(model_path)):
                        if self.stop_precache.is_set():
                            return
                        if f.lower().endswith(".xlsx"):
                            file_path = os.path.join(model_path, f)
                            for sheet_type in ["front", "back"]:
                                if self.stop_precache.is_set():
                                    return
                                sheet = self.excel_converter.find_sheet(file_path, sheet_type)
                                if sheet:
                                    logger.info(f"Background cache: {os.path.basename(file_path)} - {sheet_type}")
                                    self.excel_converter.convert_excel_to_png(file_path, sheet)
                                time.sleep(0.1)
                except Exception as e:
                    logger.error(f"Background precache error in {model}: {e}")
        except Exception as e:
            logger.error(f"Background precache error: {e}")import tkinter as tk
from tkinter import ttk
import os
from PIL import Image, ImageTk
import threading
import time
from datetime import datetime, timedelta
import tempfile
import shutil
import subprocess
import logging
from logging.handlers import RotatingFileHandler
import gc

LOG_FILE = "/var/log/pi-photo-viewer/app.log"
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

try:
    handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=5)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
except:
    pass

# Network drive configuration
NETWORK_BASE_PATH = "/mnt/network-drive/TS16949 Work/Standard Operating Procedure"
DEPARTMENTS = ["11 Injection", "12 Assembly", "13 Paint", "32 Repack", "New Model"]
SUPPORTED_FORMATS = (".xlsx", ".png", ".jpg", ".jpeg", ".gif", ".bmp")
LOGO_WIDTH = 175
IMAGE_CACHE_SIZE = 2
MAX_IMAGE_DIMENSION = 1920
CACHE_STALE_DAYS = 7

SHEET_MAPPING = {
    "front": ["front", "front page", "proposal"],
    "back": ["back", "back page"],
}

class ImageCache:
    def __init__(self, max_size=IMAGE_CACHE_SIZE):
        self.cache = {}
        self.max_size = max_size
        self.order = []
    
    def get(self, path):
        if path in self.cache:
            self.order.remove(path)
            self.order.append(path)
            return self.cache[path]
        return None
    
    def put(self, path, photo):
        if path in self.cache:
            self.order.remove(path)
        elif len(self.cache) >= self.max_size:
            removed = self.order.pop(0)
            if removed in self.cache:
                del self.cache[removed]
            gc.collect()
        self.cache[path] = photo
        self.order.append(path)
    
    def clear(self):
        self.cache.clear()
        self.order.clear()
        gc.collect()

class ExcelConverter:
    def __init__(self, cache_dir="/tmp/pi-photo-viewer-cache"):
        self.cache_dir = cache_dir
        self.conversion_lock = threading.Lock()
        os.makedirs(cache_dir, exist_ok=True)
    
    def find_sheet(self, excel_path, sheet_type):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            patterns = SHEET_MAPPING.get(sheet_type.lower(), [])
            for pattern in patterns:
                for sheet_name in sheet_names:
                    if pattern.lower() in sheet_name.lower():
                        logger.info(f"Matched '{sheet_type}' to '{sheet_name}'")
                        return sheet_name
            return None
        except Exception as e:
            logger.error(f"Error reading sheet names: {e}")
            return None
    
    def get_cache_path(self, excel_path, sheet_name):
        safe_name = f"{os.path.basename(excel_path)}_{sheet_name}".replace(" ", "_")
        return os.path.join(self.cache_dir, f"{safe_name}.png")
    
    def get_meta_path(self, cache_png_path):
        """Get metadata file path safely"""
        base = os.path.splitext(cache_png_path)[0]
        return base + ".meta"
    
    def get_sheet_index(self, excel_path, sheet_name):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            if sheet_name in sheet_names:
                return sheet_names.index(sheet_name)
            return None
        except Exception as e:
            logger.error(f"Error getting sheet index: {e}")
            return None
    
    def is_cache_valid(self, cache_path, source_excel_path=None):
        """Check if cache is valid"""
        if not os.path.exists(cache_path):
            return False
        
        # Check source file modification time
        if source_excel_path and os.path.exists(source_excel_path):
            try:
                excel_mod_time = os.path.getmtime(source_excel_path)
                meta_path = self.get_meta_path(cache_path)
                if os.path.exists(meta_path):
                    try:
                        with open(meta_path, 'r') as f:
                            cached_time = float(f.read().strip())
                        if excel_mod_time > cached_time:
                            logger.info(f"Source file modified, cache invalid")
                            return False
                    except Exception as e:
                        logger.warning(f"Could not read metadata: {e}")
            except Exception as e:
                logger.error(f"Error checking modification time: {e}")
        
        # Check cache age
        try:
            file_age = datetime.now() - datetime.fromtimestamp(os.path.getmtime(cache_path))
            if file_age >= timedelta(days=CACHE_STALE_DAYS):
                logger.info(f"Cache older than {CACHE_STALE_DAYS} days")
                return False
            return True
        except Exception as e:
            logger.error(f"Error checking cache age: {e}")
            return False
    
    def save_metadata(self, cache_path, source_excel_path):
        """Save source file modification time"""
        try:
            excel_mod_time = os.path.getmtime(source_excel_path)
            meta_path = self.get_meta_path(cache_path)
            with open(meta_path, 'w') as f:
                f.write(str(excel_mod_time))
        except Exception as e:
            logger.error(f"Error saving metadata: {e}")
    
    def convert_excel_to_png(self, excel_path, sheet_name):
        """Convert Excel sheet to PNG with proper locking"""
        with self.conversion_lock:
            temp_dir = None
            try:
                cache_path = self.get_cache_path(excel_path, sheet_name)
                if self.is_cache_valid(cache_path, excel_path):
                    logger.info(f"Using cached: {os.path.basename(cache_path)}")
                    return cache_path
                
                logger.info(f"Converting: {os.path.basename(excel_path)} - {sheet_name}")
                sheet_index = self.get_sheet_index(excel_path, sheet_name)
                if sheet_index is None:
                    logger.error(f"Sheet not found: {sheet_name}")
                    return None
                
                temp_dir = tempfile.mkdtemp()
                output_prefix = os.path.splitext(cache_path)[0]
                
                # Convert to PDF
                cmd = ["libreoffice", "--headless", "--invisible", "--nocrashreport",
                       "--nodefault", "--nofirststartwizard", "--nologo", "--norestore",
                       "--convert-to", "pdf", "--outdir", temp_dir, excel_path]
                result = subprocess.run(cmd, capture_output=True, timeout=45, text=True)
                
                if result.returncode != 0:
                    logger.error(f"LibreOffice failed: {result.stderr[:200]}")
                    return None
                
                pdf_files = [f for f in os.listdir(temp_dir) if f.endswith(".pdf")]
                if not pdf_files:
                    logger.error("No PDF generated")
                    return None
                
                pdf_path = os.path.join(temp_dir, pdf_files[0])
                pdf_page = sheet_index + 1
                
                # Convert PDF to PNG
                cmd = ["pdftoppm", "-png", "-f", str(pdf_page), "-l", str(pdf_page),
                       "-singlefile", "-r", "150", pdf_path, output_prefix]
                result = subprocess.run(cmd, capture_output=True, timeout=30, text=True)
                
                if result.returncode != 0:
                    logger.warning(f"pdftoppm failed, trying ImageMagick")
                    cmd = ["convert", "-density", "100", f"{pdf_path}[{sheet_index}]", cache_path]
                    result = subprocess.run(cmd, capture_output=True, timeout=30, text=True)
                    if result.returncode != 0:
                        logger.error(f"ImageMagick failed: {result.stderr[:200]}")
                        return None
                
                if os.path.exists(cache_path):
                    self.save_metadata(cache_path, excel_path)
                    logger.info(f"Conversion complete: {os.path.basename(cache_path)}")
                    return cache_path
                
                logger.error("PNG file not created")
                return None
            
            except subprocess.TimeoutExpired:
                logger.error("Conversion timeout")
                return None
            except Exception as e:
                logger.error(f"Conversion error: {e}")
                return None
            finally:
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except Exception as e:
                        logger.error(f"Temp cleanup error: {e}")
                gc.collect()

class FullscreenImageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pi Standards Viewer - Network")
        logger.info("App starting")
        
        self.root.after(200, lambda: self.root.attributes("-fullscreen", True))
        self.root.configure(bg="black")
        
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox", fieldbackground="#374151", background="#4B5563",
                       foreground="#06B6D4", arrowcolor="#06B6D4", arrowsize=30)
        self.root.option_add("*TCombobox*Listbox.font", ("Helvetica", 22, "bold"))
        self.root.option_add("*TCombobox*Listbox.background", "#1F2937")
        self.root.option_add("*TCombobox*Listbox.foreground", "#06B6D4")
        
        self.control_bar_collapsed_height = 80
        self.control_bar_expanded_height = 220
        self.collapse_timer = None
        self.collapse_delay = 30000
        
        # Control frame
        self.control_frame = tk.Frame(root, bg="#1F2937", pady=10, padx=30)
        self.control_frame.pack(side="bottom", fill="x")
        self.control_frame.pack_propagate(False)
        self.control_frame.config(height=self.control_bar_collapsed_height)
        
        self.expand_indicator = tk.Label(self.control_frame, text="▲ TAP TO SELECT FILE ▲",
                                         bg="#1F2937", fg="#06B6D4", font=("Helvetica", 14, "bold"),
                                         cursor="hand2")
        self.expand_indicator.pack(side="top", pady=(0, 5))
        self.expand_indicator.bind("<Button-1>", lambda e: self.expand_controls())
        
        self.collapsed_container = tk.Frame(self.control_frame, bg="#1F2937")
        self.collapsed_container.pack(fill="both", expand=True)
        
        self.expanded_container = tk.Frame(self.control_frame, bg="#1F2937")
        
        # Collapsed view - page buttons only
        collapsed_label = tk.Label(self.collapsed_container, text="Page:", bg="#1F2937",
                                  fg="#06B6D4", font=("Helvetica", 20, "bold"))
        collapsed_label.pack(side="left", padx=(0, 15))
        
        self.collapsed_button_frame = tk.Frame(self.collapsed_container, bg="#1F2937")
        self.collapsed_button_frame.pack(side="left", expand=True, fill="both")
        
        self.front_button = tk.Button(self.collapsed_button_frame, text="FRONT PAGE",
                                     font=("Helvetica", 22, "bold"), bg="#06B6D4", fg="#1F2937",
                                     activebackground="#0891B2", relief="sunken", bd=3, cursor="hand2",
                                     command=lambda: self.on_page_click("Front"))
        self.back_button = tk.Button(self.collapsed_button_frame, text="BACK PAGE",
                                    font=("Helvetica", 22, "bold"), bg="#4B5563", fg="#06B6D4",
                                    activebackground="#374151", relief="raised", bd=3, cursor="hand2",
                                    command=lambda: self.on_page_click("Back"))
        
        self.front_button.pack(side="left", expand=True, fill="both", padx=(0, 10))
        self.back_button.pack(side="left", expand=True, fill="both")
        
        # Expanded view - all dropdowns
        # Row 1: Department and Part Model
        row1 = tk.Frame(self.expanded_container, bg="#1F2937")
        row1.pack(fill="x", pady=(0, 10))
        
        tk.Label(row1, text="Department:", bg="#1F2937", fg="#06B6D4",
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.dept_var = tk.StringVar(root)
        self.dept_dropdown = ttk.Combobox(row1, textvariable=self.dept_var,
                                         font=("Helvetica", 22, "bold"), state="readonly", height=5,
                                         postcommand=self.on_dropdown_open)
        self.dept_dropdown["values"] = DEPARTMENTS
        self.dept_dropdown.pack(side="left", expand=True, fill="both", padx=(0, 20))
        self.dept_dropdown.bind("<<ComboboxSelected>>", self.on_dept_select)
        
        tk.Label(row1, text="Part Model:", bg="#1F2937", fg="#06B6D4",
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.model_var = tk.StringVar(root)
        self.model_dropdown = ttk.Combobox(row1, textvariable=self.model_var,
                                          font=("Helvetica", 22, "bold"), state="readonly", height=5,
                                          postcommand=self.on_dropdown_open)
        self.model_dropdown.pack(side="left", expand=True, fill="both")
        self.model_dropdown.bind("<<ComboboxSelected>>", self.on_model_select)
        
        # Row 2: File
        row2 = tk.Frame(self.expanded_container, bg="#1F2937")
        row2.pack(fill="x", pady=(0, 10))
        
        tk.Label(row2, text="File:", bg="#1F2937", fg="#06B6D4",
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.file_var = tk.StringVar(root)
        self.file_dropdown = ttk.Combobox(row2, textvariable=self.file_var,
                                         font=("Helvetica", 22, "bold"), state="readonly", height=5,
                                         postcommand=self.on_dropdown_open)
        self.file_dropdown.pack(side="left", expand=True, fill="both")
        self.file_dropdown.bind("<<ComboboxSelected>>", self.on_file_select)
        
        # Row 3: Page buttons
        row3 = tk.Frame(self.expanded_container, bg="#1F2937")
        row3.pack(fill="x")
        tk.Label(row3, text="Page:", bg="#1F2937", fg="#06B6D4",
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        
        self.button_frame = tk.Frame(row3, bg="#1F2937")
        self.button_frame.pack(side="left", expand=True, fill="both")
        
        self.front_button_exp = tk.Button(self.button_frame, text="FRONT PAGE",
                                         font=("Helvetica", 22, "bold"), bg="#06B6D4", fg="#1F2937",
                                         activebackground="#0891B2", relief="sunken", bd=3, cursor="hand2",
                                         command=lambda: self.on_page_click("Front"))
        self.back_button_exp = tk.Button(self.button_frame, text="BACK PAGE",
                                        font=("Helvetica", 22, "bold"), bg="#4B5563", fg="#06B6D4",
                                        activebackground="#374151", relief="raised", bd=3, cursor="hand2",
                                        command=lambda: self.on_page_click("Back"))
        
        self.front_button_exp.pack(side="left", expand=True, fill="both", padx=(0, 10))
        self.back_button_exp.pack(side="left", expand=True, fill="both")
        
        # Main image display
        self.image_label = tk.Label(root, bg="black", fg="white", font=("Helvetica", 24))
        self.image_label.pack(expand=True, fill="both")
        self.image_label.bind("<Button-1>", lambda e: self.expand_controls())
        
        # State
        self.current_page = "Front"
        self.is_expanded = False
        self.current_file_path = None
        self.current_model_path = None
        self.files_list = []
        self.image_cache = ImageCache()
        self.excel_converter = ExcelConverter()
        self.stop_precache = threading.Event()
        self.precache_thread = None
        self.fg_precache_thread = None
        
        # Initialize
        if DEPARTMENTS:
            self.dept_var.set(DEPARTMENTS[0])
            self.root.after(100, lambda: self.on_dept_select(None))
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def on_dropdown_open(self):
        if self.collapse_timer:
            self.root.after_cancel(self.collapse_timer)
            self.collapse_timer = None
        self.reset_collapse_timer()
    
    def expand_controls(self):
        if self.is_expanded:
            return
        self.is_expanded = True
        self.collapsed_container.pack_forget()
        self.expanded_container.pack(fill="both", expand=True)
        self.expand_indicator.config(text="▼ TAP IMAGE TO HIDE ▼")
        self.control_frame.config(height=self.control_bar_expanded_height)
        self.reset_collapse_timer()
    
    def collapse_controls(self):
        if not self.is_expanded:
            return
        self.is_expanded = False
        self.expanded_container.pack_forget()
        self.collapsed_container.pack(fill="both", expand=True)
        self.expand_indicator.config(text="▲ TAP TO SELECT FILE ▲")
        self.control_frame.config(height=self.control_bar_collapsed_height)
        if self.collapse_timer:
            self.root.after_cancel(self.collapse_timer)
            self.collapse_timer = None
    
    def reset_collapse_timer(self):
        if self.collapse_timer:
            self.root.after_cancel(self.collapse_timer)
        self.collapse_timer = self.root.after(self.collapse_delay, self.collapse_controls)
    
    def on_page_click(self, page):
        self.current_page = page
        for front, back in [(self.front_button, self.back_button),
                           (self.front_button_exp, self.back_button_exp)]:
            if page == "Front":
                front.config(bg="#06B6D4", fg="#1F2937", relief="sunken")
                back.config(bg="#4B5563", fg="#06B6D4", relief="raised")
            else:
                back.config(bg="#06B6D4", fg="#1F2937", relief="sunken")
                front.config(bg="#4B5563", fg="#06B6D4", relief="raised")
        
        if self.current_file_path:
            self.display_file(self.current_file_path, page)
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def on_close(self):
        logger.info("App closing")
        self.stop_precache.set()
        self.root.destroy()
    
    def on_dept_select(self, event):
        logger.info(f"Department: {self.dept_var.get()}")
        self.update_models()
        self.stop_precache.set()
        time.sleep(0.1)
        self.stop_precache.clear()
        if self.precache_thread is None or not self.precache_thread.is_alive():
            self.precache_thread = threading.Thread(
                target=self.precache_dept,
                args=(self.dept_var.get(),),
                daemon=True
            )
            self.precache_thread.start()
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def update_models(self):
        dept = self.dept_var.get()
        if not dept:
            self.model_dropdown["values"] = []
            return
        
        dept_path = os.path.join(NETWORK_BASE_PATH, dept)
        try:
            models = sorted([d for d in os.listdir(dept_path)
                           if os.path.isdir(os.path.join(dept_path, d))])
            self.model_dropdown["values"] = models
            if models:
                self.model_var.set(models[0])
                self.on_model_select(None)
        except Exception as e:
            logger.error(f"Error listing models: {e}")
            self.model_dropdown["values"] = []
    
    def on_model_select(self, event):
        self.update_files()
        # Start aggressive foreground precaching for this model
        if self.fg_precache_thread is None or not self.fg_precache_thread.is_alive():
            self.fg_precache_thread = threading.Thread(
                target=self.precache_model_aggressive,
                args=(self.current_model_path,),
                daemon=True
            )
            self.fg_precache_thread.start()
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def update_files(self):
        dept = self.dept_var.get()
        model = self.model_var.get()
        if not dept or not model:
            self.file_dropdown["values"] = []
            return
        
        self.current_model_path = os.path.join(NETWORK_BASE_PATH, dept, model)
        try:
            self.files_list = sorted([os.path.join(self.current_model_path, f)
                                     for f in os.listdir(self.current_model_path)
                                     if os.path.isfile(os.path.join(self.current_model_path, f))
                                     and f.lower().endswith(SUPPORTED_FORMATS)])
            names = [os.path.splitext(os.path.basename(f))[0] for f in self.files_list]
            self.file_dropdown["values"] = names
            if names:
                self.file_var.set(names[0])
                self.on_file_select(None)
        except Exception as e:
            logger.error(f"Error listing files: {e}")
            self.file_dropdown["values"] = []
            self.files_list = []
    
    def on_file_select(self, event):
        name = self.file_var.get()
        if not name:
            return
        
        for f in self.files_list:
            if os.path.splitext(os.path.basename(f))[0] == name:
                self.current_file_path = f
                break
        
        if not self.current_file_path:
            return
        
        if self.current_file_path.lower().endswith(".xlsx"):
            for btn in [self.front_button, self.back_button, self.front_button_exp, self.back_button_exp]:
                btn.config(state="normal")
            self.on_page_click("Front")
        else:
            for btn in [self.front_button, self.back_button, self.front_button_exp, self.back_button_exp]:
                btn.config(state="disabled")
            self.display_file(self.current_file_path, "Image")
        
        self.root.after(2000, self.collapse_controls)
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def display_file(self, path, page):
        if not path:
            return
        
        cache_key = f"{path}_{page}"
        cached = self.image_cache.get(cache_key)
        if cached:
            self.image_label.config(image=cached, text="")
            self.image_label.image = cached
        else:
            if path.lower().endswith(".xlsx"):
                self.image_label.config(image="", text=f"Loading {page}...")
            else:
                self.image_label.config(image="", text="Loading...")
            threading.Thread(target=self.load_file, args=(path, page, cache_key), daemon=True).start()
    
    def load_file(self, path, page, cache_key):
        try:
            if path.lower().endswith(".xlsx"):
                sheet_type = page.lower() if page != "Image" else "front"
                sheet = self.excel_converter.find_sheet(path, sheet_type)
                if not sheet:
                    self.root.after(0, lambda: self.image_label.config(image="", text=f"Sheet not found: {page}"))
                    return
                png_path = self.excel_converter.convert_excel_to_png(path, sheet)
                if not png_path:
                    self.root.after(0, lambda: self.image_label.config(image="", text=f"Conversion failed"))
                    return
                img = Image.open(png_path)
            else:
                img = Image.open(path)
            
            w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight() - self.control_bar_collapsed_height
            img.thumbnail((min(w, MAX_IMAGE_DIMENSION), min(h, MAX_IMAGE_DIMENSION)), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.image_cache.put(cache_key, photo)
            self.root.after(0, lambda: (self.image_label.config(image=photo, text=""),
                                       setattr(self.image_label, 'image', photo)))
        except Exception as e:
            logger.error(f"Load error: {e}")
            self.root.after(0, lambda: self.image_label.config(image="", text=f"Error: {os.path.basename(path)}"))
    
    def precache_dept(self, dept):
        logger.info(f"Precaching department: {dept}")
        dept_path = os.path.join(NETWORK_BASE_PATH, dept)
        try:
            models = [d for d in os.listdir(dept_path) if os.path.isdir(os.path.join(dept_path, d))]
            for model in sorted(models):
                if self.stop_precache.is_set():
                    return
                model_path = os.path.join(dept_path, model)
                try:
                    for f in os.listdir(model_path):
                        if self.stop_precache.is_set():
                            return
                        if f.lower().endswith(".xlsx"):
                            file_path = os.path.join(model_path, f)
                            for sheet_type in ["front", "back"]:
                                if self.stop_precache.is_set():
                                    return
                                sheet = self.excel_converter.find_sheet(file_path, sheet_type)
                                if sheet:
                                    logger.info(f"Precaching: {os.path.basename(file_path)} - {sheet_type}")
                                    self.excel_converter.convert_excel_to_png(file_path, sheet)
                                time.sleep(0.3)
                except Exception as e:
                    logger.error(f"Precache error in {model}: {e}")
        except Exception as e:
            logger.error(f"Precache error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FullscreenImageApp(root)
    root.mainloop()
