def precache_excel_files(self, folder_path):
        """Precache all Excel files in selected part model"""
        logger.info(f"Pre-caching Excel files in {folder_path}")
        try:
            files = self.files_in_part_model
            excel_files = [f for f in files if f.lower().endswith(".xlsx")]
            for excel_file in excel_files:
                if self.stop_background_precache:
                    return
                for sheet_type in ["front", "back"]:
                    if self.stop_background_precache:
                        return
                    actual_sheet = self.excel_converter.find_sheet(excel_file, sheet_type)
                    if actual_sheet:
                        logger.info(f"Pre-caching {os.path.basename(excel_file)} - {sheet_type}")
                        self.excel_converter.convert_excel_to_png(excel_file, actual_sheet)
            logger.info("Pre-caching complete")
        except Exception as e:
            logger.error(f"Pre-cache error: {e}")
    
    def background_precache_department(self, department):
        """Background precache all Excel files in entire department"""
        logger.info(f"Starting background precache of department: {department}")
        try:
            dept_path = os.path.join(NETWORK_BASE_PATH, department)
            
            if not os.path.exists(dept_path):
                logger.error(f"Department path does not exist: {dept_path}")
                return
            
            # Iterate through all part models (folders) in department
            try:
                part_models = [item for item in os.listdir(dept_path) 
                              if os.path.isdir(os.path.join(dept_path, item))]
                part_models.sort()
            except Exception as e:
                logger.error(f"Error listing part models: {e}")
                return
            
            logger.info(f"Background precaching {len(part_models)} part models")
            
            for part_model in part_models:
                if self.stop_background_precache:
                    logger.info("Background precache stopped")
                    return
                
                part_model_path = os.path.join(dept_path, part_model)
                logger.info(f"Background precaching part model: {part_model}")
                
                try:
                    # Get all Excel files in this part model
                    files = []
                    for item in os.listdir(part_model_path):
                        if self.stop_background_precache:
                            return
                        item_path = os.path.join(part_model_path, item)
                        if os.path.isfile(item_path) and item.lower().endswith(".xlsx"):
                            files.append(item_path)
                    
                    files.sort()
                    
                    # Precache each Excel file
                    for excel_file in files:
                        if self.stop_background_precache:
                            return
                        
                        for sheet_type in ["front", "back"]:
                            if self.stop_background_precache:
                                return
                            
                            try:
                                actual_sheet = self.excel_converter.find_sheet(excel_file, sheet_type)
                                if actual_sheet:
                                    logger.info(f"Background caching {os.path.basename(excel_file)} - {sheet_type}")
                                    self.excel_converter.convert_excel_to_png(excel_file, actual_sheet)
                            except Exception as e:
                                logger.error(f"Error precaching {excel_file}: {e}")
                            
                            time.sleep(0.5)  # Small delay between conversions
                
                except Exception as e:
                    logger.error(f"Error processing part model {part_model}: {e}")
                    continue
            
            logger.info(f"Background precache of department {department} complete")
        except Exception as e:
            logger.error(f"Background precache error: {e}")import tkinter as tk
from tkinter import ttk
import os
from PIL import Image, ImageTk
import threading
import time
from collections import Counter
import subprocess
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime, timedelta
import tempfile
import shutil
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

# Network drive path
NETWORK_BASE_PATH = "/mnt/network-drive/TS16949 Work/Standard Operating Procedure"

# Department folders to show
DEPARTMENTS = ["11 Injection", "12 Assembly", "13 Paint", "32 Repack", "New Model"]

SUPPORTED_FORMATS = (".xlsx", ".png", ".jpg", ".jpeg", ".gif", ".bmp")
LOGO_WIDTH = 175
IMAGE_CACHE_SIZE = 2
MAX_IMAGE_DIMENSION = 1920
CACHE_STALE_DAYS = 7

SHEET_MAPPING = {
    "front": ["front", "front page", "proposal"],
    "back": ["back", "back page"],
    "hidden": ["changelog", "revision history"]
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
        if not os.path.exists(cache_path):
            return False
        
        if source_excel_path and os.path.exists(source_excel_path):
            try:
                excel_mod_time = os.path.getmtime(source_excel_path)
                meta_path = cache_path.replace(".png", ".meta")
                if os.path.exists(meta_path):
                    with open(meta_path, 'r') as f:
                        cached_excel_mod_time = float(f.read().strip())
                    if excel_mod_time > cached_excel_mod_time:
                        logger.info(f"Excel file modified, cache is stale: {source_excel_path}")
                        return False
            except Exception as e:
                logger.error(f"Error checking cache validity: {e}")
        
        try:
            file_age = datetime.now() - datetime.fromtimestamp(os.path.getmtime(cache_path))
            is_stale = file_age >= timedelta(days=CACHE_STALE_DAYS)
            if is_stale:
                logger.info(f"Cache file older than {CACHE_STALE_DAYS} days, refreshing")
            return not is_stale
        except Exception as e:
            logger.error(f"Error checking cache age: {e}")
            return False
    
    def save_cache_metadata(self, cache_path, source_excel_path):
        try:
            excel_mod_time = os.path.getmtime(source_excel_path)
            meta_path = cache_path.replace(".png", ".meta")
            with open(meta_path, 'w') as f:
                f.write(str(excel_mod_time))
            logger.info(f"Saved cache metadata for {cache_path}")
        except Exception as e:
            logger.error(f"Error saving cache metadata: {e}")
    
    def convert_excel_to_png(self, excel_path, sheet_name):
        with self.conversion_lock:
            temp_dir = None
            try:
                cache_path = self.get_cache_path(excel_path, sheet_name)
                if self.is_cache_valid(cache_path, excel_path):
                    logger.info(f"Using cached image: {cache_path}")
                    return cache_path
                
                logger.info(f"Converting Excel sheet '{sheet_name}' to PNG...")
                sheet_index = self.get_sheet_index(excel_path, sheet_name)
                if sheet_index is None:
                    return None
                
                temp_dir = tempfile.mkdtemp()
                logger.info(f"Created temp dir: {temp_dir}")
                
                try:
                    output_prefix = cache_path.replace(".png", "")
                    cmd = ["libreoffice", "--headless", "--invisible", "--nocrashreport", 
                           "--nodefault", "--nofirststartwizard", "--nologo", "--norestore",
                           "--convert-to", "pdf", "--outdir", temp_dir, excel_path]
                    
                    logger.info("Starting LibreOffice conversion...")
                    result = subprocess.run(cmd, capture_output=True, timeout=45, text=True)
                    
                    if result.returncode != 0:
                        logger.error(f"LibreOffice conversion failed: {result.stderr}")
                        return None
                    
                    pdf_files = [f for f in os.listdir(temp_dir) if f.endswith(".pdf")]
                    if not pdf_files:
                        logger.error("No PDF generated")
                        return None
                    
                    pdf_path = os.path.join(temp_dir, pdf_files[0])
                    pdf_page = sheet_index + 1
                    logger.info(f"PDF created, extracting page {pdf_page}...")
                    
                    cmd = ["pdftoppm", "-png", "-f", str(pdf_page), "-l", str(pdf_page), 
                           "-singlefile", "-r", "150", pdf_path, output_prefix]
                    result = subprocess.run(cmd, capture_output=True, timeout=30, text=True)
                    
                    if result.returncode != 0:
                        logger.warning(f"pdftoppm failed: {result.stderr}")
                        logger.info("Trying ImageMagick fallback...")
                        cmd = ["convert", "-density", "100", f"{pdf_path}[{sheet_index}]", cache_path]
                        result = subprocess.run(cmd, capture_output=True, timeout=30, text=True)
                        if result.returncode != 0:
                            logger.error(f"ImageMagick also failed: {result.stderr}")
                            return None
                    
                    if os.path.exists(cache_path):
                        logger.info(f"Successfully converted to PNG: {cache_path}")
                        self.save_cache_metadata(cache_path, excel_path)
                        return cache_path
                    else:
                        logger.error("PNG file was not created")
                        return None
                except subprocess.TimeoutExpired:
                    logger.error("Conversion timed out")
                    return None
                except Exception as e:
                    logger.error(f"Conversion step error: {e}")
                    return None
            except Exception as e:
                logger.error(f"Conversion error: {e}")
                return None
            finally:
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        logger.info(f"Cleaning up temp dir: {temp_dir}")
                        shutil.rmtree(temp_dir, ignore_errors=True)
                        time.sleep(0.1)
                    except Exception as e:
                        logger.error(f"Failed to clean temp dir: {e}")
                gc.collect()

class FullscreenImageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pi Standards Viewer")
        logger.info("Application starting")
        
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
        
        self.control_frame = tk.Frame(root, bg="#1F2937", pady=10, padx=30)
        self.control_frame.pack(side="bottom", fill="x")
        self.control_frame.pack_propagate(False)
        self.control_frame.config(height=self.control_bar_collapsed_height)
        
        self.expand_indicator = tk.Label(self.control_frame, text="▲ TAP TO SELECT FILE ▲", 
                                         bg="#1F2937", fg="#06B6D4", font=("Helvetica", 14, "bold"), cursor="hand2")
        self.expand_indicator.pack(side="top", pady=(0, 5))
        self.expand_indicator.bind("<Button-1>", lambda e: self.expand_controls())
        
        self.collapsed_container = tk.Frame(self.control_frame, bg="#1F2937")
        self.collapsed_container.pack(fill="both", expand=True)
        
        self.expanded_container = tk.Frame(self.control_frame, bg="#1F2937")
        
        collapsed_label = tk.Label(self.collapsed_container, text="Page:", bg="#1F2937", 
                                   fg="#06B6D4", font=("Helvetica", 20, "bold"))
        collapsed_label.pack(side="left", padx=(0, 15))
        
        self.collapsed_button_frame = tk.Frame(self.collapsed_container, bg="#1F2937")
        self.collapsed_button_frame.pack(side="left", expand=True, fill="both")
        
        self.front_button = tk.Button(self.collapsed_button_frame, text="FRONT PAGE", 
                                     font=("Helvetica", 22, "bold"), bg="#06B6D4", fg="#1F2937",
                                     activebackground="#0891B2", relief="sunken", bd=3, cursor="hand2",
                                     command=lambda: self.on_page_button_click("Front"))
        self.back_button = tk.Button(self.collapsed_button_frame, text="BACK PAGE", 
                                    font=("Helvetica", 22, "bold"), bg="#4B5563", fg="#06B6D4",
                                    activebackground="#374151", relief="raised", bd=3, cursor="hand2",
                                    command=lambda: self.on_page_button_click("Back"))
        
        self.front_button.pack(side="left", expand=True, fill="both", padx=(0, 10))
        self.back_button.pack(side="left", expand=True, fill="both")
        
        # Row 1: Department and Part Model dropdowns
        exp_row1 = tk.Frame(self.expanded_container, bg="#1F2937")
        exp_row1.pack(fill="x", pady=(0, 10))
        
        tk.Label(exp_row1, text="Department:", bg="#1F2937", fg="#06B6D4", 
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.department_variable = tk.StringVar(root)
        self.department_dropdown = ttk.Combobox(exp_row1, textvariable=self.department_variable, 
                                               font=("Helvetica", 22, "bold"), state="readonly", 
                                               height=6, postcommand=lambda: self.on_dropdown_open())
        self.department_dropdown["values"] = DEPARTMENTS
        self.department_dropdown.pack(side="left", expand=True, fill="both", padx=(0, 20))
        self.department_dropdown.bind("<<ComboboxSelected>>", self.on_department_select)
        
        tk.Label(exp_row1, text="Part Model:", bg="#1F2937", fg="#06B6D4", 
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.part_model_variable = tk.StringVar(root)
        self.part_model_dropdown = ttk.Combobox(exp_row1, textvariable=self.part_model_variable, 
                                               font=("Helvetica", 22, "bold"), state="readonly", 
                                               height=6, postcommand=lambda: self.on_dropdown_open())
        self.part_model_dropdown.pack(side="left", expand=True, fill="both")
        self.part_model_dropdown.bind("<<ComboboxSelected>>", self.on_part_model_select)
        
        # Row 2: File dropdown and Page buttons
        exp_row2 = tk.Frame(self.expanded_container, bg="#1F2937")
        exp_row2.pack(fill="x", pady=(0, 10))
        
        tk.Label(exp_row2, text="File:", bg="#1F2937", fg="#06B6D4", 
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.file_variable = tk.StringVar(root)
        self.file_dropdown = ttk.Combobox(exp_row2, textvariable=self.file_variable, 
                                         font=("Helvetica", 22, "bold"), state="readonly", 
                                         height=6, postcommand=lambda: self.on_dropdown_open())
        self.file_dropdown.pack(side="left", expand=True, fill="both")
        self.file_dropdown.bind("<<ComboboxSelected>>", self.on_file_select)
        
        # Row 3: Page selection buttons
        exp_row3 = tk.Frame(self.expanded_container, bg="#1F2937")
        exp_row3.pack(fill="x")
        tk.Label(exp_row3, text="Page:", bg="#1F2937", fg="#06B6D4", 
                font=("Helvetica", 20, "bold")).pack(side="left", padx=(0, 10))
        self.expanded_button_frame = tk.Frame(exp_row3, bg="#1F2937")
        self.expanded_button_frame.pack(side="left", expand=True, fill="both")
        
        self.front_button_exp = tk.Button(self.expanded_button_frame, text="FRONT PAGE", 
                                         font=("Helvetica", 22, "bold"), bg="#06B6D4", fg="#1F2937",
                                         activebackground="#0891B2", relief="sunken", bd=3, cursor="hand2",
                                         command=lambda: self.on_page_button_click("Front"))
        self.back_button_exp = tk.Button(self.expanded_button_frame, text="BACK PAGE", 
                                        font=("Helvetica", 22, "bold"), bg="#4B5563", fg="#06B6D4",
                                        activebackground="#374151", relief="raised", bd=3, cursor="hand2",
                                        command=lambda: self.on_page_button_click("Back"))
        
        self.front_button_exp.pack(side="left", expand=True, fill="both", padx=(0, 10))
        self.back_button_exp.pack(side="left", expand=True, fill="both")
        
        self.image_label = tk.Label(root, bg="black", fg="white", font=("Helvetica", 24))
        self.image_label.pack(expand=True, fill="both")
        self.image_label.bind("<Button-1>", lambda e: self.expand_controls())
        
        self.logo_label = tk.Label(root, bg="black")
        self.logo_label.place(relx=1.0, rely=0.0, anchor="ne", x=-10, y=10)
        
        self.current_page = "Front"
        self.is_expanded = False
        self.current_department = None
        self.current_part_model = None
        self.current_part_model_path = None
        self.files_in_part_model = []
        self.current_file_path = None
        self.image_cache = ImageCache()
        self.excel_converter = ExcelConverter()
        self.precache_thread = None
        self.current_photoimage = None
        self.background_precache_thread = None
        self.stop_background_precache = False
        
        # Initialize by setting first department
        if DEPARTMENTS:
            self.department_variable.set(DEPARTMENTS[0])
            self.on_department_select(None)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
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
    
    def on_page_button_click(self, page):
        self.current_page = page
        buttons = [(self.front_button, self.back_button), (self.front_button_exp, self.back_button_exp)]
        for front_btn, back_btn in buttons:
            if page == "Front":
                front_btn.config(bg="#06B6D4", fg="#1F2937", relief="sunken")
                back_btn.config(bg="#4B5563", fg="#06B6D4", relief="raised")
            else:
                back_btn.config(bg="#06B6D4", fg="#1F2937", relief="sunken")
                front_btn.config(bg="#4B5563", fg="#06B6D4", relief="raised")
        
        if self.current_file_path:
            self.display_file_async(self.current_file_path, page)
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def on_closing(self):
        logger.info("Application closing")
        self.stop_background_precache = True
        self.root.destroy()
    
    def on_department_select(self, event):
        """Handle department selection"""
        department = self.department_variable.get()
        if not department:
            return
        
        logger.info(f"Department selected: {department}")
        self.current_department = department
        
        # Get part models for this department
        self.update_part_models()
        
        # Start background precaching of entire department
        self.stop_background_precache = False
        if self.background_precache_thread is None or not self.background_precache_thread.is_alive():
            self.background_precache_thread = threading.Thread(
                target=self.background_precache_department,
                args=(department,),
                daemon=True
            )
            self.background_precache_thread.start()
        
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def update_part_models(self):
        """Get list of folders (part models) in selected department"""
        if not self.current_department:
            self.part_model_dropdown["values"] = []
            return
        
        dept_path = os.path.join(NETWORK_BASE_PATH, self.current_department)
        
        try:
            if not os.path.exists(dept_path):
                logger.error(f"Department path does not exist: {dept_path}")
                self.part_model_dropdown["values"] = []
                return
            
            # Get only folders, ignore files
            part_models = []
            for item in os.listdir(dept_path):
                item_path = os.path.join(dept_path, item)
                if os.path.isdir(item_path):
                    part_models.append(item)
            
            part_models.sort()
            logger.info(f"Found {len(part_models)} part models in {self.current_department}")
            self.part_model_dropdown["values"] = part_models
            
            if part_models:
                self.part_model_variable.set(part_models[0])
                self.on_part_model_select(None)
            else:
                self.part_model_variable.set("")
                self.file_dropdown["values"] = []
                self.file_variable.set("")
        except Exception as e:
            logger.error(f"Error reading part models: {e}")
            self.part_model_dropdown["values"] = []
    
    def on_part_model_select(self, event):
        """Handle part model selection"""
        part_model = self.part_model_variable.get()
        if not part_model or not self.current_department:
            return
        
        logger.info(f"Part model selected: {part_model}")
        self.current_part_model = part_model
        self.current_part_model_path = os.path.join(NETWORK_BASE_PATH, self.current_department, part_model)
        
        # Get files in this part model
        self.update_file_list()
        
        # Start precaching for this part model
        if self.precache_thread is None or not self.precache_thread.is_alive():
            self.precache_thread = threading.Thread(
                target=self.precache_excel_files, 
                args=(self.current_part_model_path,), 
                daemon=True
            )
            self.precache_thread.start()
        
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def update_file_list(self):
        """Get list of files in selected part model"""
        if not self.current_part_model_path:
            self.file_dropdown["values"] = []
            self.file_variable.set("")
            return
        
        try:
            if not os.path.exists(self.current_part_model_path):
                logger.error(f"Part model path does not exist: {self.current_part_model_path}")
                self.file_dropdown["values"] = []
                return
            
            # Get all supported files
            files = []
            for item in os.listdir(self.current_part_model_path):
                item_path = os.path.join(self.current_part_model_path, item)
                if os.path.isfile(item_path) and item.lower().endswith(SUPPORTED_FORMATS):
                    files.append(item_path)
            
            files.sort()
            self.files_in_part_model = files
            
            # Display without extension
            display_names = [os.path.splitext(os.path.basename(f))[0] for f in files]
            logger.info(f"Found {len(files)} files in {self.current_part_model}")
            self.file_dropdown["values"] = display_names
            
            if display_names:
                self.file_variable.set(display_names[0])
                self.on_file_select(None)
            else:
                self.file_variable.set("")
                self.image_label.config(image="", text="No files found")
        except Exception as e:
            logger.error(f"Error reading files: {e}")
            self.file_dropdown["values"] = []
    
    def on_file_select(self, event):
        """Handle file selection"""
        selected_name = self.file_variable.get()
        if not selected_name:
            return
        
        # Find matching file
        file_path = None
        for f in self.files_in_part_model:
            if os.path.splitext(os.path.basename(f))[0] == selected_name:
                file_path = f
                break
        
        if not file_path:
            return
        
        logger.info(f"File selected: {file_path}")
        self.current_file_path = file_path
        
        if file_path.lower().endswith(".xlsx"):
            for btn in [self.front_button, self.back_button, self.front_button_exp, self.back_button_exp]:
                btn.config(state="normal")
            self.on_page_button_click("Front")
        else:
            for btn in [self.front_button, self.back_button, self.front_button_exp, self.back_button_exp]:
                btn.config(state="disabled")
            self.display_file_async(file_path, "Image")
        
        self.root.after(2000, self.collapse_controls)
        if self.is_expanded:
            self.reset_collapse_timer()
    
    def display_file_async(self, file_path, page):
        if not file_path:
            return
        
        cache_key = f"{file_path}_{page}"
        cached = self.image_cache.get(cache_key)
        
        if cached:
            self.update_ui_with_image(cached)
        else:
            if file_path.lower().endswith(".xlsx"):
                self.image_label.config(image="", text=f"Converting Excel sheet: {page}\n\nPlease wait...")
            else:
                self.image_label.config(image="", text="Loading...")
            thread = threading.Thread(target=self.load_file_threaded, args=(file_path, page, cache_key), daemon=True)
            thread.start()
    
    def load_file_threaded(self, file_path, page, cache_key):
        try:
            logger.info(f"load_file_threaded START: {file_path}, page={page}")
            if file_path.lower().endswith(".xlsx"):
                logger.info("Processing Excel file")
                sheet_type = page.lower() if page != "Image" else "front"
                logger.info(f"Sheet type: {sheet_type}")
                actual_sheet = self.excel_converter.find_sheet(file_path, sheet_type)
                
                if not actual_sheet:
                    logger.error(f"Could not find sheet: {sheet_type}")
                    self.root.after(0, self.update_ui_with_error, f"Could not find '{page}' sheet")
                    return
                
                logger.info(f"Found sheet: {actual_sheet}")
                png_path = self.excel_converter.convert_excel_to_png(file_path, actual_sheet)
                logger.info(f"Conversion result: {png_path}")
                
                if not png_path:
                    logger.error("Conversion returned None")
                    self.root.after(0, self.update_ui_with_error, f"Failed to convert {page}")
                    return
                
                try:
                    logger.info(f"Opening PNG: {png_path}")
                    img = Image.open(png_path)
                    logger.info(f"PNG opened successfully, size: {img.size}")
                except Exception as e:
                    logger.error(f"Failed to open PNG: {e}", exc_info=True)
                    self.root.after(0, self.update_ui_with_error, f"Error opening converted image")
                    return
            else:
                logger.info("Processing regular image file")
                try:
                    logger.info(f"Opening image: {file_path}")
                    img = Image.open(file_path)
                    logger.info(f"Image opened successfully, size: {img.size}")
                except Exception as e:
                    logger.error(f"Failed to open image: {e}", exc_info=True)
                    self.root.after(0, self.update_ui_with_error, os.path.basename(file_path))
                    return
            
            try:
                logger.info("Getting screen dimensions")
                screen_width = self.root.winfo_screenwidth()
                available_height = self.root.winfo_screenheight() - self.control_bar_collapsed_height
                logger.info(f"Screen: {screen_width}x{available_height}")
                
                max_dim = (min(screen_width, MAX_IMAGE_DIMENSION), 
                          min(available_height, MAX_IMAGE_DIMENSION))
                logger.info(f"Thumbnailing to: {max_dim}")
                img.thumbnail(max_dim, Image.LANCZOS)
                logger.info(f"Thumbnail created, new size: {img.size}")
                
                logger.info("Creating PhotoImage")
                photo = ImageTk.PhotoImage(img)
                logger.info("PhotoImage created successfully")
                
                logger.info("Adding to cache")
                self.image_cache.put(cache_key, photo)
                logger.info("Cache updated, scheduling UI update")
                
                self.root.after(0, self.update_ui_with_image, photo)
                logger.info("UI update scheduled")
            except Exception as e:
                logger.error(f"Error processing image: {e}", exc_info=True)
                self.root.after(0, self.update_ui_with_error, "Image processing error")
        except Exception as e:
            logger.error(f"Error loading file {file_path}: {e}", exc_info=True)
            self.root.after(0, self.update_ui_with_error, os.path.basename(file_path))
    
    def update_ui_with_image(self, photo):
        try:
            self.image_label.config(image=photo, text="")
            self.image_label.image = photo
        except Exception as e:
            logger.error(f"Error updating UI with image: {e}")
    
    def update_ui_with_error(self, filename):
        self.image_label.config(image="", text=f"Error loading:\n{filename}")
    
    def load_logo_threaded(self, path):
        try:
            img = Image.open(path)
            w, h = img.size
            new_height = int((LOGO_WIDTH / float(w)) * h)
            img = img.resize((LOGO_WIDTH, new_height), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.root.after(0, self.update_ui_with_logo, photo)
        except Exception as e:
            logger.error(f"Error loading logo: {e}")
    
    def update_ui_with_logo(self, photo):
        self.logo_label.config(image=photo)
        self.logo_label.image = photo

if __name__ == "__main__":
    root = tk.Tk()
    app = FullscreenImageApp(root)
    root.mainloop()
