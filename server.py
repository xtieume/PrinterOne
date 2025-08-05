#!/usr/bin/env python3
"""
PrinterOne - Unified Server and GUI Application
A comprehensive TCP print server with integrated GUI management and test client functionality
"""

# Critical startup logging - Log everything from the very beginning
import os
import sys
import time
import json
import socket
import threading
import subprocess
import signal
import tempfile
import logging
import glob
import psutil
import winreg
import traceback
from datetime import datetime, timedelta

# Setup early logging to capture startup issues
def setup_early_logging():
    """Setup logging as early as possible to capture startup issues"""
    try:
        # Create logs directory if it doesn't exist - use user temp if permission denied
        logs_dir = "logs"
        try:
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)
        except PermissionError:
            # Fallback to user temp directory if permission denied
            import tempfile
            logs_dir = os.path.join(tempfile.gettempdir(), "PrinterOne", "logs")
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir, exist_ok=True)
        
        # Generate startup log filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        startup_log_filename = f"{timestamp}_startup.log"
        startup_log_path = os.path.join(logs_dir, startup_log_filename)
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(startup_log_path, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        logger = logging.getLogger('startup')
        logger.info("=== PrinterOne Startup Log ===")
        logger.info(f"Python version: {sys.version}")
        logger.info(f"Platform: {sys.platform}")
        logger.info(f"Current working directory: {os.getcwd()}")
        logger.info(f"Script path: {os.path.abspath(__file__)}")
        logger.info(f"Command line arguments: {sys.argv}")
        logger.info(f"Environment variables: USERPROFILE={os.environ.get('USERPROFILE', 'NOT_SET')}")
        
        return logger
    except Exception as e:
        print(f"CRITICAL: Failed to setup early logging: {e}")
        print(f"Exception type: {type(e).__name__}")
        print(f"Traceback: {traceback.format_exc()}")
        return None

# Initialize early logging
startup_logger = setup_early_logging()

try:
    if startup_logger:
        startup_logger.info("Starting import phase...")
    
    # Import GUI modules
    try:
        if startup_logger:
            startup_logger.info("Importing GUI modules...")
        import tkinter as tk
        from tkinter import ttk
        if startup_logger:
            startup_logger.info("GUI modules imported successfully")
    except ImportError as e:
        if startup_logger:
            startup_logger.error(f"Failed to import GUI modules: {e}")
        raise
    
    # Import Windows-specific modules
    try:
        if startup_logger:
            startup_logger.info("Importing Windows print modules...")
        import win32print
        if startup_logger:
            startup_logger.info("Windows print modules imported successfully")
    except ImportError as e:
        if startup_logger:
            startup_logger.error(f"Failed to import Windows modules: {e}")
        raise
    
    # Import PDF generation
    try:
        if startup_logger:
            startup_logger.info("Importing PDF modules...")
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        if startup_logger:
            startup_logger.info("PDF modules imported successfully")
    except ImportError as e:
        if startup_logger:
            startup_logger.error(f"Failed to import PDF modules: {e}")
        raise

except Exception as e:
    error_msg = f"CRITICAL: Import phase failed: {e}"
    if startup_logger:
        startup_logger.critical(error_msg)
        startup_logger.critical(f"Exception type: {type(e).__name__}")
        startup_logger.critical(f"Traceback: {traceback.format_exc()}")
    else:
        print(error_msg)
        print(f"Exception type: {type(e).__name__}")
        print(f"Traceback: {traceback.format_exc()}")
    sys.exit(1)

# System tray imports (optional)
try:
    import pystray
    from PIL import Image, ImageTk
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False
    print("pystray not available, system tray disabled")

# Global variables
SERVER_RUNNING = True
AUTO_START_MODE = False

class PrinterOneServer:
    """PrinterOne TCP Server"""
    
    def __init__(self, log_callback=None):
        try:
            if startup_logger:
                startup_logger.info("Initializing PrinterOneServer...")
            
            self.config = self.load_config()
            
            if startup_logger:
                startup_logger.info(f"Configuration loaded: {self.config}")
            
            self.server_socket = None
            self.server_thread = None
            self.running = False
            self.log_callback = log_callback  # Callback function for logging to GUI
            
            if startup_logger:
                startup_logger.info("PrinterOneServer initialized successfully")
                
        except Exception as e:
            error_msg = f"Failed to initialize PrinterOneServer: {e}"
            if startup_logger:
                startup_logger.critical(error_msg)
                startup_logger.critical(f"Traceback: {traceback.format_exc()}")
            raise
    
    def log(self, message):
        """Log message to console and GUI if callback is set"""
        print(message)  # Always print to console
        if self.log_callback:
            # Remove the timestamp and brackets from message for GUI (GUI adds its own)
            clean_message = message
            if message.startswith("[") and "]" in message:
                # Extract message after first "]"
                bracket_end = message.find("]")
                if bracket_end != -1:
                    clean_message = message[bracket_end + 1:].strip()
            self.log_callback(clean_message)
    
    def load_config(self):
        """Load configuration from config.json"""
        default_config = {
            "printer_name": "",
            "port": 9100,
            "use_pdf_conversion": True,
            "save_pdf_file": False,
            "auto_start": False,
            "service_name": "PrinterOne",
            "service_description": "PrinterOne - Network print server for raw print data",
            "manual": False,
            "minimize_to_tray": True
        }
        
        # Try multiple config file locations
        config_paths = [
            'config.json',  # Current directory first
            os.path.join(os.path.expanduser('~'), 'PrinterOne', 'config.json'),  # User home
            os.path.join(os.environ.get('APPDATA', ''), 'PrinterOne', 'config.json'),  # AppData
            os.path.join(tempfile.gettempdir(), 'PrinterOne', 'config.json')  # Temp directory
        ]
        
        # Store the successful config path for saving
        self.config_path = None
        
        try:
            for config_path in config_paths:
                try:
                    config_path = os.path.abspath(config_path)
                    if startup_logger:
                        startup_logger.info(f"Trying to load configuration from: {config_path}")
                    
                    if os.path.exists(config_path):
                        if startup_logger:
                            startup_logger.info(f"Config file exists at: {config_path}")
                        
                        with open(config_path, 'r') as f:
                            config = json.load(f)
                            
                        if startup_logger:
                            startup_logger.info(f"Config loaded from file: {config}")
                        
                        # Merge with defaults
                        for key, value in default_config.items():
                            if key not in config:
                                config[key] = value
                                
                        if startup_logger:
                            startup_logger.info(f"Final merged config: {config}")
                        
                        self.config_path = config_path  # Remember successful path
                        return config
                except PermissionError:
                    if startup_logger:
                        startup_logger.warning(f"Permission denied accessing: {config_path}")
                    continue
                except Exception as e:
                    if startup_logger:
                        startup_logger.warning(f"Error loading from {config_path}: {e}")
                    continue
            
            # No config file found or accessible
            if startup_logger:
                startup_logger.info("No config.json found or accessible, using defaults")
                    
        except Exception as e:
            error_msg = f"Error in config loading process: {e}"
            if startup_logger:
                startup_logger.error(error_msg)
                startup_logger.error(f"Traceback: {traceback.format_exc()}")
            self.log(f"[!] {error_msg}")
        
        if startup_logger:
            startup_logger.info(f"Using default config: {default_config}")
        
        # Set a fallback config path for saving
        if not self.config_path:
            # Try to create a writable config directory
            for base_path in [os.path.expanduser('~'), os.environ.get('APPDATA', ''), tempfile.gettempdir()]:
                try:
                    config_dir = os.path.join(base_path, 'PrinterOne')
                    if not os.path.exists(config_dir):
                        os.makedirs(config_dir, exist_ok=True)
                    self.config_path = os.path.join(config_dir, 'config.json')
                    break
                except:
                    continue
            
            if not self.config_path:
                self.config_path = os.path.join(tempfile.gettempdir(), 'PrinterOne_config.json')
        
        return default_config
    
    def save_config(self, printer_name=None, port=None, use_pdf_conversion=None, save_pdf_file=None):
        """Save configuration to config.json"""
        try:
            if printer_name is not None:
                self.config["printer_name"] = printer_name
            if port is not None:
                self.config["port"] = port
            if use_pdf_conversion is not None:
                self.config["use_pdf_conversion"] = use_pdf_conversion
            if save_pdf_file is not None:
                self.config["save_pdf_file"] = save_pdf_file
            
            self.config["manual"] = True
            
            # Use the config path determined during load, or try fallback locations
            config_paths_to_try = []
            
            if hasattr(self, 'config_path') and self.config_path:
                config_paths_to_try.append(self.config_path)
            
            # Fallback locations in order of preference
            config_paths_to_try.extend([
                'config.json',  # Current directory
                os.path.join(os.path.expanduser('~'), 'PrinterOne', 'config.json'),  # User home
                os.path.join(os.environ.get('APPDATA', ''), 'PrinterOne', 'config.json'),  # AppData
                os.path.join(tempfile.gettempdir(), 'PrinterOne', 'config.json')  # Temp directory
            ])
            
            # Try to save to each location until one succeeds
            for config_path in config_paths_to_try:
                try:
                    # Ensure directory exists
                    config_dir = os.path.dirname(config_path)
                    if config_dir and not os.path.exists(config_dir):
                        os.makedirs(config_dir, exist_ok=True)
                    
                    # Try to save
                    with open(config_path, 'w') as f:
                        json.dump(self.config, f, indent=4)
                    
                    self.log(f"[SAVE] Configuration saved to {config_path}")
                    self.config_path = config_path  # Remember successful path
                    return True
                    
                except PermissionError:
                    self.log(f"[WARN] Permission denied saving to {config_path}, trying next location...")
                    continue
                except Exception as e:
                    self.log(f"[WARN] Error saving to {config_path}: {e}, trying next location...")
                    continue
            
            # If all locations failed
            self.log(f"[!] Failed to save configuration to any location")
            return False
            
        except Exception as e:
            self.log(f"[!] Error in save_config: {e}")
            return False
    
    def list_printers(self):
        """List all available printers"""
        try:
            printers = []
            for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1):
                printers.append(printer[2])
            return printers
        except Exception as e:
            self.log(f"[!] Error listing printers: {e}")
            return []
    
    def convert_raw_to_pdf(self, raw_data, save_file=False):
        """Convert raw data to PDF for testing with PDF printers (test client only)"""
        try:
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            
            self.log(f"[INFO] Creating PDF at: {temp_pdf_path}")
            
            c = canvas.Canvas(temp_pdf_path, pagesize=letter)
            data_format = self.analyze_raw_data(raw_data)
            
            # Add title
            c.setFont("Helvetica-Bold", 16)
            c.drawString(100, 750, "PrinterOne - Test Print Job")
            
            # Add info
            c.setFont("Helvetica", 12)
            c.drawString(100, 720, f"Data length: {len(raw_data)} bytes")
            c.drawString(100, 700, f"Data format: {data_format}")
            c.drawString(100, 680, f"Received at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            y_position = 640
            
            # Try to extract and display readable text first
            try:
                text_content = self.extract_readable_text(raw_data)
                if text_content and len(text_content.strip()) > 0:
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(100, y_position, "Content:")
                    y_position -= 30
                    
                    # Display text content
                    c.setFont("Helvetica", 11)
                    lines = text_content.split('\n')
                    
                    for line in lines:
                        if y_position < 80:
                            c.showPage()  # New page
                            y_position = 750
                        
                        # Wrap long lines
                        if len(line) > 80:
                            for i in range(0, len(line), 80):
                                chunk = line[i:i+80]
                                c.drawString(100, y_position, chunk)
                                y_position -= 15
                                if y_position < 80:
                                    c.showPage()
                                    y_position = 750
                        else:
                            c.drawString(100, y_position, line)
                            y_position -= 15
                    
                    y_position -= 20
                else:
                    # If no readable text, show hex and ASCII as before
                    self.add_hex_dump_to_pdf(c, raw_data, y_position)
                    
            except Exception as e:
                self.log(f"[WARN] Error extracting text, using hex dump: {e}")
                self.add_hex_dump_to_pdf(c, raw_data, y_position)
            
            c.save()
            
            # Read the PDF file
            with open(temp_pdf_path, 'rb') as f:
                pdf_data = f.read()
            
            # Clean up or save file
            if save_file:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                saved_path = f"raw_data_{timestamp}.pdf"
                os.rename(temp_pdf_path, saved_path)
                self.log(f"[SAVE] PDF saved as: {saved_path}")
            else:
                try:
                    os.unlink(temp_pdf_path)
                except:
                    pass
            
            return pdf_data
        except Exception as e:
            self.log(f"[!] PDF conversion error: {e}")
            return None
    
    def extract_readable_text(self, raw_data):
        """Extract readable text from raw data"""
        try:
            # Try UTF-8 first
            try:
                text = raw_data.decode('utf-8')
                # Clean up control characters but keep printable ones
                cleaned = ''.join(char if char.isprintable() or char in '\n\r\t' else ' ' for char in text)
                return cleaned.strip()
            except UnicodeDecodeError:
                pass
            
            # Try Windows-1252 (common in Windows printing)
            try:
                text = raw_data.decode('windows-1252', errors='ignore')
                cleaned = ''.join(char if char.isprintable() or char in '\n\r\t' else ' ' for char in text)
                return cleaned.strip()
            except:
                pass
            
            # Try ASCII with error handling
            try:
                text = raw_data.decode('ascii', errors='ignore')
                cleaned = ''.join(char if char.isprintable() or char in '\n\r\t' else ' ' for char in text)
                return cleaned.strip()
            except:
                pass
            
            return None
        except Exception as e:
            self.log(f"[WARN] Text extraction error: {e}")
            return None
    
    def add_hex_dump_to_pdf(self, canvas_obj, raw_data, start_y):
        """Add hex dump to PDF (fallback method)"""
        try:
            y_position = start_y
            
            # Add hex dump section
            canvas_obj.setFont("Helvetica-Bold", 12)
            canvas_obj.drawString(100, y_position, "Raw Data (Hex):")
            y_position -= 30
            
            # Add raw data as hex (first 2000 bytes)
            canvas_obj.setFont("Courier", 8)
            hex_data = raw_data[:2000].hex()
            
            # Split hex data into lines
            for i in range(0, len(hex_data), 80):
                if y_position < 50:
                    canvas_obj.showPage()
                    y_position = 750
                
                line = hex_data[i:i+80]
                formatted_line = ' '.join([line[j:j+2] for j in range(0, len(line), 2)])
                canvas_obj.drawString(100, y_position, formatted_line)
                y_position -= 12
            
            # Add ASCII representation
            y_position -= 20
            if y_position < 100:
                canvas_obj.showPage()
                y_position = 750
            
            canvas_obj.setFont("Helvetica-Bold", 12)
            canvas_obj.drawString(100, y_position, "ASCII representation:")
            y_position -= 20
            
            canvas_obj.setFont("Courier", 8)
            ascii_data = raw_data[:1000]
            ascii_line = ''
            for i, byte in enumerate(ascii_data):
                if 32 <= byte <= 126:
                    ascii_line += chr(byte)
                else:
                    ascii_line += '.'
                
                if (i + 1) % 80 == 0:
                    if y_position < 50:
                        canvas_obj.showPage()
                        y_position = 750
                    canvas_obj.drawString(100, y_position, ascii_line)
                    y_position -= 12
                    ascii_line = ''
            
            if ascii_line:
                canvas_obj.drawString(100, y_position, ascii_line)
                
        except Exception as e:
            self.log(f"[WARN] Hex dump error: {e}")
    
    def analyze_raw_data(self, data):
        """Analyze raw data to determine format"""
        if len(data) == 0:
            return "Empty data"
        
        # Check for common printer command formats
        if data.startswith(b'\x1b'):
            return "ESC/P (Epson)"
        elif data.startswith(b'\x1b%-12345X'):
            return "PCL (HP)"
        elif data.startswith(b'%!PS'):
            return "PostScript"
        elif data.startswith(b'\x02'):
            return "ZPL (Zebra)"
        elif b'PDF' in data[:100]:
            return "PDF document"
        elif b'Microsoft Office' in data or b'Word' in data or b'.docx' in data or b'.doc' in data:
            return "Microsoft Office document"
        elif b'%PDF' in data[:100]:
            return "PDF format"
        else:
            # Try to detect if it contains printable text
            try:
                decoded = data.decode('utf-8', errors='ignore')
                if len(decoded.strip()) > 0 and any(c.isprintable() and c not in '\r\n\t' for c in decoded[:200]):
                    return f"Text document ({len(data)} bytes)"
            except:
                pass
            
            return f"Binary/Unknown format ({len(data)} bytes)"
    
    def print_raw(self, data, printer_name):
        """Send raw data to printer"""
        try:
            self.log(f"[INFO] Opening printer: {printer_name}")
            hPrinter = win32print.OpenPrinter(printer_name)
            
            job_info = ("RAW Print Job", None, "RAW")
            hJob = win32print.StartDocPrinter(hPrinter, 1, job_info)
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, data)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            
            self.log(f"[OK] Successfully printed {len(data)} bytes.")
            return True
        except Exception as e:
            self.log(f"[!] Print error: {e}")
            return False
    
    
    def handle_client(self, client_socket, address):
        """Handle a client connection"""
        self.log(f"[CONN] Client connected: {address}")
        try:
            data = b""
            while True:
                chunk = client_socket.recv(4096)
                if not chunk:
                    break
                data += chunk
            
            if data:
                self.log(f"[DATA] Received {len(data)} bytes from {address}")
                self.log(f"[INFO] Data format: {self.analyze_raw_data(data)}")
                printer_name = self.config.get("printer_name", "")
                if printer_name:
                    self.print_raw(data, printer_name)
                else:
                    self.log(f"[!] No printer configured")
            else:
                self.log(f"[!] No data received from {address}")
                
        except Exception as e:
            self.log(f"[!] Error handling client {address}: {e}")
        finally:
            client_socket.close()
            self.log(f"[CONN] Client disconnected: {address}")
    
    def kill_process_on_port(self, port):
        """Kill any process using the specified port"""
        try:
            # Use psutil instead of netstat to avoid snmpapi.dll dependency
            for conn in psutil.net_connections():
                if conn.laddr.port == port and conn.status == psutil.CONN_LISTEN:
                    try:
                        process = psutil.Process(conn.pid)
                        process.terminate()
                        self.log(f"[KILL] Terminated process {conn.pid} ({process.name()}) using port {port}")
                        time.sleep(1)
                        
                        # Force kill if still running
                        if process.is_running():
                            process.kill()
                            self.log(f"[KILL] Force killed process {conn.pid}")
                    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                        self.log(f"[!] Failed to kill process {conn.pid}: {e}")
        except Exception as e:
            self.log(f"[!] Error killing process on port {port}: {e}")
    
    def get_local_ip(self):
        """Get the actual local IP address of the machine"""
        try:
            # Method 1: Try to connect to a remote server to determine local IP
            test_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            try:
                # Connect to Google DNS (doesn't actually send data)
                test_socket.connect(("8.8.8.8", 80))
                local_ip = test_socket.getsockname()[0]
                test_socket.close()
                
                # Validate that it's not a loopback address or VirtualBox
                if (not local_ip.startswith('127.') and 
                    not local_ip.startswith('192.168.56.') and  # VirtualBox Host-Only
                    not local_ip.startswith('169.254.')):       # APIPA
                    return local_ip
            except:
                test_socket.close()
            
            # Method 2: Use psutil to get network interfaces with better filtering
            try:
                import psutil
                interfaces_with_gw = []
                interfaces_without_gw = []
                
                # Get default gateways to identify primary interfaces
                gateways = psutil.net_if_stats()
                
                for interface_name, interface_addresses in psutil.net_if_addrs().items():
                    # Skip known virtual interfaces
                    if any(skip in interface_name.lower() for skip in [
                        'virtualbox', 'vmware', 'vbox', 'hyper-v', 'loopback', 
                        'bluetooth', 'isatap', 'teredo', 'tunnel'
                    ]):
                        continue
                    
                    for address in interface_addresses:
                        if address.family == socket.AF_INET:
                            ip = address.address
                            
                            # Skip loopback, APIPA, and VirtualBox IPs
                            if (ip.startswith('127.') or 
                                ip.startswith('169.254.') or
                                ip.startswith('192.168.56.')):  # VirtualBox Host-Only
                                continue
                            
                            # Check if this interface is up and running
                            try:
                                interface_stats = psutil.net_if_stats().get(interface_name)
                                if interface_stats and interface_stats.isup:
                                    # Prefer Wi-Fi and Ethernet over other interfaces
                                    if any(pref in interface_name.lower() for pref in ['wi-fi', 'wifi', 'ethernet', 'local area']):
                                        interfaces_with_gw.append((ip, interface_name))
                                    else:
                                        interfaces_without_gw.append((ip, interface_name))
                            except:
                                pass
                
                # Return the best interface
                if interfaces_with_gw:
                    # Prefer Wi-Fi over Ethernet if both available
                    for ip, name in interfaces_with_gw:
                        if 'wi-fi' in name.lower() or 'wifi' in name.lower():
                            return ip
                    # Otherwise return first good interface
                    return interfaces_with_gw[0][0]
                
                if interfaces_without_gw:
                    return interfaces_without_gw[0][0]
                    
            except Exception as e:
                if startup_logger:
                    startup_logger.warning(f"Method 2 failed: {e}")
            
            # Method 3: Fallback to hostname resolution
            try:
                hostname = socket.gethostname()
                local_ip = socket.gethostbyname(hostname)
                if (not local_ip.startswith('127.') and 
                    not local_ip.startswith('192.168.56.')):
                    return local_ip
            except:
                pass
            
            # Method 4: Last resort - return localhost
            return '127.0.0.1'
            
        except Exception as e:
            if startup_logger:
                startup_logger.error(f"Error getting local IP: {e}")
            return '127.0.0.1'
    
    def start_server(self):
        """Start the TCP print server"""
        global SERVER_RUNNING
        SERVER_RUNNING = True
        
        printer_name = self.config.get("printer_name", "")
        port = self.config.get("port", 9100)
        
        if not printer_name:
            self.log("[!] No printer configured!")
            return False
        
        # Kill any process using the port first
        self.log(f"[KILL] Checking for processes using port {port}...")
        self.kill_process_on_port(port)
        
        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.server_socket.settimeout(1.0)
        
        try:
            self.server_socket.bind(('0.0.0.0', port))
            self.server_socket.listen(5)
            self.running = True
            
            self.log(f"[OK] Server started on port {port}")
            self.log(f"[PRINTER] Using printer: {printer_name}")
            
            # Get local IP addresses using improved method
            local_ip = self.get_local_ip()
            self.log(f"[IP] Local IP: {local_ip}")
            self.log(f"[CONNECT] Other machines can connect to: {local_ip}:{port}")
            
            while SERVER_RUNNING and self.running:
                try:
                    client_socket, address = self.server_socket.accept()
                    client_thread = threading.Thread(
                        target=self.handle_client, 
                        args=(client_socket, address)
                    )
                    client_thread.daemon = True
                    client_thread.start()
                except socket.timeout:
                    continue
                except Exception as e:
                    if self.running:
                        self.log(f"[!] Error accepting client: {e}")
                    break
                    
        except Exception as e:
            self.log(f"[!] Server error: {e}")
            return False
        finally:
            self.stop_server()
        
        return True
    
    def stop_server(self):
        """Stop the TCP print server"""
        global SERVER_RUNNING
        SERVER_RUNNING = False
        self.running = False
        
        if self.server_socket:
            try:
                self.server_socket.close()
            except:
                pass
            self.server_socket = None
        
        self.log("[DONE] Server stopped")

class TestClient:
    """Test client for the print server"""
    
    @staticmethod
    def test_connection(host='localhost', port=9100, test_data=None, log_callback=None):
        """Test connection to print server"""
        def log(message):
            print(message)  # Always print to console
            if log_callback:
                log_callback(message)
        
        try:
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(5)
            
            log(f"[CONNECT] Connecting to {host}:{port}...")
            client_socket.connect((host, port))
            log(f"[OK] Connected successfully!")
            
            # Send test data if provided
            if test_data is not None:
                log(f"[SEND] Sending {len(test_data)} bytes...")
                client_socket.send(test_data)
                log(f"[OK] Data sent successfully!")
            else:
                log(f"[OK] Connection test only - no data sent")
            
            client_socket.close()
            return True
            
        except ConnectionRefusedError:
            log(f"[ERROR] Connection refused. Is the server running on {host}:{port}?")
            return False
        except socket.timeout:
            log(f"[ERROR] Connection timeout.")
            return False
        except Exception as e:
            log(f"[ERROR] Error: {e}")
            return False

class AutoStartManager:
    """Windows auto-start management"""
    
    @staticmethod  
    def find_manager_exe():
        """Find PrinterOne Manager GUI executable"""
        # Check if running from exe (PyInstaller)
        if hasattr(sys, '_MEIPASS'):
            # Running from exe - use sys.executable which points to exe
            exe_path = os.path.abspath(sys.executable)
            # For exe files, we need to include parameters as part of the command
            return f'"{exe_path}" gui auto_start'
        else:
            # Running from Python script
            current_script = os.path.abspath(__file__)
            python_exe = os.path.abspath(sys.executable)
            return f'"{python_exe}" "{current_script}" gui auto_start'
    
    @staticmethod
    def add_to_startup():
        """Add PrinterOne Manager to Windows startup"""
        try:
            registry_path = AutoStartManager.find_manager_exe()
            
            # Add to Windows startup registry
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            winreg.SetValueEx(key, "PrinterOneManager", 0, winreg.REG_SZ, registry_path)
            winreg.CloseKey(key)
            
            return True, f"PrinterOne Manager added to Windows startup!"
            
        except Exception as e:
            return False, f"Error adding to startup: {e}"
    
    @staticmethod
    def remove_from_startup():
        """Remove PrinterOne Manager from Windows startup"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            
            winreg.DeleteValue(key, "PrinterOneManager")
            winreg.CloseKey(key)
            
            return True, "PrinterOne Manager removed from Windows startup!"
            
        except Exception as e:
            return False, f"Error removing from startup: {e}"
    
    @staticmethod
    def check_startup_status():
        """Check if PrinterOne Manager is in startup"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_READ
            )
            
            try:
                value, _ = winreg.QueryValueEx(key, "PrinterOneManager")
                winreg.CloseKey(key)
                return True, value
            except FileNotFoundError:
                winreg.CloseKey(key)
                return False, "Not in startup"
                
        except Exception as e:
            return False, f"Error checking startup status: {e}"

class PrinterOneGUI:
    """Integrated GUI for PrinterOne"""
    
    def __init__(self, root):
        # Setup GUI initialization logging
        self.init_logger = None
        try:
            self.init_logger = logging.getLogger('gui_init')
            if not self.init_logger.handlers:  # Avoid duplicate handlers
                self.init_logger.setLevel(logging.INFO)
                
                # Create logs directory if it doesn't exist - use temp fallback if permission denied
                logs_dir = "logs"
                try:
                    if not os.path.exists(logs_dir):
                        os.makedirs(logs_dir)
                except PermissionError:
                    # Fallback to user temp directory if permission denied
                    import tempfile
                    logs_dir = os.path.join(tempfile.gettempdir(), "PrinterOne", "logs")
                    if not os.path.exists(logs_dir):
                        os.makedirs(logs_dir, exist_ok=True)
                
                # Generate GUI initialization log filename
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                gui_init_log_filename = f"gui_init_{timestamp}.log"
                gui_init_log_path = os.path.join(logs_dir, gui_init_log_filename)
                
                # Create file handler for GUI initialization log
                file_handler = logging.FileHandler(gui_init_log_path, encoding='utf-8')
                file_handler.setLevel(logging.INFO)
                
                # Create formatter
                formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(formatter)
                
                # Add handler to logger
                self.init_logger.addHandler(file_handler)
        except Exception as e:
            print(f"Error setting up GUI init logging: {e}")
            self.init_logger = None
        
        try:
            if self.init_logger:
                self.init_logger.info("=== PrinterOneGUI Initialization Started ===")
                self.init_logger.info(f"Tkinter root object: {root}")
            
            self.root = root
            self.root.title("PrinterOne - Network Print Server")
            self.root.geometry("1200x700")
            self.root.resizable(True, True)
            
            if self.init_logger:
                self.init_logger.info("Basic root window configuration completed")
            
            # Initialize server with log callback
            if self.init_logger:
                self.init_logger.info("Initializing PrinterOneServer...")
            
            self.server = PrinterOneServer(log_callback=self.log_message)
            self.server_thread = None
            
            if self.init_logger:
                self.init_logger.info("PrinterOneServer initialized successfully")
            
            # GUI variables
            if self.init_logger:
                self.init_logger.info("Setting up GUI variables...")
            
            self.printer_var = tk.StringVar(value=self.server.config.get("printer_name", ""))
            self.port_var = tk.IntVar(value=self.server.config.get("port", 9100))
            self.test_host_var = tk.StringVar(value="localhost")
            self.test_port_var = tk.IntVar(value=9100)
            
            # System tray variables
            self.tray_icon = None
            self.minimize_to_tray = self.server.config.get("minimize_to_tray", True)
            self.minimize_to_tray_var = tk.BooleanVar(value=self.minimize_to_tray)
            
            if self.init_logger:
                self.init_logger.info("GUI variables setup completed")
            
            # Setup logging
            if self.init_logger:
                self.init_logger.info("Setting up application logging...")
            
            self.logger = self.setup_logging()
            
            if self.init_logger:
                self.init_logger.info("Application logging setup completed")
            
            # Set window icon
            if self.init_logger:
                self.init_logger.info("Setting window icon...")
            
            self.set_window_icon()
            
            if self.init_logger:
                self.init_logger.info("Window icon setup completed")
            
            # Create GUI
            if self.init_logger:
                self.init_logger.info("Creating GUI widgets...")
            
            self.create_widgets()
            
            if self.init_logger:
                self.init_logger.info("GUI widgets creation completed")
            
            # Update status
            if self.init_logger:
                self.init_logger.info("Updating initial status...")
            
            self.update_status()
            
            if self.init_logger:
                self.init_logger.info("Initial status update completed")
            
            # Bind window events
            if self.init_logger:
                self.init_logger.info("Binding window events...")
            
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            if self.init_logger:
                self.init_logger.info("Window events binding completed")
            
            # Start status update thread
            if self.init_logger:
                self.init_logger.info("Starting status update thread...")
            
            self.start_status_thread()
            
            if self.init_logger:
                self.init_logger.info("Status update thread started")
            
            # Setup system tray
            if self.init_logger:
                self.init_logger.info(f"Setting up system tray (TRAY_AVAILABLE: {TRAY_AVAILABLE})...")
            
            if TRAY_AVAILABLE:
                self.setup_tray()
                if self.init_logger:
                    self.init_logger.info("System tray setup completed")
            else:
                if self.init_logger:
                    self.init_logger.info("System tray not available, skipping")
            
            # Auto-start server if configured
            if self.init_logger:
                self.init_logger.info(f"Checking auto-start configuration (AUTO_START_MODE: {AUTO_START_MODE})...")
            
            if AUTO_START_MODE:
                if self.init_logger:
                    self.init_logger.info("Auto-start mode detected, hiding window to system tray")
                # Hide window to system tray in auto-start mode
                if TRAY_AVAILABLE:
                    self.root.after(100, self.hide_window)
                if self.init_logger:
                    self.init_logger.info("Scheduling server start in 2 seconds")
                self.root.after(2000, self.auto_start_server)
            else:
                # Auto-start server if printer is configured
                printer_name = self.server.config.get("printer_name", "")
                if printer_name and printer_name.strip():
                    if self.init_logger:
                        self.init_logger.info(f"Printer configured ({printer_name}), scheduling auto-start in 1 second")
                    self.log_message("Printer configured, auto-starting server...")
                    self.root.after(1000, self.auto_start_server)
                else:
                    if self.init_logger:
                        self.init_logger.info("No printer configured, server will not auto-start")
            
            if self.init_logger:
                self.init_logger.info("=== PrinterOneGUI Initialization Completed Successfully ===")
                
        except Exception as e:
            error_msg = f"Error during GUI initialization: {e}"
            print(error_msg)
            if self.init_logger:
                self.init_logger.critical(error_msg)
                self.init_logger.critical(f"Exception type: {type(e).__name__}")
                import traceback
                self.init_logger.critical(f"Traceback: {traceback.format_exc()}")
            
            # Re-raise to maintain original behavior
            raise
    
    def get_resource_path(self, filename):
        """Get absolute path to resource"""
        try:
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, filename)
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        except Exception:
            return os.path.abspath(filename)
    
    def set_window_icon(self):
        """Set window icon"""
        try:
            icon_png = self.get_resource_path("printer.png") 
            if os.path.exists(icon_png):
                img = Image.open(icon_png)
                img = img.resize((32, 32), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo)
        except Exception as e:
            print(f"Error setting window icon: {e}")
    
    def create_widgets(self):
        """Create GUI widgets"""
        # Main notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Server Management Tab
        server_frame = ttk.Frame(notebook)
        notebook.add(server_frame, text="Server Management")
        self.create_server_tab(server_frame)
        
        # Test Client Tab
        test_frame = ttk.Frame(notebook)
        notebook.add(test_frame, text="Test Client")
        self.create_test_tab(test_frame)
        
        # Settings Tab
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="Settings")
        self.create_settings_tab(settings_frame)
    
    def create_server_tab(self, parent):
        """Create server management tab"""
        # Top section - Configuration and Control
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Configuration frame
        config_frame = ttk.LabelFrame(top_frame, text="Configuration", padding="10")
        config_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Printer selection
        ttk.Label(config_frame, text="Printer:").pack(anchor=tk.W)
        printer_combo = ttk.Combobox(config_frame, textvariable=self.printer_var, width=40)
        printer_combo['values'] = self.server.list_printers()
        printer_combo.pack(fill=tk.X, pady=(5, 10))
        
        # Port configuration
        ttk.Label(config_frame, text="Port:").pack(anchor=tk.W)
        port_entry = ttk.Entry(config_frame, textvariable=self.port_var, width=10)
        port_entry.pack(anchor=tk.W, pady=(5, 10))
        
        # Save config button
        ttk.Button(config_frame, text="Save Configuration", 
                  command=self.save_configuration).pack(fill=tk.X)
        
        # Control frame
        control_frame = ttk.LabelFrame(top_frame, text="Server Control", padding="10")
        control_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
        
        # Server status
        self.server_status_label = ttk.Label(control_frame, text="[STOP] Server Stopped", 
                                           font=("Arial", 12, "bold"))
        self.server_status_label.pack(pady=10)
        
        self.server_info_label = ttk.Label(control_frame, text="", font=("Arial", 9))
        self.server_info_label.pack(pady=5)
        
        # Control buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(pady=10)
        
        self.start_button = ttk.Button(button_frame, text="Start Server", 
                                      command=self.start_server, width=15)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Stop Server", 
                                     command=self.stop_server, width=15, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Auto-start section
        autostart_frame = ttk.LabelFrame(control_frame, text="Auto-Start", padding="10")
        autostart_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.autostart_status_label = ttk.Label(autostart_frame, text="Checking...")
        self.autostart_status_label.pack()
        
        autostart_button_frame = ttk.Frame(autostart_frame)
        autostart_button_frame.pack(pady=5)
        
        self.add_autostart_button = ttk.Button(autostart_button_frame, text="Add", 
                                              command=self.add_to_startup, width=12)
        self.add_autostart_button.pack(side=tk.LEFT, padx=2)
        
        self.remove_autostart_button = ttk.Button(autostart_button_frame, text="Remove", 
                                                 command=self.remove_from_startup, width=12)
        self.remove_autostart_button.pack(side=tk.LEFT, padx=2)
        
        # Log section
        log_frame = ttk.LabelFrame(parent, text="Server Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Create text widget with scrollbar
        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_text_frame, height=15, font=("Consolas", 9))
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_test_tab(self, parent):
        """Create test client tab"""
        # Test configuration frame
        test_config_frame = ttk.LabelFrame(parent, text="Test Configuration", padding="10")
        test_config_frame.pack(fill=tk.X, padx=10, pady=10)
        
        config_grid = ttk.Frame(test_config_frame)
        config_grid.pack(fill=tk.X)
        
        # Host/Server input
        ttk.Label(config_grid, text="Server:").grid(row=0, column=0, sticky=tk.W, pady=5)
        host_entry = ttk.Entry(config_grid, textvariable=self.test_host_var, width=20)
        host_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 20), pady=5)
        
        # Port input
        ttk.Label(config_grid, text="Port:").grid(row=0, column=2, sticky=tk.W, pady=5)
        port_entry = ttk.Entry(config_grid, textvariable=self.test_port_var, width=10)
        port_entry.grid(row=0, column=3, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Test button
        ttk.Button(config_grid, text="Test Connection", 
                  command=self.test_connection, width=15).grid(row=0, column=4, padx=(20, 0), pady=5)
        
        # Test data options
        test_data_frame = ttk.LabelFrame(parent, text="Test Data", padding="10")
        test_data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Test data button
        button_frame = ttk.Frame(test_data_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="Send Test Data", 
                  command=lambda: self.send_test_data("test")).pack(side=tk.LEFT, padx=5)
        
        # Test log area
        log_frame = ttk.LabelFrame(test_data_frame, text="Test Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        # Test log text widget with scrollbar
        test_log_container = ttk.Frame(log_frame)
        test_log_container.pack(fill=tk.BOTH, expand=True)
        
        self.test_log_text = tk.Text(test_log_container, height=8, font=("Consolas", 9), wrap=tk.WORD)
        test_log_scrollbar = ttk.Scrollbar(test_log_container, orient="vertical", command=self.test_log_text.yview)
        self.test_log_text.configure(yscrollcommand=test_log_scrollbar.set)
        
        self.test_log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        test_log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Initialize with instruction text
        instruction_text = """Test Log Area
=============

Click 'Test Connection' to ping the server.
Click 'Send Test Data' to send a test print job.

Test results will appear here...
"""
        
        self.test_log_text.insert("1.0", instruction_text)
    
    def create_settings_tab(self, parent):
        """Create settings tab"""
        # Application settings
        app_frame = ttk.LabelFrame(parent, text="Application Settings", padding="10")
        app_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Minimize to tray option
        minimize_check = ttk.Checkbutton(app_frame, 
                                       text="Minimize to system tray when closing window",
                                       variable=self.minimize_to_tray_var,
                                       command=self.on_minimize_option_changed)
        minimize_check.pack(anchor=tk.W, pady=5)
        
        if not TRAY_AVAILABLE:
            minimize_check.config(state="disabled")
            ttk.Label(app_frame, text="(System tray not available - pystray not installed)",
                     font=("Arial", 8), foreground="gray").pack(anchor=tk.W)
        
        # About section
        about_frame = ttk.LabelFrame(parent, text="About", padding="10")
        about_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        about_text = """PrinterOne - Network Print Server
Version 1.0
Copyright (c) 2025 xtieume@gmail.com
GitHub: https://github.com/xtieume/PrinterOne

A simple TCP server that receives raw print data and sends it directly to a local printer.
Includes a GUI management interface and test client with PDF conversion for testing."""
        
        ttk.Label(about_frame, text=about_text, justify=tk.LEFT, font=("Arial", 9)).pack(anchor=tk.W)
    
    def log_test_message(self, message):
        """Add message to test log"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Add to test log
        self.test_log_text.insert(tk.END, log_entry)  
        self.test_log_text.see(tk.END)
        self.test_log_text.update_idletasks()
        
        # Limit test log size
        if int(self.test_log_text.index('end-1c').split('.')[0]) > 100:
            self.test_log_text.delete('1.0', '10.0')
    
    def log_message(self, message):
        """Add message to log"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Add to GUI log
        self.log_text.insert(tk.END, log_entry)  
        self.log_text.see(tk.END)
        self.log_text.update_idletasks()
        
        # Also log to file
        if hasattr(self, 'logger'):
            self.logger.info(message)
        
        # Limit GUI log size
        if int(self.log_text.index('end-1c').split('.')[0]) > 1000:
            self.log_text.delete('1.0', '100.0')
    
    def save_configuration(self):
        """Save the current configuration"""
        printer_name = self.printer_var.get()
        port = self.port_var.get()
        
        if not printer_name:
            self.log_message("[WARN] Please select a printer first!")
            return
        
        if self.server.save_config(printer_name=printer_name, port=port):
            self.log_message("[OK] Configuration saved successfully!")
        else:
            self.log_message("[ERROR] Failed to save configuration!")
        
        # Update port in test client if not manually changed
        if self.test_port_var.get() == 9100 or self.test_port_var.get() == self.server.config.get("port", 9100):
            self.test_port_var.set(port)
    
    def start_server(self):
        """Start the print server"""
        if self.server_thread and self.server_thread.is_alive():
            self.log_message("[WARN] Server is already running!")
            return
        
        printer_name = self.printer_var.get()
        port = self.port_var.get()
        
        if not printer_name:
            self.log_message("[WARN] Please select a printer first!")
            return
        
        # Save current configuration
        if self.server.save_config(printer_name=printer_name, port=port):
            self.log_message("[OK] Configuration saved")
        
        # Start server in separate thread
        self.server_thread = threading.Thread(target=self.server.start_server, daemon=True)
        self.server_thread.start()
        
        self.log_message("[START] Starting server...")
        self.update_server_status()
        
        # Give server time to start
        self.root.after(1000, self.update_server_status)
    
    def stop_server(self):
        """Stop the print server"""
        self.server.stop_server()
        self.log_message("[STOP] Server stopped")
        self.update_server_status()
    
    def auto_start_server(self):
        """Auto-start server when launched from startup or when printer is configured"""
        try:
            printer_name = self.server.config.get("printer_name", "")
            
            if not printer_name or not printer_name.strip():
                if AUTO_START_MODE:
                    self.log_message("[WARN] Auto-start mode: No printer configured, running in system tray")
                else:
                    self.log_message("[INFO] No printer configured, server not started")
                return
            
            if AUTO_START_MODE:
                self.log_message("[AUTO] Auto-start mode: Starting server in background, check system tray...")
            else:
                self.log_message("[AUTO] Auto-starting server with configured printer...")
            
            # Start server automatically
            self.start_server()
        except Exception as e:
            self.log_message(f"[ERROR] Error in auto-start: {e}")
    
    def update_status(self):
        """Update all status displays"""
        self.update_server_status()
        self.update_autostart_status()
    
    def update_server_status(self):
        """Update server status display"""
        if self.server.running:
            self.server_status_label.config(text="[OK] Server Running", foreground="green")
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            
            port = self.server.config.get("port", 9100)
            try:
                local_ip = self.server.get_local_ip()
                info_text = f"Port: {port} | IP: {local_ip}"
            except:
                info_text = f"Port: {port}"
            
            self.server_info_label.config(text=info_text)
        else:
            self.server_status_label.config(text="[STOP] Server Stopped", foreground="red")
            self.start_button.config(state="normal")
            self.stop_button.config(state="disabled")
            self.server_info_label.config(text="")
    
    def update_autostart_status(self):
        """Update auto-start status"""
        is_in_startup, path_or_error = AutoStartManager.check_startup_status()
        
        if is_in_startup:
            self.autostart_status_label.config(text="[OK] Auto-start enabled", foreground="green")
            self.add_autostart_button.config(state="disabled")
            self.remove_autostart_button.config(state="normal")
        else:
            self.autostart_status_label.config(text="[STOP] Auto-start disabled", foreground="red")
            self.add_autostart_button.config(state="normal")
            self.remove_autostart_button.config(state="disabled")
    
    def add_to_startup(self):
        """Add to Windows startup"""
        success, message = AutoStartManager.add_to_startup()
        if success:
            self.log_message(f"[OK] {message}")
        else:
            self.log_message(f"[ERROR] {message}")
        self.update_autostart_status()
    
    def remove_from_startup(self):
        """Remove from Windows startup"""
        success, message = AutoStartManager.remove_from_startup()
        if success:
            self.log_message(f"[OK] {message}")
        else:
            self.log_message(f"[ERROR] {message}")
        self.update_autostart_status()
    
    def test_connection(self):
        """Test connection to server (ping only)"""
        host = self.test_host_var.get()
        port = self.test_port_var.get()
        
        self.log_test_message(f"[CONNECT] Testing connection to {host}:{port}...")
        
        def run_test():
            # Only test connection, don't send any data
            success = TestClient.test_connection(host, port, test_data=None, log_callback=self.log_test_message)
            if success:
                self.root.after(0, lambda: self.log_test_message("[OK] Connection test completed!"))
            else:
                self.root.after(0, lambda: self.log_test_message("[ERROR] Connection test failed!"))
        
        threading.Thread(target=run_test, daemon=True).start()
    
    def send_test_data(self, data_type):
        """Send test data to server"""
        host = self.test_host_var.get()
        port = self.test_port_var.get()
        
        # Prepare test data (default test data)
        test_data = b"""PrinterOne Test Data
====================

This is a test print job sent from PrinterOne test client.
Date: """ + time.strftime("%Y-%m-%d %H:%M:%S").encode() + b"""

Test content:
- Line 1: Testing printer functionality
- Line 2: Checking data transmission
- Line 3: Verifying print server operation
- Line 4: Testing raw data handling
- Line 5: End of test data

If you can see this printed output, the PrinterOne server is working correctly!
"""
        
        # Check if target printer is PDF printer and convert data if needed
        printer_name = self.server.config.get("printer_name", "")
        use_pdf_conversion = self.server.config.get("use_pdf_conversion", True)
        
        if printer_name == "Microsoft Print to PDF" and use_pdf_conversion:
            self.log_test_message(f"[PDF] Converting test data to PDF for PDF printer...")
            try:
                pdf_data = self.server.convert_raw_to_pdf(test_data, save_file=False)
                if pdf_data:
                    test_data = pdf_data
                    self.log_test_message(f"[OK] Test data converted to PDF ({len(test_data)} bytes)")
                else:
                    self.log_test_message("[WARN] PDF conversion failed, using raw data")
            except Exception as e:
                self.log_test_message(f"[WARN] PDF conversion error: {e}")
        
        self.log_test_message(f"[SEND] Sending test data to {host}:{port} ({len(test_data)} bytes)")
        
        def run_test():
            success = TestClient.test_connection(host, port, test_data, log_callback=self.log_test_message)
            if success:
                self.root.after(0, lambda: self.log_test_message("[OK] Test data sent successfully!"))
            else:
                self.root.after(0, lambda: self.log_test_message("[ERROR] Failed to send test data!"))
        
        threading.Thread(target=run_test, daemon=True).start()
    
    def start_status_thread(self):
        """Start thread to periodically update status"""
        def status_updater():
            while True:
                try:
                    self.root.after(0, self.update_server_status) 
                    time.sleep(2)
                except:
                    break
        
        threading.Thread(target=status_updater, daemon=True).start()
    
    def on_closing(self):
        """Handle window closing"""
        if TRAY_AVAILABLE and self.tray_icon and self.minimize_to_tray:
            self.hide_window()
        else:
            self.quit_app()
    
    def on_minimize_option_changed(self):
        """Handle minimize to tray option change"""
        self.minimize_to_tray = self.minimize_to_tray_var.get()
        # Save to config
        self.server.config["minimize_to_tray"] = self.minimize_to_tray
        self.server.save_config()
    
    def quit_app(self):
        """Quit the application"""
        try:
            self.log_message("[BYE] Shutting down...")
            
            # Stop server
            if self.server.running:
                self.log_message("[STOP] Stopping server...")
                self.server.stop_server()
                time.sleep(1)
            
            # Stop tray icon
            if TRAY_AVAILABLE and self.tray_icon:
                try:
                    self.tray_icon.stop()
                    self.log_message("[TRAY] Tray icon stopped")
                except:
                    pass
            
            self.log_message("[OK] Application closed")
            self.root.quit()
            self.root.destroy()
            sys.exit(0)
        except Exception as e:
            print(f"Error quitting app: {e}")
            sys.exit(0)
    
    def setup_tray(self):
        """Setup system tray icon"""
        if not TRAY_AVAILABLE:
            return
        
        try:
            # Load icon - try multiple paths
            tray_image = None
            
            # Try bundled resource first
            icon_path = self.get_resource_path("printer.png")
            try:
                tray_image = Image.open(icon_path)
            except Exception:
                pass
            
            # If bundled resource fails, try direct path
            if tray_image is None:
                try:
                    direct_path = "printer.png"
                    tray_image = Image.open(direct_path)
                except Exception:
                    pass
            
            # If all fails, create a default icon
            if tray_image is None:
                tray_image = Image.new('RGB', (64, 64), color='blue')
            
            # Create menu
            menu = pystray.Menu(
                pystray.MenuItem("Show Window", self.show_window, default=True),
                pystray.MenuItem("Hide Window", self.hide_window),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Start Server", self.start_server_tray),
                pystray.MenuItem("Stop Server", self.stop_server_tray),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self.quit_app)
            )
            
            self.tray_icon = pystray.Icon("PrinterOne", tray_image, "PrinterOne", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            
        except Exception as e:
            print(f"Error setting up tray: {e}")
    
    def show_window(self, icon=None, item=None):
        """Show the main window"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
    
    def hide_window(self, icon=None, item=None):
        """Hide window to tray"""
        self.root.withdraw()
    
    def start_server_tray(self, icon=None, item=None):
        """Start server from tray"""
        self.root.after(0, self.start_server)
    
    def stop_server_tray(self, icon=None, item=None):
        """Stop server from tray"""
        self.root.after(0, self.stop_server)
    
    def setup_logging(self):
        """Setup logging system"""
        # Create logs directory - use temp fallback if permission denied
        logs_dir = "logs"
        try:
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)
        except PermissionError:
            # Fallback to user temp directory if permission denied
            import tempfile
            logs_dir = os.path.join(tempfile.gettempdir(), "PrinterOne", "logs")
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir, exist_ok=True)
        
        # Clean old logs
        self.cleanup_old_logs(logs_dir)
        
        # Generate log filename with simple timestamp format
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"{timestamp}.log"
        log_path = os.path.join(logs_dir, log_filename)
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        return logging.getLogger(__name__)
    
    def cleanup_old_logs(self, logs_dir, days_to_keep=30):
        """Clean up old log files (retention: 30 days)"""
        try:
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            log_pattern = os.path.join(logs_dir, "*.log")
            
            for log_file in glob.glob(log_pattern):
                try:
                    file_time = datetime.fromtimestamp(os.path.getctime(log_file))
                    if file_time < cutoff_date:
                        os.remove(log_file)
                        print(f"Cleaned up old log: {os.path.basename(log_file)}")
                except Exception as e:
                    print(f"Error removing log file {log_file}: {e}")
        except Exception as e:
            print(f"Error during log cleanup: {e}")

def signal_handler(signum, frame):
    """Handle Ctrl+C signal"""
    global SERVER_RUNNING
    print(f"\n[STOP] Received signal {signum}, stopping...")
    SERVER_RUNNING = False
    sys.exit(0)

def run_console_mode():
    """Run in console mode (command line interface)"""
    # Set up signal handler
    signal.signal(signal.SIGINT, signal_handler)
    
    print("PrinterOne - Network Print Server")
    print("=" * 35)
    print()
    
    server = PrinterOneServer()
    
    # Check if configuration exists
    config = server.config
    if not config.get("printer_name"):
        print("No printer configured. Please configure first:")
        printers = server.list_printers()
        
        if not printers:
            print("[!] No printers found!")
            return
        
        print(f"Found {len(printers)} printers:")
        for i, printer in enumerate(printers, 1):
            print(f"  {i}. {printer}")
        
        while True:
            try:
                selection = input(f"\nSelect printer (1-{len(printers)}): ")
                printer_index = int(selection) - 1
                if 0 <= printer_index < len(printers):
                    selected_printer = printers[printer_index]
                    break
                else:
                    print("[!] Invalid selection.")
            except ValueError:
                print("[!] Please enter a valid number.")
        
        port_input = input("Enter port (default: 9100): ")
        port = int(port_input) if port_input.strip() else 9100
        
        server.save_config(printer_name=selected_printer, port=port)
        print(f"[OK] Configuration saved: {selected_printer} on port {port}")
    
    # Start server
    print(f"Starting server with printer: {config['printer_name']}")
    print(f"Port: {config['port']}")
    print("Press Ctrl+C to stop")
    print()
    
    try:
        server.start_server()
    except KeyboardInterrupt:
        print("\n[STOP] Server stopped by user")
    except Exception as e:
        print(f"[!] Server error: {e}")

def run_gui_mode():
    """Run in GUI mode"""
    global AUTO_START_MODE
    
    # Setup GUI startup logging
    gui_logger = None
    try:
        gui_logger = logging.getLogger('gui_startup')
        if not gui_logger.handlers:  # Avoid duplicate handlers
            gui_logger.setLevel(logging.INFO)
            
            # Create logs directory if it doesn't exist - use temp fallback if permission denied
            logs_dir = "logs"
            try:
                if not os.path.exists(logs_dir):
                    os.makedirs(logs_dir)
            except PermissionError:
                # Fallback to user temp directory if permission denied
                import tempfile
                logs_dir = os.path.join(tempfile.gettempdir(), "PrinterOne", "logs")
                if not os.path.exists(logs_dir):
                    os.makedirs(logs_dir, exist_ok=True)
            
            # Generate GUI startup log filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            gui_startup_log_filename = f"gui_startup_{timestamp}.log"
            gui_startup_log_path = os.path.join(logs_dir, gui_startup_log_filename)
            
            # Create file handler for GUI startup log
            file_handler = logging.FileHandler(gui_startup_log_path, encoding='utf-8')
            file_handler.setLevel(logging.INFO)
            
            # Create formatter
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            
            # Add handler to logger
            gui_logger.addHandler(file_handler)
    except Exception as e:
        print(f"Error setting up GUI startup logging: {e}")
        gui_logger = None
    
    try:
        if gui_logger:
            gui_logger.info("=== GUI Mode Starting ===")
            gui_logger.info(f"Process ID: {os.getpid()}")
            gui_logger.info(f"Arguments: {sys.argv}")
        
        # Check if this is an auto-start instance (works for both script and exe)
        AUTO_START_MODE = 'auto_start' in sys.argv
        
        if gui_logger:
            gui_logger.info(f"Auto-start mode: {AUTO_START_MODE}")
        
        # Kill existing GUI instances
        killed_count = 0
        try:
            if gui_logger:
                gui_logger.info("Checking for existing GUI instances...")
            
            current_pid = os.getpid()
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['pid'] == current_pid:
                        continue
                    
                    process_name = proc.info['name'].lower()
                    cmdline = ' '.join(proc.info['cmdline']) if proc.info['cmdline'] else ''
                    
                    # Kill other GUI instances
                    if ((process_name == 'python.exe' and 'server.py' in cmdline and 'gui' in cmdline) or
                        process_name == 'printerone.exe'):
                        if gui_logger:
                            gui_logger.info(f"Killing existing instance: {process_name} (PID: {proc.info['pid']})")
                        proc.terminate()
                        try:
                            proc.wait(timeout=3)
                        except psutil.TimeoutExpired:
                            proc.kill()
                        killed_count += 1
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except Exception as e:
            error_msg = f"Error killing existing instances: {e}"
            print(error_msg)
            if gui_logger:
                gui_logger.error(error_msg)
        
        if killed_count > 0:
            info_msg = f"Killed {killed_count} existing GUI instance(s)"
            print(info_msg)
            if gui_logger:
                gui_logger.info(info_msg)
            time.sleep(1)
        elif gui_logger:
            gui_logger.info("No existing GUI instances found")
        
        # Create and run GUI
        if gui_logger:
            gui_logger.info("Creating Tkinter root window...")
        
        root = tk.Tk()
        
        if gui_logger:
            gui_logger.info("Tkinter root window created successfully")
            gui_logger.info("Initializing PrinterOneGUI...")
        
        try:
            app = PrinterOneGUI(root)
            if gui_logger:
                gui_logger.info("PrinterOneGUI initialized successfully")
                gui_logger.info("Starting Tkinter mainloop...")
            
            root.mainloop()
            
            if gui_logger:
                gui_logger.info("Tkinter mainloop completed normally")
                
        except Exception as e:
            error_msg = f"GUI error: {e}"
            print(error_msg)
            if gui_logger:
                gui_logger.critical(error_msg)
                gui_logger.critical(f"Exception type: {type(e).__name__}")
                import traceback
                gui_logger.critical(f"Traceback: {traceback.format_exc()}")
            
            import traceback
            traceback.print_exc()
            
            # Re-raise for proper error handling
            raise
            
    except Exception as e:
        error_msg = f"Critical GUI startup error: {e}"
        print(error_msg)
        if gui_logger:
            gui_logger.critical(error_msg)
            gui_logger.critical(f"Exception type: {type(e).__name__}")
            import traceback
            gui_logger.critical(f"Traceback: {traceback.format_exc()}")
        
        # Re-raise the exception to maintain original behavior
        raise

def run_test_mode():
    """Run test client"""
    print("PrinterOne - Test Client")
    print("=" * 24)
    print()
    
    # Get server details
    host = input("Server host (default: localhost): ").strip() or "localhost"
    port_input = input("Server port (default: 9100): ").strip()
    port = int(port_input) if port_input else 9100
    
    # Test connection
    TestClient.test_connection(host, port)

def show_help():
    """Show help information"""
    print("PrinterOne - Network Print Server")
    print("=" * 35)
    print()
    print("Usage:")
    print("  python server.py              - Run in console mode")
    print("  python server.py gui          - Run with GUI")
    print("  python server.py gui auto_start - Run GUI in auto-start mode")
    print("  python server.py test         - Run test client")
    print("  python server.py --help       - Show this help")
    print()
    print("Features:")
    print("  • TCP print server for raw print data")
    print("  • Automatic PDF conversion for PDF printers")
    print("  • GUI management interface")
    print("  • Built-in test client")
    print("  • Windows auto-start support")
    print("  • System tray integration")
    print()
    print("Copyright (c) 2025 xtieume@gmail.com")
    print("GitHub: https://github.com/xtieume/PrinterOne")

def setup_startup_logging():
    """Setup startup logging to track OS calls and initialization failures"""
    try:
        # Create logs directory if it doesn't exist - use temp fallback if permission denied
        logs_dir = "logs"
        try:
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)
        except PermissionError:
            # Fallback to user temp directory if permission denied
            import tempfile
            logs_dir = os.path.join(tempfile.gettempdir(), "PrinterOne", "logs")
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir, exist_ok=True)
        
        # Generate startup log filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        startup_log_filename = f"startup_{timestamp}.log"
        startup_log_path = os.path.join(logs_dir, startup_log_filename)
        
        # Setup startup logger
        startup_logger = logging.getLogger('startup')
        startup_logger.setLevel(logging.INFO)
        
        # Create file handler for startup log
        file_handler = logging.FileHandler(startup_log_path, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        
        # Add handler to logger
        startup_logger.addHandler(file_handler)
        
        return startup_logger
    except Exception as e:
        print(f"Error setting up startup logging: {e}")
        return None

def main():
    """Main function"""
    # Setup startup logging first
    startup_logger = setup_startup_logging()
    
    try:
        if startup_logger:
            startup_logger.info("=== PrinterOne Application Started ===")
            startup_logger.info(f"Python version: {sys.version}")
            startup_logger.info(f"Command line arguments: {sys.argv}")
            startup_logger.info(f"Working directory: {os.getcwd()}")
            startup_logger.info(f"Executable path: {sys.executable}")
            
            # Log if running from exe
            if hasattr(sys, '_MEIPASS'):
                startup_logger.info(f"Running from PyInstaller exe: {sys.executable}")
                startup_logger.info(f"Bundle dir: {sys._MEIPASS}")
            else:
                startup_logger.info("Running from Python script")
        
        if len(sys.argv) > 1:
            command = sys.argv[1].lower()
            
            if startup_logger:
                startup_logger.info(f"Command mode: {command}")
            
            if command in ['--help', '-h', 'help']:
                if startup_logger:
                    startup_logger.info("Showing help")
                show_help()
            elif command == 'gui':
                if startup_logger:
                    startup_logger.info("Starting GUI mode")
                run_gui_mode()
            elif command == 'test':
                if startup_logger:
                    startup_logger.info("Starting test mode")
                run_test_mode()
            else:
                if startup_logger:
                    startup_logger.error(f"Unknown command: {command}")
                print(f"Unknown command: {command}")
                show_help()
        else:
            if startup_logger:
                startup_logger.info("Default mode - attempting GUI")
            # Default to GUI mode if available, otherwise console
            try:
                run_gui_mode()
            except ImportError as e:
                if startup_logger:
                    startup_logger.error(f"GUI dependencies not available: {e}")
                    startup_logger.info("Falling back to console mode")
                print(f"GUI dependencies not available: {e}")
                print("Running in console mode...")
                run_console_mode()
        
        if startup_logger:
            startup_logger.info("Application completed successfully")
            
    except Exception as e:
        error_msg = f"Critical startup error: {e}"
        print(error_msg)
        if startup_logger:
            startup_logger.critical(error_msg)
            startup_logger.critical(f"Exception type: {type(e).__name__}")
            import traceback
            startup_logger.critical(f"Traceback: {traceback.format_exc()}")
        
        # Re-raise the exception to maintain original behavior
        raise

if __name__ == "__main__":
    try:
        if startup_logger:
            startup_logger.info("=== Main execution started ===")
            startup_logger.info(f"OS called application: {sys.executable}")
            startup_logger.info(f"Arguments passed: {sys.argv}")
            startup_logger.info(f"Environment: {dict(os.environ)}")
        
        main()
        
        if startup_logger:
            startup_logger.info("=== Main execution completed successfully ===")
            
    except Exception as e:
        error_msg = f"CRITICAL APPLICATION FAILURE: {e}"
        print(error_msg)
        
        if startup_logger:
            startup_logger.critical("=== CRITICAL APPLICATION FAILURE ===")
            startup_logger.critical(error_msg)
            startup_logger.critical(f"Exception type: {type(e).__name__}")
            startup_logger.critical(f"Full traceback: {traceback.format_exc()}")
            startup_logger.critical("=== END OF CRITICAL FAILURE LOG ===")
        else:
            print(f"Exception type: {type(e).__name__}")
            print(f"Full traceback: {traceback.format_exc()}")
        
        # Keep console open for debugging
        try:
            input("Press Enter to exit...")
        except:
            pass
            
        sys.exit(1)