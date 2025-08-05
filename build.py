#!/usr/bin/env python3
"""
Unified Build Script for PrinterOne
Builds both main server and GUI executables
"""

import os
import sys
import subprocess
import shutil
import psutil
import time

def install_requirements():
    """Install required packages"""
    print("Installing requirements...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
        print("[OK] Requirements installed successfully")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Error installing requirements: {e}")
        return False
    return True

def clean_build():
    """Clean previous build files"""
    print("Cleaning previous build...")
    try:
        # Kill any running processes first
        kill_running_processes()
        
        # Clean build folders
        if os.path.exists("build"):
            shutil.rmtree("build")
        if os.path.exists("dist"):
            # Try to remove dist folder, handle permission errors
            try:
                shutil.rmtree("dist")
            except PermissionError:
                print("[WARNING]  Permission denied removing dist folder, trying to force...")
                # Try to remove individual files first
                for root, dirs, files in os.walk("dist", topdown=False):
                    for file in files:
                        filepath = os.path.join(root, file)
                        force_remove_file(filepath)
                    for dir in dirs:
                        try:
                            os.rmdir(os.path.join(root, dir))
                        except:
                            pass
                # Try to remove the folder again
                try:
                    os.rmdir("dist")
                except:
                    print("[WARNING]  Could not fully clean dist folder, continuing...")
        
        # Clean spec files
        spec_files = ["PrinterOne.spec", "PrinterOneManager.spec"]
        for spec_file in spec_files:
            if os.path.exists(spec_file):
                force_remove_file(spec_file)
        
        print("[OK] Build cleaned successfully")
        return True
    except Exception as e:
        print(f"[ERROR] Error cleaning build: {e}")
        return False


def build_gui_exe():
    """Build the PrinterOne GUI executable (includes integrated server)"""
    print("Building PrinterOne GUI executable...")
    
    # Kill processes and clean up before building
    kill_running_processes()
    
    # Force remove existing executable if it exists
    gui_exe_path = os.path.join("dist", "PrinterOne.exe")
    if os.path.exists(gui_exe_path):
        print(f"Removing existing {gui_exe_path}...")
        force_remove_file(gui_exe_path)
    
    try:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--noconsole",
            "--name=PrinterOne",
            "--icon=printer.ico",  # Sử dụng file .ico đúng chuẩn Windows
            "--add-data=config.json;.",
            "--add-data=printer.png;.",  # Đóng gói luôn file PNG vào exe
            "--hidden-import=pystray",
            "--hidden-import=PIL",
            "--hidden-import=PIL.Image",
            "--hidden-import=psutil",
            "--hidden-import=win32print",
            "--hidden-import=win32api", 
            "--hidden-import=win32con",
            "--hidden-import=winreg",
            "server.py"
        ]
        
        subprocess.run(cmd, check=True)
        print("[OK] PrinterOne GUI executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Error building PrinterOne executable: {e}")
        return False

def check_gui_executable():
    """Check if GUI executable was built successfully"""
    print("Checking PrinterOne executable...")
    try:
        exe_path = os.path.join("dist", "PrinterOne.exe")
        if os.path.exists(exe_path):
            file_size = os.path.getsize(exe_path)
            print(f"[OK] PrinterOne.exe built successfully ({file_size:,} bytes)")
            return True
        else:
            print("[ERROR] PrinterOne.exe not found in dist directory")
            return False
    except Exception as e:
        print(f"[ERROR] Error checking executable: {e}")
        return False

def kill_running_processes():
    """Kill any running PrinterOne processes to free up files"""
    print("Checking for running PrinterOne processes...")
    killed_count = 0
    
    try:
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                process_name = proc.info['name'].lower()
                
                # Check for PrinterOne executable
                if process_name == 'printerone.exe':
                    print(f"Killing process: {proc.info['name']} (PID: {proc.info['pid']})")
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except psutil.TimeoutExpired:
                        proc.kill()
                        proc.wait(timeout=2)
                    killed_count += 1
                    
                # Also check for Python processes running server.py
                elif process_name == 'python.exe' and proc.info['exe']:
                    try:
                        cmdline = ' '.join(proc.cmdline())
                        if 'server.py' in cmdline:
                            print(f"Killing Python process running server.py: PID {proc.info['pid']}")
                            proc.terminate()
                            try:
                                proc.wait(timeout=5)
                            except psutil.TimeoutExpired:
                                proc.kill()
                                proc.wait(timeout=2)
                            killed_count += 1
                    except (psutil.AccessDenied, psutil.NoSuchProcess):
                        pass
                        
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
                
    except Exception as e:
        print(f"Error killing processes: {e}")
    
    if killed_count > 0:
        print(f"[OK] Killed {killed_count} process(es)")
        # Wait a moment for processes to fully terminate and files to be released
        time.sleep(2)
    else:
        print("[OK] No running processes found")
    
    return killed_count > 0


def force_remove_file(filepath):
    """Force remove a file, trying multiple methods"""
    if not os.path.exists(filepath):
        return True
        
    try:
        # First try normal removal
        os.remove(filepath)
        return True
    except PermissionError:
        print(f"[WARNING]  Permission denied for {filepath} - file may be in use")
        return False
    except Exception as e:
        print(f"[ERROR] Error removing {filepath}: {e}")
        return False

def main():
    """Main build function"""
    print("PrinterOne - Unified Build Script")
    print("=================================")
    print()
    print("Copyright (c) 2025 xtieume@gmail.com")
    print("GitHub: https://github.com/xtieume/PrinterOne")
    print()
    
    # Install requirements
    if not install_requirements():
        return
    
    # Clean previous build (skip if files are in use)
    try:
        clean_build()
    except Exception as e:
        print(f"[WARNING]  Some files could not be cleaned: {e}")
        print("Trying to kill processes and continue...")
        kill_running_processes()
        time.sleep(1)
    
    # Build GUI executable (includes integrated server)
    print("\n[BUILD] Building PrinterOne GUI executable...")
    if not build_gui_exe():
        print("[ERROR] Failed to build GUI executable. Stopping build process.")
        return 1
    
    # Check executable
    print("\n[VERIFY] Verifying build results...")
    if not check_gui_executable():
        print("[ERROR] Build verification failed. Executable is missing.")
        return 1
    
    print()
    print("[SUCCESS] Build completed successfully!")
    print("[FOLDER] Generated file in dist/ folder:")
    print("  • dist/PrinterOne.exe (GUI application with integrated server)")
    print()
    print("Usage:")
    print("  • Double-click PrinterOne.exe to launch the GUI")
    print("  • The GUI includes server management, test client, and settings")
    print("  • Server will auto-start when printer is configured")
    print()
    return 0

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code if exit_code is not None else 0)