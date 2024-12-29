import platform
import subprocess
import os
import json
import csv
import sys
from datetime import datetime
from pathlib import Path

class SystemInfoCollector:
    def __init__(self):
        self.system = platform.system()
        self.specs = {}

    def get_command_output(self, command):
        try:
            output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
            return output.strip()
        except:
            return "Not available"

    def get_linux_info(self):
        # CPU
        self.specs['CPU'] = self.get_command_output("lscpu | grep 'Model name' | cut -f 2 -d ':'")
        # RAM
        mem_kb = self.get_command_output("grep MemTotal /proc/meminfo | awk '{print $2}'")
        try:
            self.specs['RAM_GB'] = round(int(mem_kb) / 1024 / 1024, 2)
        except:
            self.specs['RAM_GB'] = "Unknown"
        # Storage
        self.specs['Storage'] = self.get_command_output("df -h / | awk 'NR==2 {print $2}'")
        # Serial Number
        self.specs['Serial_Number'] = self.get_command_output("sudo dmidecode -s system-serial-number")
        # Model
        self.specs['Model'] = self.get_command_output("sudo dmidecode -s system-product-name")
        # Manufacturer
        self.specs['Manufacturer'] = self.get_command_output("sudo dmidecode -s system-manufacturer")

    def get_mac_info(self):
        # System Profile
        system_profiler = "system_profiler SPHardwareDataType SPStorageDataType"
        self.specs['Serial_Number'] = self.get_command_output("system_profiler SPHardwareDataType | grep 'Serial Number' | awk '{print $4}'")
        self.specs['Model'] = self.get_command_output("system_profiler SPHardwareDataType | grep 'Model Name' | cut -f 2 -d ':'")
        self.specs['CPU'] = self.get_command_output("sysctl -n machdep.cpu.brand_string")
        # RAM
        ram_bytes = self.get_command_output("sysctl hw.memsize | awk '{print $2}'")
        try:
            self.specs['RAM_GB'] = round(int(ram_bytes) / 1024 / 1024 / 1024, 2)
        except:
            self.specs['RAM_GB'] = "Unknown"
        # Storage
        self.specs['Storage'] = self.get_command_output("diskutil info disk0 | grep 'Disk Size' | awk '{print $3,$4}'")

    def get_windows_info(self):
        try:
            import wmi
            c = wmi.WMI()
            
            system = c.Win32_ComputerSystem()[0]
            cpu = c.Win32_Processor()[0]
            bios = c.Win32_BIOS()[0]
            os_info = c.Win32_OperatingSystem()[0]
            
            self.specs['Serial_Number'] = bios.SerialNumber
            self.specs['Manufacturer'] = system.Manufacturer
            self.specs['Model'] = system.Model
            self.specs['CPU'] = cpu.Name
            self.specs['RAM_GB'] = round(float(system.TotalPhysicalMemory) / 1024**3, 2)
            
            # Storage
            disks = []
            for disk in c.Win32_DiskDrive():
                size_gb = round(float(disk.Size) / 1024**3, 2)
                disks.append(f"{disk.Model} ({size_gb} GB)")
            self.specs['Storage'] = ', '.join(disks)
            
        except ImportError:
            # Fallback to PowerShell commands if WMI is not available
            self.specs['Serial_Number'] = self.get_command_output("wmic bios get serialnumber")
            self.specs['Model'] = self.get_command_output("wmic computersystem get model")
            self.specs['CPU'] = self.get_command_output("wmic cpu get name")
            self.specs['RAM_GB'] = self.get_command_output("wmic computersystem get totalphysicalmemory")
            self.specs['Storage'] = self.get_command_output("wmic diskdrive get size,model")

    def get_gpu_info(self):
        if self.system == 'Windows':
            try:
                import wmi
                c = wmi.WMI()
                gpus = []
                for gpu in c.Win32_VideoController():
                    gpus.append({
                        'Name': gpu.Name,
                        'Memory': f"{round(float(gpu.AdapterRAM or 0) / 1024**3, 2)} GB" if gpu.AdapterRAM else "Unknown",
                        'Driver_Version': gpu.DriverVersion
                    })
                return gpus
            except:
                # Fallback for Windows without WMI
                gpu_info = self.get_command_output("wmic path win32_VideoController get name,adapterram,driverversion")
                return [{'Name': gpu_info}]

        elif self.system == 'Linux':
            # Try lspci first
            gpu_info = self.get_command_output("lspci | grep -E 'vga|3d|2d'")
            if not gpu_info:
                # Try nvidia-smi for NVIDIA cards
                gpu_info = self.get_command_output("nvidia-smi --query-gpu=gpu_name,memory.total,driver_version --format=csv,noheader")
                if not gpu_info:
                    # Try AMD specific tools
                    gpu_info = self.get_command_output("glxinfo | grep 'OpenGL renderer'")
            
            return [{'Name': gpu_info}]

        elif self.system == 'Darwin':  # macOS
            gpu_info = self.get_command_output("system_profiler SPDisplaysDataType")
            return [{'Name': gpu_info}]

    def collect_info(self):
        # Basic info for all platforms
        self.specs['OS'] = platform.platform()
        self.specs['Hostname'] = platform.node()
        self.specs['OS_Version'] = platform.version()
        self.specs['Architecture'] = platform.machine()
        self.specs['Scan_Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Platform specific info
        if self.system == 'Linux':
            self.get_linux_info()
        elif self.system == 'Darwin':  # macOS
            self.get_mac_info()
        elif self.system == 'Windows':
            self.get_windows_info()

        gpu_info = self.get_gpu_info()
        for i, gpu in enumerate(gpu_info):
            if isinstance(gpu, dict):
                self.specs[f'GPU_{i+1}_Name'] = gpu.get('Name', 'Unknown')
                self.specs[f'GPU_{i+1}_Memory'] = gpu.get('Memory', 'Unknown')
                self.specs[f'GPU_{i+1}_Driver'] = gpu.get('Driver_Version', 'Unknown')
            else:
                self.specs[f'GPU_{i+1}'] = str(gpu)

        return self.specs



    def save_results(self):
        # Get the directory where the executable is located 
        if getattr(sys, 'frozen', False):
            # running in a PyInstaller bundle
            script_dir = Path(sys._MEIPASS).parent
        else:
            # running in a normal Python environment
            script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # get filename from user
        filename_base = input("Enter a name for the report (without extension): ").strip()
        if not filename_base:
            filename_base = "system_info"  # default name if empty
        
        # timestamp
        filename_base = f"{filename_base}_{timestamp}"
        
        # create a 'reports' directory if it doesn't exist
        reports_dir = script_dir / 'reports'
        reports_dir.mkdir(exist_ok=True)
        
        # Save as CSV
        csv_file = reports_dir / f"{filename_base}.csv"
        with open(csv_file, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Specification', 'Value'])
            for key, value in self.specs.items():
                writer.writerow([key, str(value)])
        
        # Save as JSON
        json_file = reports_dir / f"{filename_base}.json"
        with open(json_file, 'w') as f:
            json.dump(self.specs, f, indent=4)
        
        return csv_file, json_file

if __name__ == "__main__":
    try:
        collector = SystemInfoCollector()
        collector.collect_info()
        csv_file, json_file = collector.save_results()
        print(f"\nSystem information collected successfully!")
        print(f"CSV saved to: {csv_file}")
        print(f"JSON saved to: {json_file}")
        print("\nCollected Information:")
        for key, value in collector.specs.items():
            print(f"{key}: {value}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        input("\nPress Enter to exit...")