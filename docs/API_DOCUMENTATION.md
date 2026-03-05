# API Documentation (CIM / WMI Hooks)

AssetVault leverages native `Get-CimInstance` and `powercfg` integrations. No external software dependencies are needed.

### Primary NameSpaces Queried
* **Win32_OperatingSystem:** Scrapes the BuildNumber, OS architecture, and install date.
* **Win32_ComputerSystem:** Used for Domain joining state, Manufacturer, and Model.
* **Win32_BIOS:** Hooks into the motherboard chipset for UUID and Serial Number.
* **Win32_Processor & Win32_PhysicalMemory:** Extracts Logical Cores, CPU labeling, and calculates true RAM Capacity conversions.
* **Win32_VideoController:** Hooks the Display Adapter VRAM limits.
* **MSFT_PhysicalDisk (Storage namespace):** Targets SMART health status strings across external and internal interfaces.

### Export Handlers
* **powercfg /batteryreport:** Exports native HTML diagnostic tables into the system `$env:TEMP` cache directory.
* **powercfg /energy:** Triggers a fast 5-second energy trace event log.
* **Edge Headless Print:** Uses built-in `msedge.exe` engine with `--headless --print-to-pdf` to convert our `$HtmlBody` concatenated payload directly into an offline PDF.
