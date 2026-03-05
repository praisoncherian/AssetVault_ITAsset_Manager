# Usage Guide

Using AssetVault to capture an audit profile of a machine is simple. 
Our tool is designed to combine manual user-inputted context (purchase dates, user assignments) with automatic hardware data scraping.

## 1. Enter Asset Lifecycle Information
Open the application and fill in the required fields under the `Asset Lifecycle Data` tab:
* **Asset Tag**: The unique physical sticker ID on the machine.
* **Assigned User**: The full name of the employee receiving the equipment.
* **Purchase Date**: The format MUST strictly be `dd/mm/yyyy`.
* **Asset Cost**: The numerical cost.

## 2. Execute WMI Extraction
Click **Generate Report**. Our script will freeze the UI main page and seamlessly transition to a live processing log. Wait ~15 seconds while we scrape:
* **BIOS / CPU Arrays**
* **RAM Dimms**
* **Battery / Powercfg HTML hooks**
* **Disk SMART Drives**

## 3. Review the PDF
The raw dataset will automatically compile into a formatted `.pdf` and route directly to your `Downloads/AssetVault/` folder, titled `[Asset_Tag]_[DateTimeStamp].pdf`. You may now print or archive this report for user signatures.
