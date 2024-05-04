# File Monitor

File Monitor is a Python script that automates the monitoring of file modification dates at user-defined intervals. If any file is detected as modified, the script automatically opens it using Microsoft PowerPoint. This tool is ideal for tracking updates in presentation files during collaborative work or frequent revisions. This project is licensed under the GNU General Public License v3 (GPLv3).

## Features

- **Automatic File Opening**: Opens modified files with Microsoft PowerPoint as soon as changes are detected.
- **Customizable Monitoring**: Allows users to set their preferred intervals for file modification checks.
- **Ease of Use**: Simple setup process for specifying files to monitor and the PowerPoint application path.

## Requirements

- Microsoft PowerPoint (POWERPNT.EXE)
- Python 3.11.7 (recommended to be packaged by Anaconda, Inc.)

## Configuration
Before running the script, you need to configure the scirpt by specifying:

- File Paths: Paths and names of files to be monitored.
- Monitoring Interval: Time interval for monitoring (hours, minutes, seconds).
- PowerPoint Path: Path to your PowerPoint executable (POWERPNT.EXE).

## Usage
Simply run the script. It will open a window that displays the status of the monitored files.

## Contribution and Maintenance
For issues and support, please contact the maintainers via email at 69543538@qq.com or open an issue in this repository.

Contributions to improve the script are welcome. Please fork the repository, make your changes, and submit a pull request.
