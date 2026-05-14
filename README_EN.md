# ChoiceDownloader User Guide

[中文说明](README.md)

ChoiceDownloader is a Windows desktop automation tool for batch-downloading company announcements and annual report files from the Choice financial terminal. It reads company names from Excel, connects to or launches the local Choice client, searches each company, opens the company announcement page, applies the “All Announcements / Financial Reports” filters, runs batch download, and filters downloaded files by filename keywords.

The tool is intended for users who already have a valid Choice account, a local Choice installation, and the required data permissions. It does not bypass login, authorization, or platform restrictions, and it does not include the Choice client itself. All files are downloaded through the user’s own locally logged-in Choice terminal.

## Key Features

- Windows GUI for daily use.
- Batch company import from Excel through the `公司名称` column.
- Attach to an existing Choice window, or launch Choice from a configured `.exe` or `.lnk` shortcut.
- Automate company search, page navigation, F9 switching, announcement-page location, announcement filtering, and batch download.
- Support coordinate calibration: test key UI targets, capture the current mouse position, and override default coordinates for Hong Kong stocks or different screen layouts.
- Process a selected Excel row range, such as companies 3 to 20.
- Configure the number of announcements to download per company and the maximum safety limit for the current run.
- Configure startup wait, Enter wait, and F9 wait times.
- Filter downloaded files by filename keywords such as `10-K,20-F,40-F`.
- Choose between “match any keyword” and “match all keywords”.
- Copy only the newest dated matching file into the `最新` folder.
- Support single-company debugging modes, such as navigating to the announcement page without downloading.
- Continue after single-company failures and write a batch summary workbook.
- The release build does not prefill user paths, company lists, download directories, log directories, or filename keywords. Users must provide all run-specific information.
- Avoid persistent GUI settings JSON by default, reducing the risk of leaking local paths when sharing the project.

## Requirements

| Item | Requirement |
| --- | --- |
| Operating system | Windows 10 / Windows 11 desktop environment |
| Data terminal | Choice financial terminal installed and logged in |
| Input file | `.xlsx` or `.xlsm` |
| Recommended display setup | Stable resolution, stable scaling, maximized Choice window |
| Permission | The user must have the required announcement access and download permissions |

This tool uses desktop automation, so it is sensitive to UI layout. During a batch run, avoid clicking inside Choice manually, changing the resolution, changing display scaling, or moving the Choice window.

## Quick Start with exe

Download this file from GitHub Releases:

```text
ChoiceDownloader-v0.3.0-windows-x64.exe
```

Steps:

1. Double-click the executable.
2. Select the Excel company list.
3. Select the local Choice executable or shortcut.
4. Select the download directory and log directory.
5. Manually fill in download count, safety limit, row range, and wait times.
6. If you need annual-report filtering, fill in filename keywords such as `10-K,20-F,40-F`; leave it empty to keep all downloaded files.
7. For the first run, process only one company.
8. Make sure Choice is logged in and ready for interaction.
9. Click “开始批量运行”.
10. After completion, check the download directory, the `最新` folder, and `choice_batch_summary.xlsx`.

## Excel Company List Template

The project provides:

```text
templates\company_list_template.xlsx
```

The template contains two worksheets:

| Worksheet | Description |
| --- | --- |
| `Companies` | Company list. The program reads the `公司名称` column from this sheet. |
| `说明` | Template instructions and notes. |

The `Companies` sheet must keep at least this column:

| 公司名称 |
| --- |
| Microsoft Corp |
| Apple Inc |

Notes:

- The first row must contain the exact column name `公司名称`.
- Company names should start from the second row.
- Sample rows can be deleted.
- `代码或简称（可选）` and `备注（可选）` are for manual reference only and are ignored by the program.
- The program reads the `公司名称` column from the active worksheet. Keeping the default template structure is recommended.

## GUI Fields

| Field | Description |
| --- | --- |
| `Excel 文件` | Excel file containing the `公司名称` column. |
| `Choice 程序/快捷方式` | Local Choice entry point. You can select either `.exe` or `.lnk`. |
| `下载目录` | Directory where Choice download results are saved. |
| `日志目录` | Directory where runtime logs are saved. |
| `每家公司下载篇数` | Number of recent announcements to download per company. |
| `本次上限` | Maximum allowed download count for the current run. Values above this limit are blocked before execution. |
| `提取范围` | Process only the selected row range from the Excel list. After selecting Excel, the program fills the full range automatically and the user can edit it. |
| `启动等待` | Seconds to wait after clicking start, useful for letting the terminal window settle. |
| `Enter 等待` | Seconds to wait after submitting the company search. |
| `F9 等待` | Seconds to wait after sending F9. |
| `文件名筛选` | Optional comma-separated filename keywords. Leave empty to disable filtering. |
| `匹配方式` | Match any keyword, or require all keywords. |
| `只保留最新日期` | Copy only the newest dated matching file into the `最新` folder. |
| `只导航到公告页，不执行下载` | Debug mode: navigate only, do not download. |
| `跳过导航，假设当前已经在公告页` | Useful for single-company debugging only; not recommended for batch runs. |
| `跳过下载弹窗里的目录选择步骤` | Use only when the Choice download dialog already keeps the correct target directory. |
| `最后一家公司完成后返回首页` | Return to the Choice home page after the last company is processed. |

## Coordinate Calibration

Choice UI positions may vary by market, terminal version, resolution, scaling, or window layout. The GUI now includes a coordinate calibration panel. By default, the tool still uses the original coordinates from the U.S. stock workflow. For Hong Kong stocks, you can adjust only the targets that differ.

Available calibration targets:

| Target | Purpose |
| --- | --- |
| `左侧导航滚动点` | Scroll point in the left navigation area after F9. |
| `公司公告入口` | Company announcement entry in the left navigation area. For Hong Kong stocks, adjust this first if the entry is in a different position. |
| `全部公告筛选` | The “All Announcements” filter on the announcement page. |
| `财务报告筛选` | The “Financial Reports” filter. |
| `批量下载按钮` | Batch download button on the announcement list page. |
| `弹窗浏览按钮` | Browse button in the batch download dialog. |
| `弹窗范围勾选` | Range checkbox in the batch download dialog. |
| `弹窗篇数输入` | Download-count input in the batch download dialog. |
| `弹窗下载按钮` | Confirm download button in the batch download dialog. |
| `返回首页按钮` | Return-home entry in Choice. |

Recommended calibration workflow:

1. Fill in the Excel file, Choice entry path, log directory, and wait times.
2. Select a target under “测试定位点”, such as `公司公告入口`.
3. Click “执行到此处并移动鼠标”. The tool uses the first company in the selected range, runs to that step, moves the mouse to the target, and stops before download.
4. If the mouse is not on the correct UI element, edit X/Y manually, or move the mouse to the correct position and click “取鼠标” in that row.
5. Click “确认” in the same row. The current run will use the coordinates shown in the GUI.
6. For Hong Kong stocks, test `公司公告入口` first. If later filters or dialog buttons differ, calibrate those targets one by one.

Coordinates are relative to the top-left corner of the Choice window. For batch runs, keep the Choice window maximized and avoid changing display scaling, resolution, or window position during automation.

## Filename Filtering

Filename filtering is used to keep target filing types. The release build does not prefill keywords. A common U.S. annual report setting is:

```text
10-K,20-F,40-F
```

When the matching mode is “match any keyword”, a file is kept if its filename contains any keyword. The matcher also tolerates common formatting differences:

| Input keyword | Compatible examples |
| --- | --- |
| `10-K` | `10K`, `10 K`, `10－K` |
| `20-F` | `20F`, `20 F`, `20－F` |
| `40-F` | `40F`, `40 F`, `40－F` |

If “latest only” is enabled, the newest dated matching file is copied into the `最新` folder. If no target file is found, the summary workbook records `没筛选成功`.

## Workflow

1. Read company names from Excel.
2. Validate row range, download count, and safety limit.
3. Attach to an existing Choice window; if none is found, launch Choice.
4. Type the company name into the bottom search box.
5. Press Enter to open the company page.
6. Send F9 and wait for the page to respond.
7. Scroll the left navigation area and open the company announcement page.
8. Click “All Announcements”.
9. Click “Financial Reports”.
10. Open the batch download dialog.
11. Set the download directory and download count.
12. Wait until downloaded files appear and become stable.
13. Filter downloaded files by filename keywords.
14. Copy matching files into the `最新` folder.
15. Return to the home page and continue with the next company.
16. Write the batch summary workbook and runtime logs.

## Output Files

Downloaded files are saved to the directory selected in the GUI. Choice may create files or subfolders according to its own download behavior.

Common outputs:

| Output | Description |
| --- | --- |
| Original announcement files | Web or document files downloaded by Choice. |
| `最新\` | Folder created by the tool to store the newest matched file. |
| `choice_batch_summary.xlsx` | Batch summary workbook for success and failure status. |
| Runtime logs | Logs saved to the selected log directory for troubleshooting. |

## Failure Status

Batch runs continue after a single-company failure. The tool tries to return to the home page and then proceeds to the next company.

| Message | Meaning | Suggested action |
| --- | --- | --- |
| `没有找到公司` | The folder browser dialog did not appear after clicking batch download. The expected company announcement page was probably not reached. | Check whether the company can be searched in Choice; if needed, use a more accurate name or CUSIP. |
| `没筛选成功` | Files were downloaded, but no target annual report matched the filename keywords. | Check the company download folder and adjust keywords if needed. |
| `Run stopped by user` | The user clicked stop. | This is a normal interruption; review completed companies. |
| `Choice executable not found` | The configured Choice executable or shortcut does not exist. | Select the correct local Choice entry point. |

## Run from Source

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Launch the GUI:

```powershell
python scripts\choice_automation_gui.py
```

The command-line automation core is located at:

```text
scripts\choice_automation.py
```

Most users should use the GUI.

## Build the Executable

The project includes a PyInstaller spec file:

```text
choice_announcement_workbench.spec
```

Run this command from the project root:

```powershell
pyinstaller choice_announcement_workbench.spec
```

The executable will be generated at:

```text
dist\ChoiceDownloader.exe
```

Recommended GitHub Releases asset name:

```text
ChoiceDownloader-v0.3.0-windows-x64.exe
```

## Project Structure

```text
ChoiceDownloader\
  scripts\
    choice_automation.py          # Core Choice automation workflow
    choice_automation_gui.py      # Tkinter GUI
  templates\
    company_list_template.xlsx    # Company list import template
    README.md                     # Template notes
  release\
    ChoiceDownloader-v0.3.0-windows-x64.exe
  choice_announcement_workbench.spec
  requirements.txt
  README.md
  README_EN.md
  .gitignore
```

## Pre-Run Checklist

Before a real batch task, check that:

- Choice is logged in.
- The first row of the Excel file contains the exact `公司名称` column.
- The download and log directories are writable.
- All path and runtime parameter fields in the GUI have been filled in by the user.
- A one-company run succeeds before a large batch.
- Resolution, display scaling, and Choice layout remain stable.
- You do not manually click the Choice page or move the mouse during automation.
- “Skip navigation” is not enabled for batch runs.
- Filename keywords match the target filing type, such as `10-K,20-F,40-F`.

## Troubleshooting

| Issue | Suggested action |
| --- | --- |
| Choice entry point not found | Select the correct local Choice `.exe` or shortcut in the GUI. |
| `公司名称` column not found | Check that the first row contains exactly `公司名称`. |
| Click positions are inaccurate | Check display scaling, resolution, and Choice layout; test with one company first. |
| No matched files after download | Check filename keywords and matching mode, or clear keywords temporarily to confirm download success. |
| Some companies fail while others work | Review `choice_batch_summary.xlsx` and logs, then rerun failed companies separately. |
| No JSON settings file is generated | This is expected in the current version to avoid local path residue. |
| Folder dialog cannot open | Type the directory path manually, or check the GUI error log in the system temp directory. |

## Privacy and Release Cleanup

Do not commit or share the following files:

- Downloaded results: `downloads/`
- Runtime logs: `logs/`
- PyInstaller intermediate folders: `build/`, `dist/`
- Python cache: `__pycache__/`, `*.pyc`
- Local GUI settings cache: `choice_automation_gui_settings.json`
- Real company lists, business Excel files, or downloaded Choice files

The current version does not write a local settings JSON by default. Logs and downloaded results are written only to the directories selected in the GUI.

## Acknowledgements

Thanks to everyone who provided feedback during Choice announcement workflow testing, coordinate tuning, batch checks, failed-company review, and template cleanup. Thanks also to `pywinauto`, `openpyxl`, and `PyInstaller` for providing the desktop automation, Excel processing, and Windows packaging capabilities used by this tool.

## Copyright

Copyright (c) 2026 Liu Juncheng. All rights reserved.

This project is intended to assist users in batch-downloading announcement files within their own authorized Choice financial terminal environment. Without the author’s permission, do not use this project’s code, templates, or packaged executables for unauthorized commercial distribution. Data downloaded or organized with this tool must be reviewed by the user, and the tool output does not constitute investment, financial, legal, audit, or data compliance advice.
