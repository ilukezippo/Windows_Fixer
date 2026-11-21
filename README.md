<p align="center">
  <img src="assets/windows-fixer.png" width="96" alt="Windows Fixer icon">
</p>

<h1 align="center">Windows Fixer</h1>

<p align="center">
  <b>Repair and clean up Windows, in one window.</b><br>
  DISM, SFC, CHKDSK, network reset and safe cleanup tools, with a progress bar and a live log. In English and Arabic.
</p>

<p align="center">
  <a href="../../releases/latest"><img src="https://img.shields.io/github/downloads/ilukezippo/Windows_Fixer/total?label=downloads&color=0a84c6" alt="Downloads"></a>
  <a href="../../releases/latest"><img src="https://img.shields.io/github/v/release/ilukezippo/Windows_Fixer?color=2b2c30" alt="Latest release"></a>
  <img src="https://img.shields.io/badge/Windows-10%20%7C%2011-2b2c30" alt="Windows 10 and 11">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-2b2c30" alt="MIT License"></a>
  <a href="../../stargazers"><img src="https://img.shields.io/github/stars/ilukezippo/Windows_Fixer?style=flat&color=0a84c6" alt="GitHub stars"></a>
</p>

<p align="center">
  <a href="https://github.com/ilukezippo/Windows_Fixer/releases/download/v1.0.0/WindowsFixer_v1.0.0_Windows_x64.zip"><img src="assets/download-button.svg" width="340" alt="Download Windows Fixer for Windows"></a>
</p>

<p align="center">
  <sub>Downloads the latest version straight away: a portable ZIP, nothing to install.</sub>
</p>

<p align="center">
  <a href="../../releases/latest">All downloads</a> ·
  <a href="#how-to-use">How to use</a> ·
  <a href="CHANGELOG.md">What's new</a> ·
  <a href="#what-it-can-fix">What it can fix</a> ·
  <a href="#build-from-source">Build</a>
</p>

<p align="center">
  <img src="shots/main.png" width="760" alt="Windows Fixer main window">
</p>

<p align="center"><sub>If Windows Fixer saves you time, a ⭐ star on this page helps other people find it.</sub></p>

---

## Why

When Windows starts misbehaving, the usual advice is a list of commands to type into an administrator Command Prompt: `DISM /RestoreHealth`, `sfc /scannow`, `chkdsk`, `netsh winsock reset`… Windows Fixer puts them all in one window. Tick what you want, click **Start**, and it runs each step in the right order, with a progress bar, a live log, and the option to skip a step or cancel.

## How to use

1. Run **Windows Fixer**. Most steps need administrator rights: click **Run as Admin**, or turn on **File → Always run as admin**.
2. Tick the repair and cleanup steps you want (or **Select All**). Each one has a short description under it.
3. Click **Start**. The steps run one after another; the progress bar and the log show what's happening.

| Do this | You get |
|---|---|
| **Skip Step** | Stops the step that's running and moves on to the next one |
| **Cancel** | Stops everything |
| **Clear Log** | An empty log |
| **File → Language** | English or العربية |
| **File → About** | The version and a link to this page |

Windows Fixer also checks GitHub for a new version when it starts.

## What it can fix

**Repair**

| Step | Runs | What it does |
|---|---|---|
| Check Windows Image Health | `DISM /Online /Cleanup-Image /ScanHealth` | Checks the Windows image for corruption. Useful before RestoreHealth |
| Repair Windows Image | `DISM /Online /Cleanup-Image /RestoreHealth` | Repairs the Windows image using Windows Update |
| Repair System Files | `sfc /scannow` | Scans and repairs protected system files. Best after DISM |
| Check Disk for errors | `chkdsk C:` or `chkdsk C: /f` | Scans the drive you pick; *Fix errors* may need a restart |
| Reset Network Stack | `netsh winsock reset`, `netsh int ip reset` | Fixes common network problems. May need a restart |

**Cleanup**

| Step | What it does |
|---|---|
| Clean Temporary Files | Deletes files from your Temp folder and Windows Temp (files in use are skipped) |
| Clean Prefetch Files | Clears the Prefetch cache; Windows rebuilds it |
| Empty Recycle Bin | Frees the space right away |
| Flush DNS Cache | `ipconfig /flushdns`, which can help with some internet problems |
| Clean Windows Component Store | `DISM /StartComponentCleanup`: removes superseded Windows components |
| Fix Windows Update downloads | Stops the update services, clears the downloaded update files and starts the services again |

## Features

- **Repair and cleanup in one place**, run in a sensible order
- **Progress bar** with step tracking and a **full live log** of every command's output
- **Skip** the current step or **cancel** everything at any time
- **Admin detection**, *Run as Admin*, and *Always run as admin*
- **English and Arabic**
- **Updates:** checks this GitHub page for a new version
- A success sound when everything is done

## Install

1. **[Download WindowsFixer_v1.0.0_Windows_x64.zip](https://github.com/ilukezippo/Windows_Fixer/releases/download/v1.0.0/WindowsFixer_v1.0.0_Windows_x64.zip)** (the latest version), or pick a file on the [Releases](../../releases/latest) page.
2. Unzip it anywhere and run **WindowsFixer_v1.0.0.exe**. It's portable: nothing is installed.

> The app is not code-signed yet, so Windows may warn about it:
>
> - **Microsoft Edge** may hold the download ("isn't commonly downloaded"): open the downloads list, point at the file, choose **… → Keep**, then **Show more → Keep anyway**.
> - **SmartScreen** may warn when you run it: choose **More info → Run anyway**.

The ZIP also has `SHA256SUMS.txt` if you want to check the download.

**Requirements:** Windows 10 or 11.

## Build from source

You need Python 3.9 or newer.

```bat
git clone https://github.com/ilukezippo/Windows_Fixer.git
cd Windows_Fixer
run.bat      :: first run creates a .venv, installs dependencies, starts the app
build.bat    :: builds dist\windows_fixer.exe
```

Every push to `main` builds the app with GitHub Actions. To publish a release, set `APP_VERSION` in `windows_fixer.py`, add the version to [CHANGELOG.md](CHANGELOG.md), and push a tag like `v1.0.1`: the workflow builds `WindowsFixer_v1.0.1_Windows_x64.zip`, publishes it with that version's changelog, and points the Download buttons on this page at it.

### Project layout

```
windows_fixer.py   the whole app (Tkinter + Pillow)
icon.ico           app icon
kuwait.png         flag shown in About
Success.wav        sound played when everything finished
assets/            pictures for this page
```

## How it works

- Each ticked step runs a standard Windows command (DISM, SFC, CHKDSK, netsh, ipconfig, net stop/start) or a small built-in cleanup, one after another.
- The output of every command streams into the log as it runs. **Skip** and **Cancel** stop the running command.
- Your language and *Always run as admin* choice are saved in `%APPDATA%\WindowsFixer\settings.json`. See the [privacy policy](PRIVACY.md).

## Known limits

- Most repair steps need administrator rights.
- DISM RestoreHealth needs Windows Update to be reachable.
- CHKDSK's *Fix errors* on the Windows drive, and the network reset, take effect after a restart.

## License

[MIT](LICENSE) © ilukezippo
