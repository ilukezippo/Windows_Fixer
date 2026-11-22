# Changelog

## 1.2.0 — November 2026

### New
- **Tabs** instead of menus: Diagnose, Repair & Cleanup, Log, Settings, About. Progress and Start / Skip / Cancel stay visible on every tab
- **Progress percentage** next to the progress bar and in the window title (so you see it on the taskbar)
- Advice when DISM can't get its repair files (0x800F0915, 0x800F081F and similar), and a warning when Windows is waiting for a restart
- **A restore point before repairs** (Settings; on by default, needs admin)
- A new About page, with a link to more apps at ilukezippo.github.io
- Segoe UI Bold everywhere
- **One EXE file** to download instead of a ZIP
- **Find the problem**: describe what's wrong, in English or Arabic, and get the steps that fix it
- **Check for problems**: a quick read-only check of disk space, temporary files, the Recycle Bin, old update downloads, a pending restart, the internet and DNS, disk health and (as admin) the Windows image. Runs at startup; you can turn that off
- A new look: Windows 11 style, with **light and dark themes** (or follow Windows), dark title bar included
- **Arabic right-to-left layout**, and everything is translated: progress, log messages, dialogs
- A new app icon
- A summary at the end of every run, and a reminder when a restart is needed
- Live progress for DISM, SFC and CHKDSK, shown as part of the progress bar
- Copy Log and Save Log; the log is read-only and colour-coded
- Your ticked steps, CHKDSK choices, theme and window size are remembered

### Fixed
- Each command flashed a black console window
- SFC's output came out garbled (it writes UTF-16), and command output in other languages showed the wrong characters
- DISM and SFC progress filled the log with hundreds of lines
- CHKDSK *Fix errors* could hang forever waiting for a Y/N answer
- Skip and Cancel didn't stop DISM's helper process or CHKDSK
- Cancelling *Fix Windows Update downloads* could leave the update services stopped
- Closing the window during a step left the command running
- Declining the UAC prompt closed the app; with *Always run as admin* on it never opened
- Steps reported success even when the command failed; the exit code is now checked
- Steps that need admin now tell you up front, and offer to restart as admin
- Cleanup could follow a junction or symbolic link out of the Temp folder; links are now removed, never followed. Read-only temp files are removed too, and the space freed is reported
- The Recycle Bin step said it was emptied even when it failed
- The About box showed today's date instead of the build date
- The disk check said “Drive C::”
- The window didn't fit on smaller screens (1366×768); the options now scroll
- Sound-debugging lines in the log removed

## 1.0.0 — February 2026

The first release.

### Repair
- Check Windows image health (DISM ScanHealth) and repair it (DISM RestoreHealth)
- Repair system files (SFC /scannow)
- Check a disk for errors (CHKDSK), scan only or fix, with a drive picker
- Reset the network stack (Winsock + TCP/IP)

### Cleanup
- Clean temporary files and Prefetch files
- Empty the Recycle Bin
- Flush the DNS cache
- Clean the Windows component store (DISM StartComponentCleanup)
- Fix Windows Update downloads (clear the update download cache)

### Everywhere
- Select all, a progress bar with step tracking and a full live log
- Skip the current step or cancel everything
- Admin detection, *Run as Admin* and *Always run as admin*
- English and Arabic
