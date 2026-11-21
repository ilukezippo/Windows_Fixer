# Changelog

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
