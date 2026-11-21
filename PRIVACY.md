# Windows Fixer privacy policy

_Last updated: November 2026 (Windows Fixer 1.0.0)_

Windows Fixer works on your own computer. It collects **no personal data**, no usage statistics and no telemetry, and it has no accounts.

It only goes online for these things:

| What | When | What is contacted |
|---|---|---|
| **Update check** | When it starts | `api.github.com` (this project's releases). A standard HTTPS request, no personal data |
| **Repair Windows Image (DISM RestoreHealth)** | Only when you run that step | Windows Update, which DISM uses as its repair source (a Windows feature, not Windows Fixer's own connection) |
| **Buy me a coffee / GitHub page** | Only when you click them | Your browser opens the link |

The repair and cleanup steps are standard Windows commands (DISM, SFC, CHKDSK, netsh, ipconfig) run on your PC. Their output is shown in the log and isn't saved or sent anywhere.

Your choices (language, *Always run as admin*) are saved in `%APPDATA%\WindowsFixer\settings.json`.

## Contact

Questions: open an issue at https://github.com/ilukezippo/Windows_Fixer/issues
