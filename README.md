# Add-ClinicPrinter

This repository contains a script pair to add a TCP/IP printer by IP address.

This is a de-identified version of a script used and validated in production
support workflows.

Files
- Add-ClinicPrinter.cmd: Wrapper that launches the PowerShell script.
- Add-ClinicPrinter.ps1: Creates the port, adds the printer, and can print a test page.
- Add-ClinicPrinter.log: Run log output stored on the configured network share.

Paths
- Local repo folder: this directory
- Example remote share: \\your-fileserver\SharedTools\Add-ClinicPrinter

Sanitization notes
- Internal server/share names were replaced with placeholders.
- Internal email addresses/domains were replaced with placeholders.
- Organization-specific user path references were removed.
- Core script behavior and control flow were intentionally kept close to the production-used version.

Usage
- Run: Add-ClinicPrinter.cmd and enter the printer IP address (optionally enter a printer name).
- Optional: run Add-ClinicPrinter.ps1 with -PrinterIp, -PrinterName, -PrintTestPage, or -DriverName

End-user run instructions (Outlook-safe)
- Copy and paste the path into the Windows search bar, then press Enter.
- Double-click Add-ClinicPrinter.cmd.
- When prompted, click Run.
- Path (link to the directory, not the .cmd file itself): \\your-fileserver\SharedTools\Add-ClinicPrinter\
