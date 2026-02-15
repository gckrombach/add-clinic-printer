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

Notes for agents
- Users must supply the printer IP address; there is no default.
- The script does not request elevation; it logs a warning if not admin.
- For non-admin users, port creation may fail depending on policy.
- Test page is optional and prompted interactively after the email is sent.
- Test page uses the Windows PrintUI entry point.
- A log email is sent every run to the configured support mailbox and CCs the user if their email is found in AD.
- The email includes only the most recent run details (no attachments).
- TODO: handle printer-name collision risk (if a user enters a name that already exists, the script may update the wrong printer).

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
