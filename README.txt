claude-safeguard v0.7.1 — Windows finance pilot

START-HERE.txt is the staff installation and daily-use guide.
SHAREPOINT-SETUP.txt is the Microsoft 365 administrator setup guide.
VALIDATION.txt describes what was tested and what remains unverified.

WHAT THIS PACKAGE DOES
Installs missing Claude Code and Python, configures a financial analyst prompt,
and provides a fixed local spreadsheet tool with prompt/file/tool checks.
Staff use desktop shortcuts and ordinary-language requests. There is no hosted
custom application to deploy. Claude accounts, internet access and permission to
install software are still required. Claude use is subject to your account plan.

- Browse a user-selected local folder and its subfolders; import XLSX copies.
- Browse one configured SharePoint document-library folder and import supported
  files through Microsoft sign-in. SharePoint access is read-only.
- Create/read/edit XLSX workbooks, clean rows, sort/filter, join/append, group
  totals, format cells and create bar/line/pie charts.
- Create budget/actual dashboards with decimal amount totals, variance, grouped
  summaries, formatting and charts. Require one explicit currency; no guessed FX.
- Fill templates with explicit source-to-cell mappings. Missing/ambiguous matches,
  formula targets, locked cells and unapproved populated targets stop the fill.
- Preserve source workbooks. Each successful edit creates a new UUID-named XLSX
  in %LOCALAPPDATA%\claude-safeguard\Workbooks.

AUDIT FIXES IN v0.7.1
A fresh install previously selected MCP 2.x, which cannot import the server's
MCP 1.x FastMCP API. Direct dependencies are now pinned to tested versions.
Both installation and repair install those versions and perform a real local
MCP handshake, tool listing, synthetic dashboard and sensitive-pattern refusal.
The launcher checks the installed dependencies, required settings and Claude
version before opening a session. These checks do not require a Claude API call.

Guard hooks now use an absolute PowerShell executable plus an argument array,
avoiding shell quoting of Windows paths. UserPromptSubmit and PreToolUse use
onFailure=block, requiring Claude Code v2.1.295 or later. Setup tries claude update
for older installations; launch refuses them. SessionStart cannot prevent a
session from starting; the prompt/tool checks are the blocking hooks.
Settings JSON is written as UTF-8 without BOM, including on PowerShell 5.1.
Repair backs up settings and scripts, migrates the package's own hook entries,
retains unrelated settings, and updates the fixed MCP server entry. Your source
folder selection, workbooks and approved text are retained.

Workbook reads are bounded to 25 MiB; Windows reparse paths are refused. Failed
imports are cleaned up. External/network formula patterns are checked on import,
creation and edits. Many-to-many joins are refused; use a validated key or
aggregate first. Empty dashboards and missing categories are refused.
Test-Installation.cmd provides a double-click local check using synthetic data.

DATA FLOW AND PRIVACY BOUNDARY
Local and SharePoint imports are scanned locally before content is returned by
the bundled tool or permitted Read. Claude DOES receive approved filenames,
scanned text/cell values and tool results that are returned to it. Computations
run locally, but this is not a fully local AI model.

Pattern checks detect examples such as email addresses, international phone
patterns, secret assignments and sensitive labels. They miss names, indirect
identifiers, many numeric IDs and other sensitive material. False positives are
also possible, including policy documents labelled confidential. A passing scan
does not anonymise data or authorise disclosure. Use public, synthetic or
explicitly approved de-identified data only. Do not remove labels just to pass.

CLAUDE.md is behavioural guidance. Hooks enforce specific tool checks only when
loaded and enabled. This user-installed package is removable and is not an OS
sandbox, enterprise DLP or a guarantee against all data disclosure. It does not
control other Claude sessions, attachments, global plugins/settings, user edits
to the package, local attackers, hard links or concurrent file replacement.
Do not attach files, use @mentions, resume unreviewed sessions, enable additional
connectors or change files while Claude reads them. Built-in /hooks and /mcp UI
commands are used for setup verification. Other tool access is denied by the guard.

The SharePoint adapter uses local Microsoft OAuth and Windows-encrypted token
caching outside the Claude workspace. IT must provide a single-tenant public
client registration, delegated Sites.Selected, a read grant on the approved
site and the configured drive/folder IDs. User permissions also apply. Folder
scoping is implemented by the local tool; the Microsoft grant applies to the
whole selected site. No client secret or hosted service is needed. Read
SHAREPOINT-SETUP.txt before configuring access. No remote upload/edit/delete.

FORMATS AND EXCEL LIMITS
Editable spreadsheet workflow: standard XLSX, at most 25 MiB compressed,
50 MiB expanded and 250,000 cells. Read responses allow up to 5,000 cells.
Table transformations require literal values and unique headers in row 1.
Formulas can be written, but Excel must recalculate them on opening. Structured
reference/external formula patterns may be refused. Native pivot-table creation,
Power Query, VBA, slicers and exact advanced-object fidelity are not supported.
Macros, external workbook links and embedded objects are refused. Advanced Excel
features may be lost when saving a copy. Always compare results in Excel.

Prepared text: UTF-8 TXT/CSV in installed Approved-Data, at most 1 MiB.
The Prepare file shortcut extracts supported DOCX/PPTX/OpenDocument, text-based
PDFs, XLSX and UTF-8 text formats to text for local review before approval.
SharePoint imports XLSX into the workbook workflow and supported documents into
prepared text. PDF extraction is limited to 200 pages; image content is omitted.
Images, scanned-page OCR, audio/video, encrypted files and legacy Office formats
are unsupported. Extraction is not a complete representation of document layout,
images or all embedded objects. Review completeness against the original.

INSTALL / REPAIR
Extract the whole ZIP. Close Claude. Double-click Install.cmd (new or existing)
or Repair.cmd (existing). Then double-click Test-Installation.cmd. Open the
claude-safeguard desktop shortcut or Start.cmd. Use the synthetic rehearsal in
START-HERE.txt before entering approved work data.
An installation interrupted before all files/configuration exist needs IT help;
setup does not automatically reconstruct an incomplete installation. Repair
restores backed-up files/settings if its update fails, but pip dependency changes
are not rolled back. Resolve the error and rerun repair. Do not delete user data.
Backups are in the installed Launcher-Backups folder. Global Claude settings and
persistent PATH are not changed. Windows installation policies still apply.

VALIDATION STATUS
82 automated checks passed on Linux with Python 3.12 and PowerShell 7.4.6.
Real MCP stdio was tested. Microsoft Graph responses were simulated.
Windows PowerShell 5.1, WinGet, COM shortcuts, live Claude hook execution,
Microsoft sign-in/DPAPI/tenant permissions and Excel rendering remain untested.
This is a pilot awaiting a real Windows/tenant acceptance run, not a certified
production release. See VALIDATION.txt for exact coverage and acceptance steps.

REFERENCE DOCUMENTATION (checked 2026-11-03)
https://code.claude.com/docs/en/hooks
https://code.claude.com/docs/en/setup
https://code.claude.com/docs/en/mcp
https://github.com/modelcontextprotocol/python-sdk
https://learn.microsoft.com/en-us/graph/permissions-selected-overview
https://learn.microsoft.com/en-us/graph/api/driveitem-get-content

REMOVAL
Close Claude. Preserve any required Workbooks and Approved-Data first. Use
Disconnect-SharePoint.cmd to clear the configured account's local cached tokens.
Remove the package's desktop shortcuts and installed folder when no longer needed.
Claude Code, Python, Claude history and global credentials are separate and are
not uninstalled or erased by this package. Temporary files are normally deleted;
normal deletion is not secure erasure.
