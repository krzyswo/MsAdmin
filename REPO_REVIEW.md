# Repository standardization notes

## Repeated patterns worth centralizing
- **Graph/Exchange/SharePoint module installation and connection logic** – scripts repeatedly check for required modules, prompt to install, and then connect using either credentials or certificate-based authentication. Examples include MS Graph setup in `export-report-365-guest-user-report.ps1` and PnP PowerShell setup in `export-report-sharepoint-links.ps1`. Extracting these checks into shared helper functions (for example, a dot-sourced `Common/Connect-MgGraph.ps1` and `Common/Connect-PnP.ps1`) would reduce duplication and keep authentication behavior consistent.
- **Large banner-style headers** – most scripts begin with the same block documenting name, version, website, and “Script Highlights.” Creating a template (or using a shared comment-based help block) would keep the metadata aligned across scripts and make it easier to update versions in one place.

## Organization opportunities
- **Directory naming** – adopt lowercase kebab-case for folders and primary script names so automation can find README/PS1 pairs predictably. Recommended pattern: `<action>-<workload>-<target>` (for example, `export-exo-mailbox-reports`, `audit-teams-membership-changes`, `manage-spo-site-permissions`). Avoid trailing spaces, underscores, and mixed casing when renaming existing folders.
- **Directory rename workflow** – rename folders with `Rename-Item`/`mv` in a dedicated branch, update README links, and keep a temporary alias map until downstream automation is updated. Sorting the root by prefix (audit/export/manage/setup) makes the catalog easier to scan.
- **Script scaffolding** – consider a minimal starter script (parameters, module checks, logging conventions, and CSV export pattern) that new utilities can copy. This would align parameter naming (`ClientId` vs `ClientID`), version fields, and output paths.
- **Documentation index** – use the root `README.md` to surface every script. Once scripts share a metadata block, regenerate the README table to include workload, purpose, required modules, authentication modes, and key parameters automatically.

## Consistency checkpoints for future work
- Standardize how output files are named (currently each script embeds timestamps differently) and ensure `Export-Csv` calls use the same options (e.g., `-NoTypeInformation`, UTF-8 encoding).
- Align progress and logging messages so automation users can parse them uniformly (e.g., `Write-Progress` activity text and success/exit messages).
- Keep parameter validation consistent by using `[ValidateSet()]`/`[ValidateNotNullOrEmpty()]` where applicable instead of relying on empty-string checks.
