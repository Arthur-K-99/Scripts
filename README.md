# Automation Scripts

A growing collection of practical automation and administration scripts.

The repository is currently focused on PowerShell and Windows administration because those are the tools most relevant to my day-to-day work. It is intentionally language-agnostic and may also include useful Bash, zsh, Python, or other automation scripts over time.

## Current Areas

| Directory | Purpose |
| --- | --- |
| `Active_Directory/` | Active Directory reporting, auditing, and user or computer administration |
| `Printer_Install_Scripts/` | Printer installation and default-printer configuration |
| `Prune_Old_User_Profiles/` | Removal of stale Windows user profiles |
| `Remove_Aged_Files/` | Retention-based file cleanup with optional ACL repair and ownership takeover |

## Organization

Scripts are grouped by the task or system they support rather than by programming language. This keeps related operational workflows together even when they use different tools.

A script should move into its own repository when it grows into a standalone project with its own dependencies, tests, documentation, release cycle, or independent users.

## Usage and Safety

Review every script before running it in your environment. Administrative scripts can modify users, computers, printers, profiles, and other system state.

- Test scripts in a non-production environment first.
- Confirm required permissions and dependencies.
- Review configurable values and assumptions.
- Use appropriate change-control and backup procedures.
