# CMDT - Run as TrustedInstaller

![cmdt](images/cmdt.gif)

**The smallest fully functional TrustedInstaller elevation tool for Windows, written entirely in bare-metal x86/x64 assembly.**

CMDT launches any process under the **NT SERVICE\TrustedInstaller** security context — the highest privilege level in Windows, above both Administrator and SYSTEM. It enables all 34 Windows security privileges in the spawned process token, giving unrestricted access to every protected resource on the system.

The entire tool compiles to **under 25 KB** (x64) and **under 20 KB** (x86). No C runtime. No frameworks. No external dependencies beyond the Windows kernel and a handful of system DLLs that ship with every Windows installation since Vista.

---

## Why CMDT?

Windows protects critical system files, registry keys, and services with TrustedInstaller ownership. Even a process running as `NT AUTHORITY\SYSTEM` cannot modify these resources without taking ownership first — a destructive, auditable, and often irreversible operation.

CMDT solves this by spawning processes that **natively run as TrustedInstaller**, with the full set of 34 security privileges already enabled. No ownership changes needed. No ACL modifications. The process simply *is* the owner.

### Typical use cases

- Replacing or patching protected system binaries (WinSxS, System32)
- Modifying TrustedInstaller-owned registry keys without taking ownership
- Deleting stubborn files locked behind TrustedInstaller ACLs
- Scripted system maintenance where `SYSTEM` context is insufficient
- Debugging and forensics on protected OS components
- Repairing corrupted Windows installations at the deepest level

---

## Architecture

CMDT is a **dual-mode binary** — a single executable that operates as both a graphical desktop application and a headless command-line tool, selected at runtime based on arguments. This is not two programs stitched together; the same PE binary, the same entry point, and the same token acquisition pipeline serve both modes. The architecture is sometimes called a **hybrid subsystem** design: the executable uses the Windows subsystem (`/subsystem:windows`) but dynamically attaches to the parent console when invoked with CLI flags.

Both architectures — **x86 (IA-32)** and **x64 (AMD64)** — are built from separate, hand-written assembly source trees. No cross-compilation, no `#ifdef` macros, no shared C code. Each target is native assembly tuned to its calling convention and register set.

| Binary | Size | Architecture |
|---|---|---|
| `cmdt_x64.exe` | **under 25 KB** | x64 / AMD64 |
| `cmdt_x86.exe` | **under 20 KB** | x86 / IA-32 |

For comparison, equivalent tools written in C++ or C# typically weigh in at 50–500 KB, pulling in the CRT, .NET runtime, or static libraries. CMDT achieves full feature parity — GUI with MRU history, shortcut resolution, drag-and-drop, DPI awareness, CLI with I/O redirection, Explorer context menu integration, UAC self-elevation — in under 25 KB. This is possible only because every byte is hand-placed assembly, every API call is direct, and there is zero abstraction overhead.

---

## Features

- **Dual-mode operation** — GUI and CLI from a single binary, selected at runtime
- **UAC self-elevation** — automatically prompts for admin rights via `ShellExecuteEx("runas")` if not already elevated, forwarding all original command-line arguments to the elevated instance
- **Explorer context menu integration** — `cmdt -install` registers right-click entries for directories, executables, and shortcuts; `cmdt -uninstall` removes them (see [Context Menu Integration](#context-menu-integration))
- **CLI help** — passing an unknown switch (e.g. `cmdt -help`) prints all available options to the parent console via `AttachConsole` + `WriteConsoleW`
- **All 34 security privileges** enabled in the spawned token (see [Privilege Composition](#privilege-composition))
- **Token caching** — 30-second TTL avoids redundant privilege escalation on repeated runs
- **MRU history** — last 5 commands persisted in the registry, available in a dropdown
- **Windows shortcut (.lnk) resolution** — via COM (`IShellLinkW` + `IPersistFile`), both path and arguments
- **Drag-and-drop** with UIPI bypass — accepts drops from non-elevated Explorer windows
- **DPI-aware** — PerMonitorV2 via application manifest, sharp rendering on mixed-DPI setups
- **Modern visual styles** — Common Controls v6 through SxS manifest dependency
- **Resilient service startup** — retry loop with up to 2-second backoff when TrustedInstaller service is cold
- **I/O handle inheritance** — CLI mode preserves stdin/stdout/stderr for piping and redirection
- **Zero CRT dependency** — all string operations (copy, concatenate, compare, length) are hand-written wide-character routines
- **Proper environment block** — `CreateEnvironmentBlock` generates the correct TrustedInstaller environment for the child process

---

## Installation

No installation required. Copy `cmdt_x64.exe` (or `cmdt_x86.exe` for 32-bit systems) anywhere on your system. A natural location is `C:\Windows\System32` — this is where Microsoft places its own system utilities, and it makes CMDT available from any command prompt without modifying `PATH`.

CMDT requires Administrator privileges. If launched without elevation, it **automatically re-launches itself** with a UAC prompt via `ShellExecuteEx("runas")`, forwarding all original arguments to the elevated instance. No manual "Run as Administrator" is needed.

To register Explorer context menu entries, run:

```
cmdt -install
```

This creates right-click menu items for directories, `.exe` files, and `.lnk` shortcuts. See [Context Menu Integration](#context-menu-integration) for details.

### Requirements

- Windows 10 / Windows 11 (or Windows Vista+ with reduced feature set)
- Administrator privileges (Run as Administrator)
- TrustedInstaller service present (ships with all desktop and server editions of Windows)

---

## Usage

### GUI Mode

Launch `cmdt_x64.exe` without arguments to open the graphical interface.

```
cmdt_x64.exe
```

The window provides:

- **ComboBox** with dropdown — type a command or select from the MRU history (last 5 commands, persisted across sessions in `HKCU\Software\cmdt`)
- **Browse...** button — opens a file picker filtered to executables (`.exe`, `.lnk`)
- **Run** button — launches the command as TrustedInstaller
- **Status bar** — displays "Ready", "Launching...", "Process OK", or "Failed"
- **Drag-and-drop** — drop any `.exe` or `.lnk` file onto the window; it resolves the target and runs it immediately
- **Keyboard** — `Enter` runs the current command, `Escape` closes the window

The GUI dynamically relays out on resize. Controls stretch and reposition to fill the available client area.

### CLI Mode

Prefix any command with `-cli` (or `--cli` or `cli`) to run headless, inheriting the parent console's standard handles.

```
cmdt_x64.exe -cli <command>
cmdt_x64.exe -cli -new <command>
cmdt_x64.exe -install
cmdt_x64.exe -uninstall
```

| Switch | Description |
|---|---|
| `-cli <command>` | Run command as TrustedInstaller, inheriting the current console |
| `-cli -new <command>` | Run command in a new, separate console window |
| `-install` | Register Explorer context menu entries under HKCR |
| `-uninstall` | Remove all CMDT context menu entries |
| `(unknown switch)` | Display available options to the console |
| `(no arguments)` | Launch GUI mode |

#### Basic examples

```bash
# Launch an interactive TrustedInstaller command prompt
cmdt_x64.exe -cli cmd

# Open Registry Editor as TrustedInstaller
cmdt_x64.exe -cli regedit.exe

# Run PowerShell as TrustedInstaller
cmdt_x64.exe -cli powershell

# Launch a specific executable with full path
cmdt_x64.exe -cli notepad
```

#### I/O redirection — why it matters for scripting

In CLI mode without the `-new` flag, CMDT inherits the parent process's stdin, stdout, and stderr handles. This means the spawned TrustedInstaller process writes to the **same console and the same pipe** as the caller. Standard shell redirection works exactly as expected:

```bash
cmdt_x64.exe -cli cmd /c whoami > output.txt
```

This writes the output of `whoami` (running as TrustedInstaller) directly into `output.txt` in the caller's working directory. The redirection is handled by the parent shell before CMDT even launches — CMDT simply inherits the redirected handle and passes it through to the child process via `STARTUPINFO.hStdOutput`.

This makes CMDT suitable for **unattended automation scripts**, batch files, and CI/CD pipelines where capturing TrustedInstaller-level output is necessary:

```batch
@echo off
cmdt_x64.exe -cli cmd /c icacls "C:\Windows\servicing" > acl_report.txt
cmdt_x64.exe -cli cmd /c reg query "HKLM\SYSTEM\CurrentControlSet" /s > reg_dump.txt
cmdt_x64.exe -cli cmd /c dir "C:\Windows\WinSxS\*.manifest" /s > manifests.txt
```

Without handle inheritance, these commands would open orphaned console windows and the output would be lost. CMDT's explicit `GetStdHandle` + `STARTF_USESTDHANDLES` pipeline ensures that redirected output flows correctly through the TrustedInstaller boundary.

#### The `-new` flag — detached console

Add `-new` between the CLI switch and the command to spawn the process with `CREATE_NEW_CONSOLE`. The child gets its own independent console window:

```bash
# Open a new, standalone TrustedInstaller command prompt window
cmdt_x64.exe -cli -new cmd

# Open PowerShell in its own window as TrustedInstaller
cmdt_x64.exe -cli -new powershell
```

The difference is architectural:

| | `-cli` (default) | `-cli -new` |
|---|---|---|
| Console | Inherits parent | Creates new window |
| stdout/stderr | Shared with caller | Independent |
| I/O redirection | Works (`> file.txt`) | Not applicable |
| Use case | Scripting, automation | Interactive sessions |
| Creation flags | `CREATE_UNICODE_ENVIRONMENT` | `CREATE_NEW_CONSOLE \| CREATE_UNICODE_ENVIRONMENT` |
| STARTUPINFO | `STARTF_USESTDHANDLES` | `STARTF_USESHOWWINDOW` |

#### Shortcut (.lnk) resolution

CMDT transparently resolves Windows shortcuts. If the target path ends with `.lnk`, the tool initializes COM, instantiates `CLSID_ShellLink`, loads the shortcut via `IPersistFile::Load`, and extracts both the target path (`IShellLinkW::GetPath`) and embedded arguments (`IShellLinkW::GetArguments`). The resolved target and its arguments are concatenated and passed to `CreateProcessWithTokenW`.

This works in both GUI and CLI modes, and correctly handles quoted paths with spaces:

```bash
# Resolve shortcut and run the target as TrustedInstaller
cmdt_x64.exe -cli "C:\Users\Public\Desktop\Some App.lnk"

# Also works with drag-and-drop in GUI mode
```

The `.lnk` extension check is case-insensitive, implemented via a hand-written wide-character comparator that folds ASCII uppercase to lowercase inline.

---

## Context Menu Integration

Running `cmdt -install` registers four context menu entries under `HKEY_CLASSES_ROOT`:

| Registry Path | Menu Text | Behavior |
|---|---|---|
| `Directory\Background\shell\CMDT` | Open CMD as TrustedInstaller | Right-click on desktop or inside any folder |
| `Directory\shell\CMDT` | Open CMD as TrustedInstaller | Right-click on a folder icon |
| `exefile\shell\CMDT` | Run as TrustedInstaller | Right-click on any `.exe` file |
| `lnkfile\shell\CMDT` | Run as TrustedInstaller | Right-click on any `.lnk` shortcut |

### How it works

- **Directory entries** execute: `"<exepath>" -cli -new cmd.exe /k cd /d "%V"` — this opens a TrustedInstaller command prompt in the selected directory.
- **File entries** execute: `"<exepath>" "%1"` — CMDT receives the file path as an argument. For `.exe` files, it runs the executable directly. For `.lnk` shortcuts, it resolves the target via the COM `IShellLink` interface before execution.

Each entry displays a UAC shield icon borrowed from `shell32.dll` (icon index 104 — the "keys" icon). The binary itself contains **no embedded icon resource** (no `.ico` file). This is the same approach Microsoft uses for its own system utilities that reside in `System32`. Since CMDT is designed to live in `System32` by default, it does not need a standalone icon — Explorer resolves the `shell32.dll,104` reference at display time.

### Removal

```
cmdt -uninstall
```

This deletes all eight registry keys (four parent keys + four `command` subkeys) in leaf-first order, as the Windows registry does not allow deletion of keys that still contain subkeys.

---

## How It Works — Token Inheritance Chain

CMDT performs a multi-stage privilege escalation to obtain a fully privileged TrustedInstaller token. Each stage builds on the previous one, forming an inheritance chain:

### Stage 0: UAC Self-Elevation

Before any token work begins, CMDT checks `IsUserAnAdmin()`. If the process is not running elevated, it re-launches itself via `ShellExecuteExW` with the `"runas"` verb, forwarding the original command-line arguments to the new instance. The non-elevated process then exits immediately. This makes CMDT self-elevating — the user never needs to manually "Run as Administrator".

### Stage 1: Self-Elevation

The process enables two critical privileges in its own token using `AdjustTokenPrivileges`:

- **SeDebugPrivilege** — required to open process handles across security boundaries
- **SeImpersonatePrivilege** — required to impersonate another user's token

These are available because the process runs as Administrator (elevated).

### Stage 2: SYSTEM Impersonation

CMDT locates `winlogon.exe` by enumerating the process list via `CreateToolhelp32Snapshot` + `Process32FirstW/NextW`. The `winlogon.exe` process runs as `NT AUTHORITY\SYSTEM`.

The tool opens `winlogon.exe` with `PROCESS_QUERY_INFORMATION | PROCESS_DUP_HANDLE`, extracts its process token with `OpenProcessToken`, and duplicates it with `DuplicateTokenEx` at `MAXIMUM_ALLOWED` access and `SecurityImpersonation` level. It then calls `ImpersonateLoggedOnUser` to assume SYSTEM identity on the current thread.

After this stage, the calling thread runs as SYSTEM — necessary to interact with the Service Control Manager and the TrustedInstaller service process.

### Stage 3: TrustedInstaller Service Activation

CMDT opens the Service Control Manager and queries the **TrustedInstaller** service (`OpenServiceW` with `SERVICE_QUERY_STATUS | SERVICE_START`).

If the service is stopped, it calls `StartServiceW` and enters a **retry loop**: up to 10 iterations with 200 ms sleep intervals (~2 seconds total). Each iteration re-queries the service status via `QueryServiceStatusEx`. This resilient approach handles slow or heavily loaded machines where a single 200 ms wait would be insufficient.

Once `SERVICE_RUNNING` is confirmed, the service's **Process ID** is extracted from the `SERVICE_STATUS_PROCESS` structure (offset `dwProcessId`).

### Stage 4: Token Duplication

CMDT opens the TrustedInstaller process with `PROCESS_QUERY_INFORMATION`, extracts its token, and duplicates it via `DuplicateTokenEx` with `MAXIMUM_ALLOWED` access. This duplicated token becomes the foundation for the child process.

### Stage 5: Full Privilege Enablement

The duplicated token has all 34 Windows security privileges enabled via a loop that calls `LookupPrivilegeValueW` + `AdjustTokenPrivileges` for each privilege in the table (see [Privilege Composition](#privilege-composition) below).

### Stage 6: Process Creation

The fully privileged token is passed to `CreateProcessWithTokenW`. CMDT generates a proper environment block via `CreateEnvironmentBlock` (keyed to the TrustedInstaller token) and sets the working directory to `GetSystemDirectoryW`. The spawned process runs natively as TrustedInstaller with all 34 privileges enabled.

### Token Caching

The duplicated, fully privileged token is cached in memory with a 30-second TTL (tracked via `GetTickCount`). Subsequent invocations within the TTL window skip stages 2–5 entirely and reuse the cached token. This dramatically reduces overhead when running multiple commands in sequence — the expensive service startup, process enumeration, and privilege loop execute only once.

---

## Privilege Composition

CMDT enables all **34** Windows security privileges in the spawned token. This is the complete set that exists in the TrustedInstaller token:

| # | Privilege | Description |
|---|---|---|
| 0 | SeAssignPrimaryTokenPrivilege | Replace process-level token |
| 1 | SeBackupPrivilege | Bypass ACLs for read access (backup) |
| 2 | SeRestorePrivilege | Bypass ACLs for write access (restore) |
| 3 | SeDebugPrivilege | Debug any process |
| 4 | SeImpersonatePrivilege | Impersonate a client after authentication |
| 5 | SeTakeOwnershipPrivilege | Take ownership of any securable object |
| 6 | SeLoadDriverPrivilege | Load and unload device drivers |
| 7 | SeSystemEnvironmentPrivilege | Modify firmware environment variables |
| 8 | SeManageVolumePrivilege | Perform volume maintenance tasks |
| 9 | SeSecurityPrivilege | Manage auditing and security log |
| 10 | SeShutdownPrivilege | Shut down the system |
| 11 | SeSystemtimePrivilege | Change the system time |
| 12 | SeTcbPrivilege | Act as part of the operating system |
| 13 | SeIncreaseQuotaPrivilege | Adjust memory quotas for a process |
| 14 | SeAuditPrivilege | Generate security audits |
| 15 | SeChangeNotifyPrivilege | Bypass traverse checking |
| 16 | SeUndockPrivilege | Remove computer from docking station |
| 17 | SeCreateTokenPrivilege | Create a token object |
| 18 | SeLockMemoryPrivilege | Lock pages in memory |
| 19 | SeCreatePagefilePrivilege | Create a pagefile |
| 20 | SeCreatePermanentPrivilege | Create permanent shared objects |
| 21 | SeSystemProfilePrivilege | Profile system performance |
| 22 | SeProfileSingleProcessPrivilege | Profile a single process |
| 23 | SeCreateGlobalPrivilege | Create global objects |
| 24 | SeTimeZonePrivilege | Change the time zone |
| 25 | SeCreateSymbolicLinkPrivilege | Create symbolic links |
| 26 | SeIncreaseBasePriorityPrivilege | Increase scheduling priority |
| 27 | SeRemoteShutdownPrivilege | Force shutdown from a remote system |
| 28 | SeIncreaseWorkingSetPrivilege | Increase a process working set |
| 29 | SeRelabelPrivilege | Modify an object label |
| 30 | SeDelegateSessionUserImpersonatePrivilege | Obtain impersonation token for another user in same session |
| 31 | SeTrustedCredManAccessPrivilege | Access Credential Manager as a trusted caller |
| 32 | SeEnableDelegationPrivilege | Enable computer and user accounts to be trusted for delegation |
| 33 | SeSyncAgentPrivilege | Synchronize directory service data |

### Binary-level string decomposition

The privilege names are not stored as complete strings in the binary. Instead, CMDT uses a **prefix-suffix decomposition technique** that splits every privilege name into three parts:

| Part | Value | Storage |
|---|---|---|
| Prefix | `Se` | Single shared constant |
| Middle | e.g. `Debug`, `Backup`, `TakeOwnership` | Per-privilege unique string |
| Suffix | `Privilege` | Single shared constant |

At runtime, the `BuildPrivilegeName` procedure concatenates these three parts into a temporary buffer before passing the result to `LookupPrivilegeValueW`. The full name `SeDebugPrivilege` is assembled in memory but **never appears as a contiguous string in the binary image**.

This decomposition has two engineering consequences:

1. **Size reduction** — The prefix (`Se`, 4 bytes UTF-16) and suffix (`Privilege`, 18 bytes UTF-16) are stored once instead of 34 times, saving approximately 750 bytes. In a 20 KB binary, that is nearly 4% of the total size.

2. **Static analysis opacity** — Automated scanners and signature-based tools that grep for known privilege strings like `SeDebugPrivilege` or `SeTcbPrivilege` will find **no matches** in the binary. The strings `Se` and `Privilege` appear separately, and the middle parts (`Debug`, `Tcb`, `Backup`, etc.) are generic English words that carry no security significance on their own. This is not obfuscation — it is a natural consequence of factoring out common substrings in a size-constrained binary. But the side effect is significant: the binary's static footprint does not betray the scope of privileges it enables.

---

## Manifest and DPI Awareness

CMDT embeds a Win32 application manifest that declares three important capabilities:

### PerMonitorV2 DPI Awareness

The manifest declares both the legacy `dpiAware=true` attribute (for Vista–8.1 compatibility) and the modern `dpiAwareness=PerMonitorV2` attribute (Windows 10 1703+). On modern systems, this means:

- The window renders at native resolution on every monitor — no bitmap scaling or blurriness
- When dragged between monitors with different DPI settings, the window rescales correctly
- Text, buttons, and controls render sharp on 4K, ultrawide, and mixed-DPI configurations

This is the same DPI awareness model used by modern Windows applications like Explorer, Edge, and Terminal.

### Common Controls v6 (Visual Styles)

The manifest declares a Side-by-Side (SxS) dependency on `Microsoft.Windows.Common-Controls` version 6.0. This activates the modern visual theme for all standard controls — the ComboBox dropdown, buttons, and static labels render with the current Windows theme (Fluent, Aero, or Classic) rather than the legacy Win95 appearance.

### Execution Level

The manifest specifies `requestedExecutionLevel=asInvoker`. CMDT does not rely on the manifest for elevation — instead, it programmatically checks `IsUserAnAdmin()` at startup and re-launches itself via `ShellExecuteExW("runas")` if not elevated. This approach allows the same binary to be invoked silently from already-elevated contexts (scripts, scheduled tasks, elevated terminals) without triggering a redundant UAC prompt, while still self-elevating when launched from a standard user session.

---

## MRU (Most Recently Used) History

The GUI maintains a persistent **MRU list of the last 5 commands** in the Windows registry at `HKEY_CURRENT_USER\Software\cmdt`. Values are stored as named entries `0` through `4`, where `0` is the most recent command.

On startup, `LoadMRU` reads these values and populates the ComboBox dropdown. After each successful execution, `SaveMRU` shifts existing entries down (0→1, 1→2, ..., 3→4), deletes the oldest entry, and writes the new command at position 0. Duplicate detection is implicit — the shift operation naturally pushes older duplicates off the end of the list.

The MRU list persists across sessions, reboots, and updates. It is the only state CMDT writes to disk (via the registry).

---

## Drag-and-Drop with UIPI Bypass

CMDT accepts drag-and-drop of `.exe` and `.lnk` files. Dropping a file onto the window sets the command text and immediately executes it.

Because CMDT runs elevated (as Administrator), the default Windows behavior is to **block** drag-and-drop messages from non-elevated processes like Explorer. This is enforced by **User Interface Privilege Isolation (UIPI)** — a security boundary that prevents lower-integrity processes from sending messages to higher-integrity windows.

CMDT explicitly bypasses this restriction by calling `ChangeWindowMessageFilterEx` for both `WM_DROPFILES` and `WM_COPYGLOBALDATA` on the main window handle. This whitelists these specific messages, allowing drops from standard Explorer windows while maintaining the UIPI boundary for all other message types.

---

## Building from Source

### Prerequisites

- **Microsoft Macro Assembler** — `ml.exe` (x86) and `ml64.exe` (x64) from Visual Studio Build Tools
- **Windows SDK** — for `rc.exe` (resource compiler), import libraries, and headers
- **PowerShell** — for the build script

### Build

```powershell
.\build.ps1
```

The build script assembles all four source modules (`main`, `token`, `process`, `window`) for both architectures, compiles the resource file (`cmdt.rc`) with the manifest, and links against system import libraries only:

`kernel32.lib`, `user32.lib`, `advapi32.lib`, `shell32.lib`, `comdlg32.lib`, `ole32.lib`, `gdi32.lib`, `shlwapi.lib`, `userenv.lib`

No CRT library is linked. The entry point is `mainCRTStartup` (x64) / `start` (x86) — these are raw assembly procedures, not CRT initialization stubs.

Output binaries are placed in the `bin\` directory.

---

## Project Structure

```
cmdt_asm/
├── x64/                    # AMD64 assembly sources
│   ├── main.asm            # Entry point, CLI/GUI dispatch, privilege table
│   ├── token.asm           # Token acquisition, SYSTEM impersonation, service control
│   ├── process.asm         # CreateProcessWithTokenW wrapper
│   ├── window.asm          # GUI, MRU, drag-and-drop, .lnk resolution via COM
│   ├── consts.inc          # Windows API constants, control IDs, message codes
│   └── globals.inc         # External symbol declarations
├── x86/                    # IA-32 assembly sources (parallel structure)
│   ├── main.asm
│   ├── token.asm
│   ├── process.asm
│   ├── window.asm
│   ├── consts.inc
│   └── globals.inc
├── bin/                    # Compiled binaries
│   ├── cmdt_x64.exe        # 64-bit binary (<25 KB)
│   └── cmdt_x86.exe        # 32-bit binary (<20 KB)
├── cmdt.rc                 # Version info resource
├── cmdt.manifest           # Application manifest (DPI, visual styles, execution level)
├── build.ps1               # Build script (assembles + links both architectures)
└── README.md               # This file (documentation)
```

Every source file in `x64/` has a corresponding counterpart in `x86/`. The x86 versions use `.586` + `flat/stdcall` MASM syntax with `invoke` macros; the x64 versions use raw `proc frame` with explicit SEH prologue/epilogue annotations (`.pushreg`, `.allocstack`, `.setframe`, `.endprolog`). Both targets share the same `.rc` and `.manifest` files.

---

## String Operations — No CRT

CMDT implements all necessary string operations as hand-written wide-character (UTF-16LE) assembly routines. There is no dependency on `msvcrt.dll`, `ucrtbase.dll`, or any C runtime:

| Function | Purpose |
|---|---|
| `wcscpy_p` / `wcscpy_t` | Wide string copy |
| `wcscat_p` / `wcscat_w` | Wide string concatenation |
| `wcscmp_ci` / `wcscmp_ci_w` | Case-insensitive wide string comparison |
| `wcslen_p` / `wcslen_w` | Wide string length |
| `skip_spaces` | Skip leading whitespace in command parsing |

The `_p` / `_t` variants live in `main.asm` and `token.asm` respectively; the `_w` variants live in `window.asm`. Each is a tight loop operating on 16-bit words, with inline ASCII case folding for the comparison functions (uppercase A–Z folded to lowercase by adding 32 to the code point).

---

## Verification

After launching a command prompt as TrustedInstaller:

```
cmdt_x64.exe -cli cmd
```

Verify the security context:

```
C:\Windows\System32> whoami
nt service\trustedinstaller

C:\Windows\System32> whoami /priv
```

All 34 privileges should appear with state **Enabled**.

---

## Security Considerations

**TrustedInstaller is the highest privilege level in Windows** — higher than Administrator, higher than SYSTEM. A process running as TrustedInstaller can:

- Modify or delete any file on the system, including protected OS components
- Write to any registry key, including those owned by TrustedInstaller
- Load and unload kernel drivers
- Access and modify the firmware environment (UEFI variables)
- Debug any process, including critical system processes
- Create token objects and impersonate any security principal

Use CMDT with the same caution you would apply to a kernel debugger. Mistakes at this privilege level can render the operating system unbootable.

CMDT requires Administrator privileges to run. It does not bypass UAC — the user must explicitly elevate the process before CMDT can acquire the TrustedInstaller token.

---

## License

MIT License

Copyright (c) 2026 Marek Wesolowski

---

## Author

**Marek Wesolowski**
- Web: [https://kvc.pl](https://kvc.pl)
- E-mail: marek@kvc.pl

---

## Size Trivia

During early development, the minimal proof-of-concept builds were significantly smaller:

| Variant | Size | Notes |
|---|---|---|
| CLI-only (no GUI, no registry, no manifest) | **4 KB** | Bare token acquisition + `CreateProcessWithTokenW` |
| Hybrid GUI/CLI (no registry, no manifest) | **6 KB** | Added window creation, MRU, drag-and-drop |
| Current full build (x86) | **<20 KB** | Hybrid mode, context menu, UAC self-elevation, manifest, COM `.lnk` resolution |
| Current full build (x64) | **<25 KB** | Same feature set, 64-bit calling convention overhead |

The growth from 4–6 KB to the current size is almost entirely due to the application manifest (DPI awareness, Common Controls v6, execution level declaration), the context menu registry logic, UAC self-elevation, and the wide-character string constants for registry paths and UI text. The core token acquisition pipeline — the actual "engine" of CMDT — remains remarkably compact.

---

*Written in 100% bare-metal x86/x64 MASM assembly. No frameworks. No runtimes. No compromises.*
