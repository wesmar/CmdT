# CMDT - Run as TrustedInstaller

![cmdt](images/cmdt.gif)

**A compact TrustedInstaller process launcher for Windows, written entirely in native x86/x64 assembly.**

CMDT starts a process with a primary token duplicated from the running **Windows Modules Installer (TrustedInstaller)** service. That token normally identifies its user as `NT AUTHORITY\SYSTEM` and carries the enabled `NT SERVICE\TrustedInstaller` service SID used by ACLs protecting Windows components. CMDT also attempts to enable its table of 34 Windows token privileges before process creation; Windows still remains the final authority on privileges present in the source token and on every object-specific ACL check.

The current release binaries are **38.00 KiB (38,912 bytes)** on x64 and **29.50 KiB (30,208 bytes)** on x86, with enforced build limits of under 40 KiB and under 30 KiB respectively. No C runtime, framework, or third-party runtime is required.

---

## Why CMDT?

Windows protects critical system files, registry keys, and services with TrustedInstaller ownership. Even a process running as `NT AUTHORITY\SYSTEM` cannot modify these resources without taking ownership first — a destructive, auditable, and often irreversible operation.

CMDT solves this by duplicating the TrustedInstaller service process token and using it to create the requested process. The enabled TrustedInstaller service SID satisfies ACL entries that explicitly grant access to that service identity, avoiding the usual ownership and ACL rewrites. This is token-based access, not a new Windows privilege tier above SYSTEM, and it does not bypass an explicit deny ACE or every other security boundary.

### Typical use cases

- Replacing or patching protected system binaries (WinSxS, System32)
- Modifying TrustedInstaller-owned registry keys without taking ownership
- Deleting stubborn files locked behind TrustedInstaller ACLs
- Scripted system maintenance where `SYSTEM` context is insufficient
- Debugging and forensics on protected OS components
- Repairing corrupted Windows installations at the deepest level

---

## Architecture

CMDT is a **dual-mode binary** — a single executable that operates as both a graphical desktop application and a headless command-line tool, selected at runtime based on arguments. This is not two programs stitched together; the same PE binary, the same entry point, and the same token acquisition pipeline serve both modes.

Both architectures link as `/subsystem:console`, not `/subsystem:windows`. This is a deliberate trade-off: a console-subsystem process is what `cmd.exe` actually waits on synchronously before printing its next prompt — the PE subsystem field is checked by `cmd.exe` at `CreateProcess` time, before any of the child's own code runs, so it cannot be influenced at runtime. A GUI-subsystem binary is never waited on, which is exactly why some older TrustedInstaller launchers exhibit a glued-together prompt or a stray blank line after `-cli` output: the shell already redrew its prompt before the child finished writing. Console subsystem makes `-cli` behave like any other console tool, with correct output ordering and no cosmetic redraw hacks needed.

The cost of that choice is the default (no-argument) GUI launch: the OS attaches a console window to any console-subsystem process before its code runs, which would otherwise cause a visible flash on every double-click launch from Explorer. CMDT avoids it without spawning a second process — at startup, if no arguments were given, it calls `GetConsoleProcessList` to check whether the console belongs **exclusively** to this process (the common case: Explorer, a shortcut, the Run dialog) or is **shared** with an interactive parent shell (the user typed `cmdt` inside an already-open `cmd.exe`). In the exclusive case it calls `GetConsoleWindow` + `ShowWindow(SW_HIDE)` immediately, before falling into the GUI message loop — no `CreateProcessW`, no second process, just two Win32 calls between process start and the window being hidden. In the shared case it leaves the console alone entirely; hiding it would hide the user's whole terminal. Verified via `EnumWindows` + `IsWindowVisible` on the console's own `ConsoleWindowClass` window in both scenarios.

Both architectures — **x86 (IA-32)** and **x64 (AMD64)** — are built from separate, hand-written assembly source trees. No cross-compilation, no `#ifdef` macros, no shared C code. Each target is native assembly tuned to its calling convention and register set.

| Binary | Current release size | Enforced limit | Architecture |
|---|---:|---:|---|
| `cmdt_x64.exe` | **38.00 KiB (38,912 bytes)** | **<40 KiB** | x64 / AMD64 |
| `cmdt_x86.exe` | **29.50 KiB (30,208 bytes)** | **<30 KiB** | x86 / IA-32 |

For comparison, equivalent tools written in C++ or C# typically weigh in at 50–500 KB, pulling in the CRT, .NET runtime, or static libraries. CMDT achieves full feature parity — GUI with MRU history, shortcut resolution, drag-and-drop, DPI awareness, CLI with I/O redirection, Explorer context menu integration, Sticky Keys IFEO hook, Defender exclusion management, UAC self-elevation — in well under 40 KB on x64 and 30 KB on x86. This is possible only because every byte is hand-placed assembly, every API call is direct, and there is zero abstraction overhead.

---

## Features

- **Dual-mode operation** — GUI and CLI from a single binary, selected at runtime
- **Shift+Minimize to tray** — minimizing normally keeps CMDT on the taskbar; hold **Shift** while clicking Minimize to hide it in the notification area. Double-click the tray icon or choose **Restore** from its context menu to bring the window back; **Exit** closes CMDT directly from the tray
- **UAC self-elevation** — automatically prompts for admin rights via `ShellExecuteEx("runas")` if not already elevated, forwarding all original command-line arguments to the elevated instance
- **Explorer context menu integration** — `cmdt -install` registers right-click entries for directories, executables, and shortcuts; `cmdt -uninstall` removes them (see [Context Menu Integration](#context-menu-integration))
- **Sticky Keys IFEO hook** — `cmdt -shift` installs an Image File Execution Options debugger redirect for `sethc.exe`, so pressing Shift 5 times at the login screen opens a TrustedInstaller command prompt instead of Sticky Keys; `cmdt -unshift` reverts to default behavior (see [Sticky Keys IFEO Hook](#sticky-keys-ifeo-hook))
- **Windows Defender exclusions** — `-shift` and `-unshift` automatically add or remove process exclusions for the CMDT binary and `cmd.exe` via WMI (`MSFT_MpPreference` COM interface, `ROOT\Microsoft\Windows\Defender` namespace), preventing false-positive interference without spawning a PowerShell process
- **CLI help** — `-h`, `-help`, `--help`, `-?`, `/?`, `/h`, `/help` all print the usage banner; the check runs **before** UAC self-elevation so output always reaches the original shell — both interactive (`cmdt -help`) and redirected (`cmdt -help > out.txt`) work correctly in elevated and non-elevated sessions
- **CLI output relay** — running `cmdt -cli <command>` correctly delivers stdout/stderr to the caller's redirect target (`>> out.txt`, `| pipe`, etc.) whether the caller is already elevated or not; a temp-file relay bridges the gap transparently in both cases (see [I/O Redirection](#io-redirection--why-it-matters-for-scripting))
- **34-entry privilege enablement table** — CMDT requests each privilege with `AdjustTokenPrivileges`; Windows enables only privileges already present in the duplicated service token (see [Privilege Composition](#privilege-composition))
- **Token caching** — 30-second TTL avoids redundant privilege escalation on repeated runs
- **Opt-in MRU history, off by default** — no registry key exists until explicitly enabled via **File → Enable History** in the GUI; last 5 commands available in a dropdown once turned on; `cmdt -history-clear` wipes it from the CLI (see [MRU History](#mru-most-recently-used-history))
- **Windows shortcut (.lnk) resolution** — via COM (`IShellLinkW` + `IPersistFile`), both path and arguments
- **Drag-and-drop** with UIPI bypass — accepts drops from non-elevated Explorer windows
- **DPI-aware** — PerMonitorV2 via application manifest, sharp rendering on mixed-DPI setups; the DPI query APIs themselves are resolved dynamically at runtime (not statically imported) so the binary still loads and runs on Windows 7/8/8.1, just without per-monitor scaling (see [Manifest and DPI Awareness](#manifest-and-dpi-awareness))
- **System-aware dark mode** — reads `AppsUseLightTheme` from the registry and applies it to the title bar, Mica backdrop, ComboBox, buttons, status label, and native popup-menu theme. Updates on `WM_SETTINGCHANGE`/`WM_THEMECHANGED` without restarting. The classic top-level `HMENU` strip remains rendered by USER32; CMDT deliberately does not force the experimental owner-draw path that proved re-entrant on x64 (see [Dark Mode and Mica Backdrop](#dark-mode-and-mica-backdrop))
- **Mica backdrop** — `DWMWA_SYSTEMBACKDROP_TYPE = DWMSBT_MAINWINDOW` on Windows 11, combined with explicit light/dark client painting so minimize/restore never exposes an uninitialized DWM surface
- **Modern visual styles** — Common Controls v6 through SxS manifest dependency
- **Resilient service startup** — retry loop with up to 2-second backoff when TrustedInstaller service is cold
- **I/O handle inheritance** — CLI mode preserves stdin/stdout/stderr for piping and redirection
- **Zero CRT dependency** — all string operations (copy, concatenate, compare, length) are hand-written wide-character routines in `strutil.asm`
- **Proper environment block** — `CreateEnvironmentBlock` generates the correct TrustedInstaller environment for the child process

---

## Installation

No installation required. Copy `cmdt_x64.exe` (or `cmdt_x86.exe` for 32-bit systems) anywhere on your system. A natural location is `C:\Windows\System32` — this is where Microsoft places its own system utilities, and it makes CMDT available from any command prompt without modifying `PATH`.

CMDT requires Administrator privileges. If launched without elevation, it **automatically re-launches itself** with a UAC prompt via `ShellExecuteExW("runas")`, forwarding all original arguments to the elevated instance. No manual "Run as Administrator" is needed.

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

- **ComboBox** with dropdown — type a command or, once history is enabled, select from the MRU list (last 5 commands)
- **File menu** — Open with TrustedInstaller (browse), **Enable History** (checkbox — off by default, see [MRU History](#mru-most-recently-used-history)), About, Exit
- **Browse...** button — opens a file picker filtered to executables (`.exe`, `.lnk`)
- **Run** button — launches the command as TrustedInstaller
- **Status bar** — displays "Ready", "Launching...", "Process OK", or "Failed"
- **Drag-and-drop** — drop any `.exe` or `.lnk` file onto the window; it resolves the target and runs it immediately
- **Keyboard** — `Enter` runs the current command, `Escape` closes the window
- **Minimize** — click Minimize normally to keep CMDT on the taskbar, or hold **Shift** while clicking it to hide CMDT in the system tray. Double-click the tray icon to restore the window; right-click it for **Restore** and **Exit**

The GUI dynamically relays out on resize. Controls stretch and reposition to fill the available client area.

Tray handling is isolated in `tray.asm` for each architecture. The module owns the `NOTIFYICONDATAW` state, Shift-key detection, restore/exit menu, and icon cleanup. It also registers the `TaskbarCreated` message and allows that specific message through UIPI, so an elevated CMDT can republish its icon if Explorer restarts while the window is hidden.

### CLI Mode

Prefix any command with `-cli` (or `--cli` or `cli`) to run headless, inheriting the parent console's standard handles.

```
cmdt_x64.exe -cli <command>
cmdt_x64.exe -cli -new <command>
cmdt_x64.exe -install
cmdt_x64.exe -uninstall
cmdt_x64.exe -shift
cmdt_x64.exe -unshift
cmdt_x64.exe -history-clear
```

| Switch | Description |
|---|---|
| `-cli <command>` | Run command as TrustedInstaller, inheriting the current console |
| `-cli -new <command>` | Run command in a new, separate console window |
| `-install` | Register Explorer context menu entries under HKCR |
| `-uninstall` | Remove all CMDT context menu entries |
| `-shift` | Install Sticky Keys IFEO hook + Defender exclusions |
| `-unshift` | Remove Sticky Keys IFEO hook + Defender exclusions |
| `-history-clear` | Wipe the MRU command-history registry key, if present |
| `-h`, `-help`, `--help`, `-?`, `/?`, `/h`, `/help` | Display available options |
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
cmdt_x64.exe -cli net session >> out.txt
```

**Every `-cli` invocation goes through the same temp-file relay**, whether the caller is already elevated or not. Earlier builds only relayed the non-admin path and let an already-elevated shell call `CreateProcessWithTokenW` directly with `STARTF_USESTDHANDLES` — but `CreateProcessWithTokenW` does not reliably hand the spawned child usable inherited std handles, so an admin `cmd.exe` running `cmdt -cli net session > out.txt` silently produced an empty file ([issue #1](https://github.com/wesmar/CmdT/issues/1)). The relay now runs unconditionally for `-cli`:

1. The parent (admin or not) creates **two** unique temp files via `GetTempFileNameW` — one for stdout, one for stderr — so redirecting only one stream (`2>err.txt` without touching stdout, or vice versa) behaves correctly instead of interleaving both into a single stream.
2. It inserts internal `-outfile <path>` and `-errfile <path>` tokens into the argument string and launches an elevated copy of itself via `ShellExecuteExW("runas")` — a no-op prompt if already elevated, a real UAC prompt otherwise.
3. The elevated child opens both temp files with inheritable `GENERIC_WRITE` handles and uses them as the spawned process's `hStdOutput`/`hStdError` respectively; `hStdInput` is a handle to the `NUL` device rather than `NULL`, since some console-aware targets refuse to start with a null stdin handle. It waits on the child via `WaitForSingleObject` and captures its real exit code with `GetExitCodeProcess`.
4. After the elevated process exits, the parent opens each temp file, streams stdout to its own `STD_OUTPUT_HANDLE` and stderr to `STD_ERROR_HANDLE` (both wired up by cmd.exe before launch — so `>`, `>>`, `2>`, and `|` all work transparently and independently), deletes both temp files, and **exits with the same exit code the spawned command produced** rather than always returning 0. This matters for scripts that check `%ERRORLEVEL%`/`$LASTEXITCODE` after a `cmdt -cli` call — a failing command inside the relay now fails the caller's script too.

The relay declines — falling back to direct dispatch — for `-new` (a detached console has no output to capture), for interactive shells (`cmd`, `powershell`, `pwsh` invoked with no further arguments, which need a real attachable console), for the internal `-outfile`/`-errfile` relay child itself (guarding against relaying itself again on re-entry), and if temp-file creation fails.

One known limitation survives this fix: if the `-cli` command *is itself* `cmd.exe` with embedded I/O redirection (`cmd /c foo > file`, or a `.lnk` shortcut targeting that), `cmd.exe` refuses to run under the relay's `CREATE_NO_WINDOW` context with `"Input redirection is not supported, exiting the process immediately."` Plain commands and `.lnk` targets without their own embedded redirection are unaffected.

This makes CMDT suitable for **unattended automation scripts**, batch files, and CI/CD pipelines where capturing TrustedInstaller-level output is necessary:

```batch
@echo off
cmdt_x64.exe -cli cmd /c icacls "C:\Windows\servicing" > acl_report.txt
cmdt_x64.exe -cli cmd /c reg query "HKLM\SYSTEM\CurrentControlSet" /s > reg_dump.txt
cmdt_x64.exe -cli cmd /c dir "C:\Windows\WinSxS\*.manifest" /s > manifests.txt
cmdt_x64.exe -cli net session >> audit.txt
```

Without the relay, these commands would silently lose their output — orphaned console windows for a non-admin caller, or an empty file for an already-elevated one. With the relay applied unconditionally to `-cli`, both cases land the full output at the redirect target.

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

#### Help display

All of the following are recognized help switches and print the usage banner:

```
cmdt_x64.exe -help
cmdt_x64.exe -h
cmdt_x64.exe --help
cmdt_x64.exe -?
cmdt_x64.exe /?
cmdt_x64.exe /h
cmdt_x64.exe /help
```

The help check runs **before** UAC self-elevation. This matters because the elevated process is detached from the original shell's console and redirect targets. By printing usage from the non-elevated process, CMDT stays attached to the launching console — so both interactive display and `cmdt -help > out.txt` work correctly regardless of elevation state.

Output is routed via `GetFileType` on the stdout handle: `WriteConsoleW` for a real console (native UTF-16), `WriteFile` for a file or pipe (raw UTF-16 LE bytes). This ensures redirected output is never silently dropped.

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

## Sticky Keys IFEO Hook

Running `cmdt -shift` installs a login-screen backdoor that replaces the Sticky Keys accessibility helper (`sethc.exe`) with a TrustedInstaller command prompt. Running `cmdt -unshift` reverses the process completely.

### How it works

Windows supports **Image File Execution Options (IFEO)** — a documented registry mechanism under `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\` that allows attaching a debugger to any executable. When an IFEO `Debugger` value exists for a given binary, Windows launches the debugger *instead of* the original program, passing the original path as an argument.

`cmdt -shift` creates the following registry entry:

| Key | Value | Data |
|---|---|---|
| `HKLM\...\Image File Execution Options\sethc.exe` | `Debugger` | `<exepath> -cli -new cmd.exe` |

When the user presses **Shift five times** at the Windows login screen, the OS launches `sethc.exe` — but the IFEO redirect intercepts this and runs CMDT instead. CMDT then performs its full token acquisition pipeline (SYSTEM impersonation → TrustedInstaller service start → token duplication → privilege enablement) and opens an interactive `cmd.exe` window running as `NT SERVICE\TrustedInstaller`.

This provides a **pre-login recovery console** carrying the TrustedInstaller service token — useful for:

- Emergency system repair when the machine cannot be logged into
- Resetting locked-out local accounts via `net user`
- Fixing corrupted registry hives or group policies that prevent login
- Recovering from ransomware or malware that blocks the desktop

### Windows Defender exclusions

Both `-shift` and `-unshift` manage Windows Defender process exclusions automatically. This prevents Defender from flagging CMDT or the spawned `cmd.exe` as suspicious during the IFEO-redirected launch at the login screen, where no user session exists to dismiss alerts.

| Switch | Action |
|---|---|
| `-shift` | Adds `<filename>` and `cmd.exe` as process exclusions via WMI |
| `-unshift` | Removes `<filename>` and `cmd.exe` process exclusions via WMI |

Exclusions are managed by calling the `MSFT_MpPreference` WMI class methods (`Add` / `Remove`) directly through the COM interface — no PowerShell or child process is spawned. CMDT connects to the `ROOT\Microsoft\Windows\Defender` WMI namespace, instantiates `IWbemServices`, and invokes `ExecMethod` with a `SAFEARRAY` of `BSTR` values for the `ExclusionProcess` property. This approach is faster, produces no visible console window, and has no dependency on the PowerShell execution policy.

### Removal

```
cmdt -unshift
```

This deletes the `Debugger` value from the IFEO registry key (leaving the key itself intact, as it may contain other unrelated values) and removes both Defender process exclusions. Normal Sticky Keys behavior is fully restored.

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

CMDT locates `winlogon.exe` by enumerating the process list via `CreateToolhelp32Snapshot` + `Process32FirstW/NextW`, using a case-insensitive wide-character comparison (`wcscmp_ci`) against each entry's `szExeFile` field. The `winlogon.exe` process runs as `NT AUTHORITY\SYSTEM`.

The tool opens `winlogon.exe` with `PROCESS_QUERY_INFORMATION | PROCESS_DUP_HANDLE`, extracts its process token with `OpenProcessToken`, and duplicates it with `DuplicateTokenEx` at `MAXIMUM_ALLOWED` access and `SecurityImpersonation` level. It then calls `ImpersonateLoggedOnUser` to assume SYSTEM identity on the current thread.

After this stage, the calling thread runs as SYSTEM — necessary to interact with the Service Control Manager and the TrustedInstaller service process.

### Stage 3: TrustedInstaller Service Activation

CMDT opens the Service Control Manager and queries the **TrustedInstaller** service (`OpenServiceW` with `SERVICE_QUERY_STATUS | SERVICE_START`).

If the service is stopped, it calls `StartServiceW` and enters a **retry loop**: up to 10 iterations with 200 ms sleep intervals (~2 seconds total). Each iteration re-queries the service status via `QueryServiceStatusEx`. This resilient approach handles slow or heavily loaded machines where a single 200 ms wait would be insufficient.

Once `SERVICE_RUNNING` is confirmed, the service's **Process ID** is extracted from the `SERVICE_STATUS_PROCESS` structure (offset `dwProcessId`).

### Stage 4: Token Duplication

CMDT opens the TrustedInstaller process with `PROCESS_QUERY_INFORMATION`, extracts its token, and duplicates it via `DuplicateTokenEx` with `MAXIMUM_ALLOWED` access. This duplicated token becomes the foundation for the child process.

### Stage 5: Privilege Enablement

CMDT iterates a table of 34 privilege names and calls `LookupPrivilegeValueW` plus `AdjustTokenPrivileges` for each one. `AdjustTokenPrivileges` can enable a privilege already present in a token; it cannot add one the source token does not contain. The final set therefore depends on the TrustedInstaller service token supplied by that Windows version and configuration.

### Stage 6: Process Creation

The adjusted token is passed to `CreateProcessWithTokenW`. CMDT generates an environment block via `CreateEnvironmentBlock` and sets the working directory to `GetSystemDirectoryW`. The resulting process normally reports its user as `NT AUTHORITY\SYSTEM`; its token also carries the TrustedInstaller service SID used by protected-resource ACLs.

### Token Caching

The duplicated token is cached in memory with a 30-second TTL (tracked via `GetTickCount`). Repeated launches within the same long-lived CMDT process — principally GUI commands — can skip stages 2–5 and reuse it. Separate CLI invocations are separate processes and do not share this in-memory cache.

---

## Privilege Composition

CMDT contains the following **34-entry request table**. It attempts to enable each entry, but the resulting token can contain and enable only privileges present in the source TrustedInstaller service token:

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

1. **Size reduction** — The prefix (`Se`, 4 bytes UTF-16) and suffix (`Privilege`, 18 bytes UTF-16) are stored once instead of 34 times, saving approximately 750 bytes. In a sub-30 KB binary, that is roughly 2.5% of the total size.

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

Declaring `PerMonitorV2` in the manifest only stops Windows from doing automatic bitmap stretching — it does not scale anything by itself. `window_lifecycle.inc` queries `GetDpiForSystem()`/`GetDpiForWindow(hWnd)` at window creation, `WM_CREATE`, and `WM_SIZE`, scaling every design-time (96 DPI) coordinate by `value * dpi / 96` before use; the main window's total size is computed from the desired client area via `AdjustWindowRectExForDpi` rather than a guessed pixel constant, so caption/menu/border overhead is correct regardless of Windows theme or font scale. Both functions were introduced in Windows 10 1607 and are **resolved dynamically** — `InitDpiApis` calls `GetModuleHandleA("user32.dll")` + `GetProcAddress` at runtime instead of a static `EXTRN` import. A static import table entry the loader can't resolve fails the whole process to start with "the procedure entry point ... could not be located" — a real crash report from a Windows 7 user running an earlier build that imported these statically. Every call site checks the resulting function pointer for NULL and falls back to a fixed 96-DPI (100%) calculation when it's missing, so the binary still runs correctly on Windows 7/8/8.1 — just without per-monitor scaling, since the concept doesn't exist there.

### Common Controls v6 (Visual Styles)

The manifest declares a Side-by-Side (SxS) dependency on `Microsoft.Windows.Common-Controls` version 6.0. This activates the modern visual theme for all standard controls — the ComboBox dropdown, buttons, and static labels render with the current Windows theme (Fluent, Aero, or Classic) rather than the legacy Win95 appearance.

### Dark Mode and Mica Backdrop

CMDT reads the system app theme preference from the registry at startup and whenever the user changes the Windows color scheme:

```
HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize
  AppsUseLightTheme (DWORD) — 0 = dark, 1 = light
```

This drives three cooperating layers. Window and control attributes are reapplied on `WM_SETTINGCHANGE`/`WM_THEMECHANGED`, so a live theme toggle in Windows Settings does not require restarting CMDT:

**1. DWM chrome.** `DwmSetWindowAttribute` is called with `DWMWA_USE_IMMERSIVE_DARK_MODE` (20, falling back to attribute 19 on Windows 10 builds before 20H1) for the title bar, and `DWMWA_SYSTEMBACKDROP_TYPE = DWMSBT_MAINWINDOW` (2) for the Mica material backdrop on Windows 11. The class uses a stock white background brush for the delegated light-mode erase path, while the dark-mode `WM_ERASEBKGND` handler paints the custom background explicitly. This prevents a restored light window from exposing a black, uninitialized DWM surface.

**2. Client controls.** The ComboBox, Browse/Run buttons, and status label are switched to comctl32's dark visual style via `SetWindowTheme(hwnd, "DarkModeExplorer", NULL)`, and `WM_CTLCOLOREDIT`/`WM_CTLCOLORBTN`/`WM_CTLCOLORSTATIC`/`WM_CTLCOLORLISTBOX` handlers in `window_dispatch.inc` paint the exact background/text colors (four cached `CreateSolidBrush` handles — background, surface, hover, input — built once in `EnsureThemeBrushes` and reused for the process lifetime) so the result matches Notepad's palette instead of a jarring light-control-on-dark-window mismatch.

**3. Native menu preference.** Two undocumented, ordinal-only `uxtheme.dll` exports let USER32 use the system's immersive popup-menu palette: ordinal 135 (`AllowDarkModeForApp` before 1903 / `SetPreferredAppMode` on 1903+, called once with value `1`) and ordinal 136 (`FlushMenuThemes`). Value `1` is important: it means TRUE/AllowDark across the two ABIs, whereas value `2` means ForceDark on newer Windows and incorrectly makes the dropdown dark in light mode. Both exports are resolved dynamically and gated behind `RtlGetVersion` build ≥ 17763, so older Windows versions never call them. `AllowDarkModeForWindow` receives the current `AppsUseLightTheme` state for the CMDT window.

The classic top-level menu strip is a USER32 non-client surface and is not reliably recolored by those APIs. The source retains an isolated `ApplyDarkMenuBar` owner-draw experiment and matching handlers for reference, but no production path calls it: activating it caused re-entrant menu messages and an unresponsive x64 window on a current Windows build. Consequently, controls and the native popup follow the system theme, while the top strip may retain the classic system rendering. This is an explicit stability trade-off, not a manifest setting; the manifest controls visual styles, DPI awareness, and execution level, but cannot force a dark classic `HMENU`.

The manifest specifies `requestedExecutionLevel=asInvoker`. CMDT does not rely on the manifest for elevation — instead, it programmatically checks `IsUserAnAdmin()` at startup and re-launches itself via `ShellExecuteExW("runas")` if not elevated. This approach allows the same binary to be invoked silently from already-elevated contexts (scripts, scheduled tasks, elevated terminals) without triggering a redundant UAC prompt, while still self-elevating when launched from a standard user session.

---

## MRU (Most Recently Used) History

History is **opt-in and off by default** ([issue #2](https://github.com/wesmar/CmdT/issues/2)) — some users don't want *any* trace of run commands left on disk, and CMDT now respects that out of the box. There is no separate "enabled" flag stored anywhere: **the existence of `HKEY_CURRENT_USER\Software\cmdt` *is* the on/off state.**

- **File → Enable History** (GUI menu checkbox) reflects whether the key exists at startup.
- **Checking it** flips the in-memory flag only — the key is created lazily by the next executed command, same mechanism as before.
- **Unchecking it** flips the flag and calls `RegDeleteTreeW` immediately, wiping the key and every stored value on the spot.
- **`cmdt -history-clear`** does the same wipe from the CLI, for anyone who never opens the GUI.
- While disabled, `LoadMRU` is never called and `SaveMRU` exits immediately — zero registry reads or writes, not even a failed `RegOpenKeyExW` probe.

When enabled, values are stored as named entries `0` through `4`, where `0` is the most recent command. On startup, `LoadMRU` reads these values and populates the ComboBox dropdown. After each successful execution, `SaveMRU` shifts existing entries down (0→1, 1→2, ..., 3→4), deletes the oldest entry, and writes the new command at position 0. Duplicate detection is implicit — the shift operation naturally pushes older duplicates off the end of the list.

The MRU list, when enabled, persists across sessions, reboots, and updates. It remains the only state CMDT writes to disk (via the registry), and now only when the user has explicitly asked for it.

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

The build script auto-detects the newest installed Visual Studio C++ toolchain and Windows SDK (`Get-LatestVCToolsPath`/`Get-LatestWinSDK` — see the 2026 changelog entry below for why this isn't a hardcoded path), assembles all source modules for both architectures as `/subsystem:console`, compiles the resource file (`cmdt.rc`) with the manifest, and links against system import libraries only:

`kernel32.lib`, `user32.lib`, `advapi32.lib`, `shell32.lib`, `comdlg32.lib`, `ole32.lib`, `gdi32.lib`, `shlwapi.lib`, `userenv.lib`, `dwmapi.lib`, `uxtheme.lib`, `OleAut32.lib`

No CRT library is linked. The entry point is `mainCRTStartup` (x64) / `start` (x86) — these are raw assembly procedures, not CRT initialization stubs.

Output binaries are placed in the `bin\` directory (gitignored — rebuilt fresh by every `.\build.ps1` run; the tracked release copies live under `data\` for packaging, kept in sync manually).

---

## Project Structure

```
cmdt/
├── x64/                          # AMD64 assembly sources
│   ├── main.asm                  # Entry point, CLI/GUI dispatch, privilege table, console-hide-on-GUI-launch
│   ├── cli.asm                   # CLI mode and file-run dispatch, -outfile/-errfile relay protocol
│   ├── help.asm                  # Usage banner and help-switch recognition
│   ├── relay.asm                 # Non-admin AND admin output relay (temp-file bridge, exit-code passthrough)
│   ├── token.asm                 # Token acquisition, SYSTEM impersonation, service control
│   ├── process.asm               # CreateProcessWithTokenW wrapper
│   ├── install.asm               # Context menu registration and Sticky Keys IFEO hook
│   ├── window.asm                # GUI entry: WNDCLASSW registration, `include`s the six window_*.inc below
│   ├── tray.asm                  # Shift+Minimize tray icon, restore/exit menu, Explorer-restart recovery
│   ├── window_lifecycle.inc      # InitDpiApis, window/control creation, DPI scaling math
│   ├── window_theme.inc          # Dark mode: DWM chrome, controls, native UxTheme menu preference
│   ├── window_commands.inc       # RunCommand — reads the ComboBox, dispatches to RunAsTrustedInstaller
│   ├── window_dispatch.inc       # WndProc message dispatch table + WM_CREATE/WM_COMMAND handlers
│   ├── window_events.inc         # WndProc continued: WM_SIZE, input, WM_MEASUREITEM/WM_DRAWITEM, teardown
│   ├── window_mru.inc            # LoadMRU/SaveMRU — registry-backed Most Recently Used list
│   ├── strutil.asm               # Wide-character string helpers (wcscpy_p, wcscat_p, …)
│   ├── consts.inc                # Windows API constants, control IDs, message codes
│   └── globals.inc               # External symbol declarations shared across modules
├── x86/                          # IA-32 assembly sources (parallel structure)
│   └── …
├── bin/                          # Compiled binaries (gitignored, rebuilt by build.ps1)
├── data/                         # Tracked release copies used for packaging (pack-data.sh)
│   ├── cmdt_x64.exe
│   ├── cmdt_x86.exe
│   └── test.bat
├── cmdt.rc                       # Version info resource
├── cmdt.manifest                 # Application manifest (DPI, visual styles, execution level)
├── build.ps1                     # Build script (auto-detects VS/SDK, assembles + links both architectures)
└── README.md                     # This file (documentation)
```

Every source file in `x64/` has a corresponding counterpart in `x86/`. The x86 versions use `.586` + `flat/stdcall` MASM syntax with `invoke` macros; the x64 versions use raw `proc frame` with explicit SEH prologue/epilogue annotations (`.pushreg`, `.allocstack`, `.setframe`, `.endprolog`). Both targets share the same `.rc` and `.manifest` files.

Both source trees were originally monolithic — a single ~90 KB `main.asm` on each side. A first pass split the dispatcher, help, relay, install, and string helpers into their own files. `window.asm` itself remained a second, GUI-specific tapeworm (the entire `WndProc`, DPI handling, dark-mode theming, MRU list, and command dispatch all in one file) until a second pass broke it into the six `window_*.inc` files listed above, `include`d back into `window.asm` at assembly time. They deliberately remain one MASM translation unit, preserving internal labels, stack-frame relationships, and call boundaries without adding runtime indirection. The linked binaries remain within the same size limits; the refactor does not claim byte-for-byte object identity because include ordering may change procedure layout. The split is identical on both architectures, so any reader who learns one tree can navigate the other without re-orientation.

---

## String Operations — No CRT

CMDT implements all necessary string operations as hand-written wide-character (UTF-16LE) assembly routines in `strutil.asm`. There is no dependency on `msvcrt.dll`, `ucrtbase.dll`, or any C runtime:

| Function | Purpose |
|---|---|
| `wcscpy_p` | Wide string copy |
| `wcscat_p` | Wide string concatenation |
| `wcscmp_ci` | Case-insensitive wide string comparison |
| `wcscmp_token` | Token-prefix comparison (match up to first space) |
| `wcslen_p` | Wide string length |
| `skip_spaces` | Skip leading whitespace in command parsing |
| `DecryptWideStr` | In-place string decryption for obfuscated constants |

All routines are declared `PUBLIC` in `strutil.asm` and referenced via `EXTRN` in the modules that use them. Each is a tight loop operating on 16-bit words, with inline ASCII case folding for the comparison functions (uppercase A–Z folded to lowercase by adding 32 to the code point).

---

## Verification

After launching a command prompt as TrustedInstaller:

```
cmdt_x64.exe -cli cmd
```

Verify the security context:

```
C:\Windows\System32> whoami
nt authority\system

C:\Windows\System32> whoami /groups | findstr /i trustedinstaller
C:\Windows\System32> whoami /priv
```

The groups output should contain the TrustedInstaller service SID. `whoami /priv` shows the privileges actually present and enabled on that machine; it is not expected to invent every entry from CMDT's request table.

---

## Security Considerations

**TrustedInstaller is a highly privileged Windows service identity, not a formal privilege level above SYSTEM.** Its service SID is granted access to many Windows resources that ordinary administrators and a plain SYSTEM token may not satisfy directly. Combined with the service token's privileges, this can permit operations such as:

- Modifying protected OS components whose ACL grants access to TrustedInstaller
- Writing TrustedInstaller-protected registry locations
- Exercising sensitive token privileges that are present and enabled
- Starting processes in a security context unsuitable for routine work

Use CMDT with extreme caution. Modifying protected Windows components or registry state can render the operating system unbootable even though the process remains subject to ACLs, integrity policy, protected-process rules, and kernel security boundaries.

CMDT requires Administrator privileges to run. It does not bypass UAC — the user must explicitly elevate the process before CMDT can acquire the TrustedInstaller token.

---

## Changelog

<details open>
<summary><strong>17.07.2026 — converted to console subsystem, system-aware dark mode, window.asm split into focused modules</strong></summary>

### Architecture change: `/subsystem:windows` → `/subsystem:console`

Both `cmdt_x64.exe` and `cmdt_x86.exe` now link as `/subsystem:console` instead of `/subsystem:windows`. The PE subsystem field is read by the parent shell's `CreateProcess` call before any of the child's own code runs, so it can never be influenced at runtime — a GUI-subsystem process is simply never waited on by `cmd.exe`. That is the root cause of the glued-output/blank-double-prompt symptom some users saw after `-cli` commands: `cmd.exe` had already redrawn its prompt before the child finished writing, because it never blocked on the child at all. Every prior workaround (`NudgeConsolePrompt` posting a synthetic `VK_RETURN` into the console input queue to force a prompt redraw) only patched the symptom. Console subsystem fixes the actual cause: `cmd.exe` now waits synchronously and prints its next prompt only after CMDT has genuinely finished, with output in the correct order every time. `NudgeConsolePrompt` and the double-UAC-prompt-avoidance dance it was compensating for are gone from the relay paths entirely.

The trade-off: a console-subsystem process gets a console window attached by the OS before its own code runs, which would otherwise cause a visible flash on every argument-less GUI launch (Explorer double-click, shortcut, Run dialog). Fixed without spawning a second process — see the next entry.

### New: in-process console hiding for GUI launches (replaces the old detach-and-relaunch trick)

At startup, if no arguments were given, CMDT calls `GetConsoleProcessList` to determine whether the console it was just attached to belongs **exclusively** to this process or is **shared** with an interactive parent (the user typed `cmdt` inside an already-open `cmd.exe`). In the exclusive case — the common Explorer/shortcut/Run-dialog launch — it calls `GetConsoleWindow` + `ShowWindow(SW_HIDE)` immediately and falls straight into the GUI message loop in the *same process*. In the shared case it does nothing further; hiding a console window that's actually the user's whole terminal would be a bug, not a feature, and opening the GUI on top of an already-visible shell is the correct, expected behavior there. This replaces an earlier approach (`LaunchDetachedGui`, since removed as dead code) that spawned a whole `DETACHED_PROCESS` copy of the executable and exited the parent — functionally similar but with the overhead of a second `CreateProcessW`/image-load round-trip; the in-process hide is two Win32 calls and measurably faster to hide the window. Verified via `EnumWindows` + `IsWindowVisible` against the console's own `ConsoleWindowClass` window in both the exclusive and shared scenarios, and via `whoami`/exit-code checks that `-cli` behavior is unaffected by the subsystem change.

### Fix: `-cli` stdout/stderr now relay through **separate** temp files, and the real exit code propagates

Previously the relay captured stdout and stderr into a single shared temp file, so `2>err.txt` without also redirecting stdout (or the reverse) interleaved both streams into whichever target was specified. The relay now creates two temp files (`-outfile`/`-errfile`) and streams each back to the caller's own `STD_OUTPUT_HANDLE`/`STD_ERROR_HANDLE` independently. Separately, the relay previously always exited `0` regardless of what the spawned command actually returned — `GetExitCodeProcess` is now captured after `WaitForSingleObject` and the relay process exits with that same code, so `%ERRORLEVEL%`/`$LASTEXITCODE` checks in a caller's script correctly see a `-cli` command's real failure instead of always seeing success.

### Dark mode now covers the client controls and follows the system preference

Previously "dark mode" meant only `DwmSetWindowAttribute(DWMWA_USE_IMMERSIVE_DARK_MODE)` on the title bar and a Mica backdrop. `SetWindowTheme` plus `WM_CTLCOLOR*` handlers now cover the ComboBox, buttons, status label, and client background. Runtime-resolved UxTheme ordinals 135/136 (gated behind an `RtlGetVersion` check for Windows 10 1809+) allow the native popup menu to follow the system preference; ordinal 135 uses value `1` (AllowDark), never `2` (ForceDark), so light mode stays light. An owner-drawn top menu experiment was retained in isolated source but disabled after it caused re-entrant USER32 messages and an x64 hang. The classic top strip can therefore remain system-rendered.

### Refactor: `window.asm` split into six focused `.inc` files

`window.asm` was the last remaining tapeworm file — the entry point/dispatcher/relay/install code had already been split out in earlier releases, but the entire GUI (`WndProc`, DPI handling, theming, MRU, command dispatch) stayed in one file. It's now `window.asm` `include`-ing `window_lifecycle.inc`, `window_theme.inc`, `window_commands.inc`, `window_dispatch.inc`, `window_events.inc`, and `window_mru.inc` — see [Project Structure](#project-structure) for what each one owns. MASM still assembles them as one translation unit, so there are no new runtime call boundaries or global-state handoffs. Behavior and size limits are unchanged; byte-for-byte object identity is not asserted because include ordering can alter layout.

### `build.ps1`: auto-detects the newest Visual Studio + Windows SDK instead of hardcoded paths

Extended the existing `Get-LatestVCToolsPath` (see the 13.07.2026 entry below) with a matching `Get-LatestWinSDK` that picks the highest-versioned directory under `Windows Kits\10\Include` instead of a hardcoded SDK version string — the previous hardcoded version stopped existing on disk after a Windows SDK update, breaking the build with no indication why. `uxtheme.lib` added to the link libraries for the dark-mode `SetWindowTheme` call (a documented, statically-linkable export, unlike the undocumented ordinal-only functions above).

</details>

<details>
<summary><strong>13.07.2026 — PerMonitorV2 DPI scaling actually implemented, build script fixed for newer VS releases</strong></summary>

### Fix: GUI didn't scale on PerMonitorV2-aware displays

The manifest has declared `dpiAwareness=PerMonitorV2` since early builds, but that only stops Windows from doing automatic bitmap stretching — it does **not** scale anything for you. Every control coordinate in `window.asm` (`CreateWindowExW` calls in `WM_CREATE`, the main window's own creation size, and the `WM_SIZE` repositioning math) was a raw pixel constant authored for 96 DPI (100%). On a scaled display (125%/150%/etc.) the window and every control inside it rendered at their literal 96-DPI pixel size — tiny, with text and buttons cramped or clipped.

Fixed on both architectures:

- `GetDpiForSystem()` / `GetDpiForWindow(hWnd)` (Windows 10 1607+) are queried at window creation, at `WM_CREATE`, and at `WM_SIZE`. Every design-time coordinate is scaled by `value * dpi / 96` before use.
- The main window's total size is no longer a guessed pixel constant either. A fixed total-window guess has to assume how tall the caption bar, menu bar, and borders are — that overhead varies by Windows theme and font scale across machines, and a wrong guess either clips the bottom row of controls or leaves dead space. The window size is now computed from the **desired client area** via `AdjustWindowRectExForDpi`, which asks Windows for the exact non-client overhead at the current DPI instead of assuming one.

Verified via terminal launches at several DPI scales, both elevated and non-elevated, on real high-DPI hardware where the clipping was originally reported.

### Fix: `build.ps1` couldn't find the VC++ toolchain on newer Visual Studio releases

`vswhere -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64` started returning nothing on a newer VS release even though the C++ toolchain was installed — the component id vswhere expects there no longer matches what the installer registers. Compounding it, VS also renamed the `VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt` marker file the script used to read the active toolchain version from, replacing it with versioned filenames.

`Get-LatestVCToolsPath` in `build.ps1` no longer depends on either. It now lists every VS-Installer-registered product newest-first (`vswhere -all -sort`, no `-requires` filter — a bare `-latest` isn't safe by itself either, since other VS-Installer-registered products like SQL Server Management Studio can sort as "latest" and obviously have no C++ tools), and for each one looks directly for `VC\Tools\MSVC\<version>\bin\Hostx64\x64\ml64.exe`, taking the highest-versioned directory from the first installation that has it.

### File version bumped to 1.0.0.2

Both fixes above (DPI scaling, the double-UAC-prompt regression) shipped in this version.

</details>

<details>
<summary><strong>12.07.2026 — history made opt-in, `-cli` output relay fixed for already-elevated shells</strong></summary>

Two independent fixes driven by user-filed issues, both verified end-to-end on x64 and mirrored to x86.

### Fix: command history is now off by default (issue #2)

Previously CMDT always wrote the last 5 commands to `HKCU\Software\cmdt` with no way to disable it. The model changed: **the registry key's existence *is* the on/off state**, nothing else is persisted anywhere.

A new **File → Enable History** checkbox was added to the GUI menu, between "Open with TrustedInstaller" and "About". At `WM_CREATE`, CMDT does a read-only `RegOpenKeyExW` probe against `HKCU\Software\cmdt` to decide the checkbox's initial state, then immediately closes the handle. Checking the box just flips an in-memory flag — the key is still created lazily by the next executed command, same mechanism as before. Unchecking it flips the flag *and* calls `RegDeleteTreeW` immediately, wiping the key and every stored value on the spot. `LoadMRU`/`SaveMRU` both now start with a check against that flag and return immediately if history is off — no registry touch at all in that case, not even a failed open.

A new CLI switch, `cmdt -history-clear`, performs the same `RegDeleteTreeW` wipe for anyone who never opens the GUI (added to the usage banner and About dialog on both architectures).

### Fix: `cmdt -cli <cmd> > out.txt` from an already-elevated shell (issue #1)

Reported symptom: running `cmdt_x64.exe -cli net session > out.txt` from an **admin** `cmd.exe` produced an empty file and two blank prompts; the same command from a non-admin shell (which goes through UAC) worked fine.

Root cause: `RunAsTrustedInstaller` always launches the target through `CreateProcessWithTokenW`, which — unlike plain `CreateProcess` — does not reliably hand the spawned child usable inherited std handles. The already-admin dispatch path called it directly with `STARTF_USESTDHANDLES` and inherited/duplicated handles; the non-admin path already avoided this by routing through the temp-file relay (`relay.asm`). The fix: a new `AdminRelayLaunch` routine captures output entirely **in-process** — it opens a local temp file, sets `g_relayHandle`, and calls `RunAsTrustedInstaller` directly (token duplication only), then streams the temp file to the caller's own `STD_OUTPUT_HANDLE`. Critically, this path never calls `ShellExecuteExW("runas")`: an early version of this fix reused the non-admin relay (which *does* shell out via `"runas"`) for the already-admin case too, and re-elevating an already-elevated process that way can still pop a second UAC consent dialog — that regression was caught before release and replaced with the in-process approach described here. `admin_dispatch` (x64) / `uac_already_admin` (x86) routes `-cli` through `AdminRelayLaunch` unconditionally, falling through to direct dispatch only for `-new`, interactive shells (`cmd`/`powershell`/`pwsh` with no further arguments), and file-run mode. A guard checks `argv[2] == "-outfile"` and declines in that case too, so the non-admin relay's own elevated child (which re-enters this same dispatch code as an admin process) doesn't try to relay itself again — it falls through to `cli.asm`'s existing `-outfile` handling instead, which sets `g_relayHandle` and lets `RunAsTrustedInstaller` write straight to the temp file.

Verified: `net session`, `net user`, and `whoami` through `-cli`, with `>`, `>>`, and `| findstr`, from both an admin `cmd.exe` (no UAC prompt) and a non-admin shell (real UAC prompt). `whoami` specifically returns `nt authority\system` — confirmation that the relayed child genuinely carries the TrustedInstaller token, not merely plain Administrator.

A secondary bug surfaced during this testing: relay mode (`process.asm`, Mode 3) set `hStdInput = NULL`. Combined with `CREATE_NO_WINDOW`, that's enough to make some console-aware targets refuse to start. Fixed on both architectures by opening an inheritable handle to the `NUL` device instead and closing it after the child exits (or on failure, mirroring the existing duplicate-handle cleanup).

**Known limitation, not fixed:** when the `-cli` command *is itself* `cmd.exe` with its own embedded I/O redirection (e.g. a `.lnk` shortcut targeting `cmd.exe /c foo > file`, or a literal `cmdt -cli cmd /c foo > file`), `cmd.exe` refuses with `"Input redirection is not supported, exiting the process immediately."` Tried giving it a real-but-hidden console (`CREATE_NEW_CONSOLE` + `STARTF_USESHOWWINDOW`/`SW_HIDE`) instead of `CREATE_NO_WINDOW`; it made no difference, so that change was reverted in favor of the known-good baseline. This is `cmd.exe` itself declining to redirect without what it considers a proper console — pre-existing in the relay path before this release, not a regression. Plain commands, `.lnk` targets without embedded redirection, and `cmd /c <command>` without a trailing redirect all work correctly. Tracked for a future pass.

### File version bumped to 1.0.0.1

Every previous release shipped `FileVersion`/`ProductVersion` hardcoded at `1.0.0.0`, making it impossible to tell builds apart from Explorer's file properties ([raised on the forum](https://forums.mydigitallife.net/threads/cmdt-ultra-lightweight-run-as-trustedinstaller-pure-x64-x86-assembly.90096/)). Going forward, the version resource in `cmdt.rc` is bumped on every functional change instead of staying frozen.

</details>

<details>
<summary><strong>17.05.2026 — interactive-shell guard, x86 console handling synced with x64</strong></summary>

Two follow-up fixes on top of the 16.05.2026 release after non-admin testing in real shells.

### Fix: `cmdt -cli cmd` / `powershell` / `pwsh` from a non-admin shell

The temp-file relay introduced on 16.05.2026 is designed for *non-interactive* commands — it spawns the elevated child with `CREATE_NO_WINDOW` and captures stdout/stderr into a temp file. Pointing it at an interactive shell (`cmd`, `powershell`, `pwsh`) was incompatible: the child had no console to draw a prompt on, no stdin to read from, and the parent's `WaitForSingleObject` blocked the caller indefinitely. The user saw a flash and no shell.

`NonAdminRelayLaunch` now declines the relay when argv looks like exactly `cmdt -cli <shell>` (argc == 3 and `argv[2]` matches one of `cmd` / `cmd.exe` / `powershell` / `powershell.exe` / `pwsh` / `pwsh.exe`, case-insensitive). On decline, the caller falls through to plain UAC self-elevation (`ShellExecuteExW("runas")`), which spawns a real, visible console for the TrustedInstaller shell. The original cmd's prompt returns immediately — important for batch scripts that invoke `cmdt -cli cmd /c …` patterns and need control back.

Anything with extra arguments (`cmd /c dir`, `powershell -Command Get-Process`) still goes through the relay so its output is streamed back to the caller. The decline list lives next to its own pointer table in `relay.asm` so adding more interactive shells is one line.

### Sync: x86 console handling brought in line with x64

The 16.05.2026 fix for `AttachConsole` clobbering inherited std handles and the `NudgeConsolePrompt` helper that posts a fake VK_RETURN to redraw cmd's prompt both landed on x64 only. On x86, `cmdt_x86.exe -cli net session` with no redirect produced no output (`GetStdHandle(STD_OUTPUT_HANDLE)` returned `NULL` because `cmdt_x86.exe` is `/subsystem:windows`, so cmd.exe doesn't wire up the console handle for it), and the prompt-redraw nudge only ran for the help-banner path, not the relay path.

x86 now mirrors x64:

- `start` (in `main.asm`) does the same early `AttachConsole(ATTACH_PARENT_PROCESS)` guarded by `GetStdHandle` + `GetFileType` — attach when stdout is `NULL`/`INVALID` or `FILE_TYPE_CHAR`, leave alone for `FILE_TYPE_DISK` / `FILE_TYPE_PIPE` so redirects and pipes keep their inherited handles.
- `NudgeConsolePrompt` extracted into `help.asm` as a proper proc (was previously inlined inside `ShowUsageAndExit`), and `relay.asm` calls it before `DeleteFileW` so the cursor doesn't sit idle one line above the new prompt after relay output finishes.

### Cases now covered end-to-end on both architectures

| Command | Behavior |
|---|---|
| `cmdt -cli cmd` (non-admin) | Declines relay → plain UAC → new TI cmd window |
| `cmdt -cli cmd /c dir` (non-admin) | Relay fires, output streams to caller console |
| `cmdt -cli net session` (non-admin) | Relay fires, output streams to caller console |
| `cmdt -cli net session > out.txt` (non-admin) | Relay fires, writes to `out.txt` |
| `cmdt -cli net session \| findstr X` (non-admin) | Relay fires, writes to pipe |
| `cmdt -cli cmd -new` (non-admin) | `-new` declines relay (preexisting) → plain UAC → new console |

The relay declines only when it has no useful captured-output story to tell. Every redirect / pipe case still gets the relay, so `>`, `>>`, and `|` keep delivering TrustedInstaller output to the caller in every shell.

</details>

<details>
<summary><strong>16.05.2026 — modular split for both architectures, MRU policy change, redirect/help fixes</strong></summary>

This release covers a multi-day overhaul. Both source trees are now modular, the GUI no longer pre-fills the last command at startup, and the CLI redirect path works in every elevation / shell combination.

### Source refactoring — monolith split into focused modules (both x86 and x64)

Each architecture previously had one large monolithic `main.asm` (~90 KB) that bundled the entry point, dispatcher, helpers, GUI, context-menu installer, and Sticky-Keys hook. The split happened on x64 first (15.05.2026); x86 was migrated to the same shape on 16.05.2026 so both trees are now mirror images of each other:

- `cli.asm` — CLI mode and file-run dispatch; owns the `-outfile` relay protocol and (x86 only) the WOW64 regedit fallback (`FixRegeditPath`)
- `help.asm` — usage banner text and help-switch recognition (`HelpCheckAndExit` / `ShowUsage` on x64; `IsHelpSwitch` / `ShowUsageAndExit` on x86)
- `relay.asm` — non-admin output relay (`NonAdminRelayLaunch`)
- `install.asm` — context menu registration, Sticky-Keys IFEO hook, Defender exclusion management (WMI on x64, PowerShell on x86)
- `strutil.asm` — wide-character string helpers (`wcscpy_p`, `wcscat_p`, `wcscmp_ci`, `wcscmp_token`, `wcslen_p`, `skip_spaces`; plus `DecryptWideStr` on x64)
- `process.asm` — `CreateProcessWithTokenW` wrapper, environment block setup
- `token.asm`, `window.asm` — unchanged in responsibility, trimmed by extracting helpers
- `main.asm` — entry point, CLI/GUI dispatch, privilege table, global data, `mode_gui` as a standalone proc

To make the x86 split possible, several formerly-local variables of `start` (`pArgv`, `argc`, `argv1`, `sa`) were promoted to globals (`g_argv`, `g_argc`, `g_argv1`, `g_sa`) so the dispatch labels can be reached cross-module while still sharing the parsed command-line state. `mode_gui` was lifted out of `start` into its own `proc` for the same reason — MASM does not let cross-module jumps land inside another procedure.

Compiled binary sizes are unchanged. Both `cmdt_x64.exe` and `cmdt_x86.exe` stay under their previous limits.

### Fix: `cmdt -cli <command> >> out.txt` from a non-admin shell

Running `cmdt -cli net session >> out.txt` (or any `>`, `>>`, `|` redirect) from a non-admin shell previously produced no output — the elevated child process was spawned in a new process tree by UAC and its stdout handle was disconnected from the caller's redirect target.

The fix introduces a **temp-file relay** in `relay.asm`:

1. Non-admin parent creates a temp file via `GetTempFileNameW`.
2. Inserts an internal `-outfile <path>` token into the argument string and launches the elevated child via `ShellExecuteExW("runas")` with `SEE_MASK_NOCLOSEPROCESS`, then waits with `WaitForSingleObject` on the returned `hProcess`.
3. Elevated child (`cli.asm`, `mode_cli_setup`) opens the temp file with an inheritable `GENERIC_WRITE` handle and uses it as `hStdOutput`/`hStdError` for the spawned process (`CREATE_NO_WINDOW`).
4. After the child exits, the non-admin parent opens the temp file, streams it to its own `STD_OUTPUT_HANDLE` (which cmd.exe wired up before launch — so `>file`, `>>file`, and `|pipe` all work transparently), then deletes the temp file and exits.

Two bugs were also fixed while implementing this path:

- **`SHELLEXECUTEINFOW.hProcess` offset on x64** — the relay was reading offset `+96` (the `hIcon`/`hMonitor` union) instead of `+104` (the real `hProcess`). That meant `WaitForSingleObject` was skipped and the parent raced the child for the temp file. The relay-decline fallback masked this on x86, where the SHELLEXECUTEINFOW layout is smaller and the offset wasn't wrong.
- **`AttachConsole` clobbering inherited std handles** — when cmd.exe passes a `>file` redirect handle to a GUI-subsystem child, `AttachConsole(ATTACH_PARENT_PROCESS)` overwrites that handle with `CONOUT$`, silently dropping the redirect target. We now query `GetStdHandle(STD_OUTPUT_HANDLE)` first and only call `AttachConsole` when stdout is `NULL` or `INVALID_HANDLE_VALUE`.

The relay is skipped when `-new` is passed (a detached console has no output to capture) and falls back to plain UAC self-elevation if temp-file creation fails.

### Fix: help switches work without elevation and with redirection

`cmdt -help`, `cmdt -h`, `cmdt --help`, `cmdt -?`, `cmdt /?`, `cmdt /h`, and `cmdt /help` previously triggered UAC self-elevation before printing anything — meaning the elevated process had no console to write to, and `cmdt -help > out.txt` produced an empty file.

The help check now runs in the entry point **before** `IsUserAnAdmin()`. The help module scans `argv[1]` for all seven recognized spellings; if a match is found, it prints usage and exits without ever calling UAC.

The usage printer selects the output API via `GetFileType(STD_OUTPUT_HANDLE)`: `WriteConsoleW` for a real console (native UTF-16), `WriteFile` for a file or pipe (raw UTF-16 LE bytes). Both interactive display and `cmdt -help > out.txt` now produce correct output in all session types. After writing to a real console, a fake VK_RETURN is posted to stdin via `WriteConsoleInputW` so cmd.exe redraws its prompt instead of leaving the cursor idle.

### MRU dropdown — no longer pre-selects the last command at startup

The GUI ComboBox used to call `CB_SETCURSEL, 0` on both `LoadMRU` and `SaveMRU`, which auto-filled the edit field with the most recent command. After the change, both call sites use `CB_SETCURSEL, -1` followed by `SetWindowTextW(g_hwndEdit, g_cmdBuf)` (g_cmdBuf is zero-initialized in BSS, so it's an empty string). The dropdown still contains the full history — clicking the arrow shows it as before — but the edit field starts empty, so the user types fresh instead of accidentally re-running the previous command.

</details>

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
| Current full build (x86) | **<30 KB** | Hybrid mode, context menu, UAC self-elevation, manifest, COM `.lnk` resolution, system-aware dark mode |
| Current full build (x64) | **<40 KB** | Same feature set, 64-bit calling convention overhead, DPI-aware layout |

The growth from 4–6 KB to the current size (under 30 KB on x86, under 40 KB on x64) is almost entirely due to the application manifest (DPI awareness, Common Controls v6, execution level declaration), the context menu registry logic, the Sticky Keys IFEO hook with Defender exclusion management, UAC self-elevation, the wide-character string constants for registry paths and UI text, and the dark-mode implementation (cached brushes, control theming, and dynamically resolved `uxtheme.dll` ordinals). The core token acquisition pipeline — the actual "engine" of CMDT — remains remarkably compact.

---

*Written in 100% bare-metal x86/x64 MASM assembly. No frameworks. No runtimes. No compromises.*
