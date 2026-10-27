# ClickHere Machine Configurator

<p align="center">
  <img src=".github/logo.png" alt="clickhere-machine-configurator logo" width="128" />
</p>

A small Windows utility that configures a machine for the ClickHere
setup automatically when it runs.

## What it does

On load it works through a series of configuration steps and logs
progress to a text box:

- Looks up the machine's IP address and derives a computer name from it,
  then renames the computer with `SetComputerName`.
- Applies registry settings, including the desktop background, the
  Winlogon shell (`Explorer.exe`) and Screen Shield settings.
- Runs helper processes, for example to set up a share.
- Reboots the machine when finished.

A **Cancel** button stops the run and restores Explorer as the shell.

## Tech

- Visual Basic .NET (Windows Forms)
- Win32 API calls (`kernel32`, `user32`, `advapi32`, `powrprof`)
- `Microsoft.Win32.Registry` and `System.DirectoryServices`

## Building

Open the `.sln` file in Visual Studio and build it. The project was
upgraded to the Visual Studio 2005 format (.NET Framework 2.0) and targets
Windows only.

The solution file is `chickhere.sln`.

## License

Licensed under the Apache License 2.0. See [LICENSE](LICENSE).
