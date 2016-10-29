[![Build status](https://ci.appveyor.com/api/projects/status/k04ikl250e64lwcq?svg=true)](https://ci.appveyor.com/project/bchabrier/portableputty)

# Portable PuTTY
Portable packaging of PuTTY

**Portable PuTTY** provides a portable version of **PuTTY** (see http://www.putty.org/).

It allows to persist saved sessions and share them across computers.

## Installation

Installation consists in copying `Portable PuTTY.vbs` in a shared folder.

A typical installation is to copy `Portable PuTTY.vbs`:

- on a USB drive
- in a Dropbox folder
- in Box folder locally mounted through BoxSync or webdav

**Portable PuTTY** will save all the needed information, including **PuTTY** sessions, on this shared folder. This way, all computers accessing this shared folder will share the same **PuTTY** sessions.

## Usage

Launch `Portable PuTTY.vbs`.

At first launch, this will perform three specific actions:

1. create a shortcut named `Portable PuTTY.lnk`. You can then use this shortcut or copy it for instance to your desktop to launch **Portable PuTTY** more conveniently.

2. automatically download Simon Tatham's PuTTY from https://the.earth.li/~sgtatham/putty/latest/x86/putty.exe.
This is done only once, if `putty.exe` does not already exist in the directory where `Portable PuTTY.vbs` is launched. If you want to use another version of `putty.exe`, you can copy it manually to this directory.

3. if any **PuTTY** sessions were existing on the computer, they will be used by **Portable PuTTY** and saved in the installation folder.

## Backup

**Portable PuTTY** keeps saved sessions in `PuTTY.reg`. The previous versions are kept in `PuTTY.bak` and `PuTTY.bak.bak`.

## Tests

The tests suite uses ScriptUnit (see http://xt1.org/scriptunit). Note that it is necessary to run `scriptunit.exe` as administrator once in order to register its libraries correctly.

They are executed on Appveyor. The status is reflected by the badge at the top of this Readme.


