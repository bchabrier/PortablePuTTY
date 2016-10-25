# PortablePuTTY
Portable packaging of PuTTY

**Portable PuTTY** provides a portable version of **PuTTY** (see http://www.putty.org/).

It allows to persist saved sessions and share them across computers.

## Installation

Installation consists in copying `Portable PuTTY.vbs` in a shared folder.

A typical use is to install **Portable PuTTY** on a USB drive or a dropbox folder.

## Usage

Launch `Portable PuTTY.vbs`. 
At first launch, this creates a shortcut named `Portable PuTTY.lnk`. You can then use this shortcut or copy it for instance to your desktop to launch **Portable PuTTY** more conveniently.

## Backup

**Portable PuTTY** keeps saved sessions in `PuTTY.reg`. The previous versions are kept in `PuTTY.bak` and `PuTTY.bak.bak`.

## Tests

The tests suite uses ScriptUnit (see http://xt1.org/scriptunit). Note that it is necessary to run `scriptunit.exe` as administrator once in order to register its libraries correctly.

