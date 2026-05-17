# Samsung S23 MTP Backup Script (PowerShell)

A PowerShell script to back up files from a Samsung S23 (or similar Android device) to a local PC directory via MTP — no third-party software required.

---

## Features

- **MTP-based transfer** — uses Windows Shell COM directly, no ADB or extra drivers needed
- **Incremental backup** — skips files already backed up, only copies new ones
- **Two-phase file verification** — waits for file to appear, then waits for size to stabilize before marking as complete (prevents partial/corrupt copies)
- **Auto-retry** — retries failed files up to 2 times automatically
- **Optional delete** — after backup, optionally remove originals from the phone
- **Summary report** — shows copied / skipped / failed count at the end

---

## Requirements

- Windows 10 / 11
- PowerShell 5.1 or later
- Samsung S23 connected via USB in **File Transfer (MTP)** mode
- Phone screen **unlocked** during transfer

---

## Setup

**1. Clone or download the script**

```powershell
git clone https://github.com/714625449/mobile_backup.git
