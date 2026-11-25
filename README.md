# VDOT LRS Linear Milepost

A MicroStation 2026 / OpenRoads Designer 2025 tool that draws VDOT LRS events.

It reads VDOT's public **LRS Route Master** REST service. You choose a jurisdiction and a route from searchable drop-downs, then pick the From and To measures on the route in the drawing. A **Draw it?** bar then lets you create the event along the route centreline. The event gets LRS metadata (an Item Type), labels, start/stop markers and milepost signs. **Delete created** removes everything the tool placed, after asking for confirmation inside the dialog.

There are two editions with the same dialog and behaviour:

| Edition | File | Runs with |
|---|---|---|
| Python | `VDOT_LRS_Linear_Milepost.py` | `PYTHON LOAD "<path>"` or Utilities > Python |
| VBA | `VDOT_LRS_Linear_Milepost.mvba` | `VBA LOAD "<path>"`, then `VBA RUN [VDOT_LRS_Linear_Milepost]modLrsMain.VdotLrsEvent` |

## Install

1. Download `VDOT_LRS_Linear_Milepost_<date>.zip` from **[Releases](../../releases)** and extract the whole zip.
2. Close MicroStation / OpenRoads, then double-click `Install.cmd`. No administrator rights are needed.

The installer copies both editions to `%LOCALAPPDATA%\VDOT_LRS_Linear_Milepost` and adds the Python edition to the Python Manager list. To remove the tool, run `Uninstall.cmd`. See [`package/VDOT_LRS_Linear_Milepost/README.txt`](package/VDOT_LRS_Linear_Milepost/README.txt) for details.

## Python libraries

The Python edition needs the standard library, Bentley's MSPy modules (part of the product) and **PyQt5 5.15**.

MicroStation and OpenRoads ship with PyQt5, so normally the installer adds nothing. The release zip also carries offline wheels in `wheels\`: PyQt5 5.15.10, PyQt5-Qt5 5.15.2 and PyQt5-sip. If the product's Python lacks PyQt5, the installer puts these wheels in `lib\` beside the script. The script falls back to that copy only when the built-in PyQt5 cannot be imported, so Bentley's PyQt5 is never replaced. The bundled PyQt5 is GPL v3; see `THIRD_PARTY_NOTICES.txt`.

The wheels are not committed to git. They are attached to the release zip only. To rebuild them, use Bentley's Python:

```
python -m pip download PyQt5==5.15.10 PyQt5-Qt5==5.15.2 --only-binary=:all: --platform win_amd64 --python-version 3.12 --implementation cp -d package\VDOT_LRS_Linear_Milepost\wheels
```

## Repository layout

```
package/VDOT_LRS_Linear_Milepost/   what ships: tool (.py, .mvba), installer, README, notices
vba/                                VBA sources and build_mvba.py (builds the .mvba from MicroStation over COM)
tests/lrs_mstn_selftest.py          in-MicroStation self-test (PYTHON LOAD it on a scratch DGN)
```

## Notes

- Tested in MicroStation 2026 (loading routes, picking, creating and deleting events). It has not been run in OpenRoads Designer.
- The VDOT logo is VDOT's mark. Confirm with VDOT before distributing the tool outside your organization.

Created by Robert Ratliff.
