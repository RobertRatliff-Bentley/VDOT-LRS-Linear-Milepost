VDOT LRS Linear Milepost
Created by Robert Ratliff - Build 2026-11-25
=====================================================================

Draws a VDOT LRS event in MicroStation or OpenRoads. Choose a jurisdiction
and route from VDOT's LRS Route Master service, pick the From and
To measures on the route in the drawing, then create the event along the
route centreline with its LRS data attached.


REQUIREMENTS
------------
* MicroStation 2026 or OpenRoads Designer 2025 (both include the Python
  Manager and the Python runtime the tool uses).
* Internet access to services.arcgis.com (VDOT LRS Route Master service).

PYTHON LIBRARIES
----------------
Apart from the Python standard library and Bentley's MSPy modules (part of
MicroStation / OpenRoads), the tool needs one library: PyQt5 5.15.
MicroStation and OpenRoads ship with it, so normally nothing is added.

The wheels folder holds an offline copy (PyQt5 5.15.10, PyQt5-Qt5 5.15.2,
PyQt5-sip) for a machine whose Bentley Python lacks PyQt5. The installer
checks first; only when PyQt5 is missing does it install that copy into
%LOCALAPPDATA%\VDOT_LRS_Linear_Milepost\lib (no internet, no admin
rights). The tool uses that copy only when the built-in PyQt5 cannot be
loaded, so it never replaces Bentley's. To force the bundled copy, run:
    powershell -ExecutionPolicy Bypass -File Install.ps1 -BundledQt
The bundled PyQt5 is GPL v3 licensed - see THIRD_PARTY_NOTICES.txt.


INSTALL
-------
1. Extract the whole zip to a folder.
2. Close MicroStation / OpenRoads.
3. Double-click Install.cmd.

The installer needs no administrator rights. For the current Windows user it:
* copies VDOT_LRS_Linear_Milepost.py (and the VBA version,
  VDOT_LRS_Linear_Milepost.mvba) to
      %LOCALAPPDATA%\VDOT_LRS_Linear_Milepost
* installs the bundled PyQt5 only if MicroStation's Python lacks it
* adds that script to the Python Manager list of each product it finds
  (one line in the product's Personal.ucf preferences file; the original
  file is kept beside it as Personal.ucf.before-vdot-lrs).

A product that is running, or that this user has never started, is skipped
and reported. Close or start it once, then run Install.cmd again.

Each person who uses the tool runs Install.cmd once under their own login.


RUN
---
Utilities > Python, choose VDOT_LRS_Linear_Milepost, Run.

Or key in (the installer prints the exact path):
    PYTHON LOAD "C:/Users/<you>/AppData/Local/VDOT_LRS_Linear_Milepost/VDOT_LRS_Linear_Milepost.py"

VBA version (same dialog and behaviour, no Python needed):
    VBA LOAD "C:/Users/<you>/AppData/Local/VDOT_LRS_Linear_Milepost/VDOT_LRS_Linear_Milepost.mvba"
    VBA RUN [VDOT_LRS_Linear_Milepost]modLrsMain.VdotLrsEvent

1. Choose Jurisdiction, then Route (type in a box to search). The Route
   list starts with the Interstate, US and VA routes that cross the
   jurisdiction, then its ramps and its own routes in route-number order.
   "Statewide routes" at the top of the Jurisdiction list holds every
   Interstate, US and VA route in Virginia.
   "Opposite direction" loads the other direction of a divided route;
   it is unavailable for roads the LRS holds as a single route.
2. Choose Linear or Point. Move over the view: the route is drawn, the
   event is highlighted. Data point sets From, the next sets To; the
   dialog marks which one the next pick sets. Reset, or the Pick button,
   ends picking. "Full route" and the swap button set or exchange the
   measures.
3. Choose labels and milepost signs if wanted, then Create alignment.
   Labels are placed to the left of the direction of travel and milepost
   signs to the right, clear of the line.
4. "Delete created" removes everything the tool placed in the active
   model - events, start/stop markers, labels and milepost signs - after
   asking. One Undo in MicroStation brings it all back. Labels, markers
   and signs drawn by builds before 2026-11-25 are not marked, so delete
   those by hand (their event lines are found).

The dialog remembers the theme, jurisdiction, coordinate system, label and
milepost choices between runs. Check the coordinate system when you open a
drawing that uses a different one from last time. Click LOG to hide or show
the message list.


UPDATE
------
Extract the new zip and run Install.cmd again. It replaces the script.


UNINSTALL
---------
Close MicroStation / OpenRoads and double-click Uninstall.cmd. It removes
the script and its Python Manager entry. Elements already drawn, and the
"VDOT LRS" item type in your DGN files, are not touched.


NOTES
-----
* The tool creates DGN graphics (line strings, cells, text) on the active
  level. In OpenRoads these are not civil geometry.
* A 2D model gets 2D elements.
* If the model has no geographic coordinate system, choose an EPSG system
  in the dialog (default EPSG:2284, NAD83 Virginia South, US survey feet).
* Save your work before creating events. This build was tested inside
  MicroStation 2026 on a 2D model (loading routes, picking, and creating
  linear, gap-crossing and point events). It has not been run in
  OpenRoads Designer or on a 3D model.
* The VDOT logo shown in the dialog is VDOT's mark; confirm with VDOT
  before distributing the tool outside your organization.


FILES
-----
VDOT_LRS_Linear_Milepost.py   the tool (self-contained, logo included)
VDOT_LRS_Linear_Milepost.mvba the VBA version of the tool
wheels\                    offline PyQt5 wheels (used only if needed)
THIRD_PARTY_NOTICES.txt    PyQt5 / Qt licence notice
Install.cmd / Install.ps1  installer
Uninstall.cmd              uninstaller
README.txt                 this file
