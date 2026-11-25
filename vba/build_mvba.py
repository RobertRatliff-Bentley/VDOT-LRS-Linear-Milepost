"""Builds the VDOT_LRS_Linear_Milepost VBA project from the sources in this folder.

Needs a running MicroStation (it uses MicroStation's own VBA editor through COM):

    python build_mvba.py <work.mvba> --shell <any existing .mvba> [--with-selftest]
    python build_mvba.py --compact <work.mvba> <VDOT_LRS_Linear_Milepost.mvba>

MicroStation's VBA editor cannot create a project through automation, so the first form
copies an existing .mvba (--shell), loads it, replaces its contents and saves it. Unload
the project (VBA UNLOAD) before running the second form, which writes a compact copy
without the shell project's left-over form storages and compiled caches.

The form's drop-downs, buttons, log frame and logo are created here as design-time
controls; their layout and colours are set by the form's own code at start-up.
"""
import os
import shutil
import sys
import time

import pythoncom
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = "VDOT_LRS_Linear_Milepost"
MSTN_TYPELIB = "{CF9F97BF-39F2-4B8E-835C-8BE9E99DAF5B}"

STANDARD, CLASS, FORM = 1, 2, 3
MODULES = [
    ("modLrsMain", STANDARD, "modLrsMain.bas"),
    ("modLrsJson", STANDARD, "modLrsJson.bas"),
    ("modLrsDraw", STANDARD, "modLrsDraw.bas"),
    ("clsLrsGeometry", CLASS, "clsLrsGeometry.cls"),
    ("clsLrsRoute", CLASS, "clsLrsRoute.cls"),
    ("clsLrsService", CLASS, "clsLrsService.cls"),
    ("clsLrsPickTool", CLASS, "clsLrsPickTool.cls"),
    ("frmLrsEvent", FORM, "frmLrsEvent.frm.txt"),
]
SELFTEST = ("modLrsSelfTest", STANDARD, "modLrsSelfTest.bas")

FORM_CONTROLS = [
    ("Forms.Image.1", "imgLogo"),
    ("Forms.ComboBox.1", "cboTheme"),
    ("Forms.ComboBox.1", "cboJurisdiction"),
    ("Forms.ComboBox.1", "cboRoute"),
    ("Forms.ComboBox.1", "cboFrom"),
    ("Forms.ComboBox.1", "cboTo"),
    ("Forms.ComboBox.1", "cboCoords"),
    ("Forms.ComboBox.1", "cboLabels"),
    ("Forms.ComboBox.1", "cboMileposts"),
    ("Forms.CommandButton.1", "cmdPick"),
    ("Forms.CommandButton.1", "cmdZoom"),
    ("Forms.Frame.1", "fraLog"),
    ("Forms.CommandButton.1", "cmdOpposite"),
    ("Forms.CommandButton.1", "cmdLinear"),
    ("Forms.CommandButton.1", "cmdPoint"),
    ("Forms.CommandButton.1", "cmdFull"),
    ("Forms.CommandButton.1", "cmdSwap"),
    ("Forms.CommandButton.1", "cmdClose"),
    ("Forms.CommandButton.1", "cmdCreate"),
    ("Forms.CommandButton.1", "cmdDelete"),
    ("Forms.CommandButton.1", "cmdOfferYes"),
    ("Forms.CommandButton.1", "cmdOfferNo"),
]


def read_source(name):
    with open(os.path.join(HERE, name), encoding="ascii") as handle:
        text = handle.read()
    if name == "modLrsMain.bas":
        with open(os.path.join(HERE, "_logo_function.txt"), encoding="ascii") as handle:
            text = text.rstrip("\n") + "\n\n" + handle.read() + "\n"
    return text.replace("\r\n", "\n").replace("\n", "\r\n")


def find_project(vbe, path):
    for project in vbe.VBProjects:
        try:
            if os.path.normcase(os.path.abspath(project.FileName)) == os.path.normcase(path):
                return project
        except pythoncom.com_error:
            pass
    return None


def save_project(vbe, project):
    """File > Save <project> in the VBA editor (the only save MicroStation offers)."""
    vbe.MainWindow.Visible = True
    time.sleep(0.5)
    vbe.ActiveVBProject = project
    for control in vbe.CommandBars("Menu Bar").Controls("File").Controls:
        caption = control.Caption.replace("&", "")
        if caption.startswith("Save ") and not caption.startswith("Save As"):
            control.Execute()
            break
    time.sleep(1.0)
    vbe.MainWindow.Visible = False


def compact(source_path, target_path):
    """Copies the live content of a project file into a new, smaller compound file."""
    from win32com import storagecon
    read = storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE
    write = storagecon.STGM_CREATE | storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE
    forms = {name for name, kind, _ in MODULES if kind == FORM}

    def copy(source, target, path):
        for element in source.EnumElements():
            name, kind, size, clsid = element[0], element[1], element[2], element[8]
            if kind == storagecon.STGTY_STORAGE:
                if path.endswith("_VBA_Project/") and name != "VBA" and name not in forms:
                    continue                        # a form of the shell project
                child = target.CreateStorage(name, write, 0, 0)
                child.SetClass(clsid)
                copy(source.OpenStorage(name, None, read, None, 0), child, path + name + "/")
            elif not name.startswith("__SRP_"):     # compiled cache, rebuilt by VBA
                stream = target.CreateStream(name, write, 0, 0)
                if size:
                    stream.Write(source.OpenStream(name, None, read, 0).Read(size))

    if os.path.exists(target_path):
        os.remove(target_path)
    source = pythoncom.StgOpenStorage(source_path, None, read, None, 0)
    target = pythoncom.StgCreateDocfile(target_path, write, 0)
    target.SetClass(source.Stat(1)[8])
    copy(source, target, "")
    target.Commit(0)
    del source, target
    print("compacted", target_path, os.path.getsize(target_path), "bytes")


def main():
    if sys.argv[1] == "--compact":
        compact(os.path.abspath(sys.argv[2]), os.path.abspath(sys.argv[3]))
        return
    output = os.path.abspath(sys.argv[1])
    modules = MODULES + ([SELFTEST] if "--with-selftest" in sys.argv else [])
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    vbe = app.VBE

    project = find_project(vbe, output)
    if project is None:
        shell = sys.argv[sys.argv.index("--shell") + 1]
        shutil.copyfile(shell, output)
        app.CadInputQueue.SendKeyin('VBA LOAD "%s"' % output)
        for _ in range(40):
            time.sleep(0.5)
            project = find_project(vbe, output)
            if project is not None:
                break
    if project is None:
        raise SystemExit("The project could not be loaded: " + output)

    # Components are reused by name (removing and re-adding one with the same name fails
    # because the editor removes it lazily); anything else in the project is dropped.
    wanted = {name: kind for name, kind, _ in modules}
    existing = {}
    for component in list(project.VBComponents):
        if component.Type not in (STANDARD, CLASS, FORM):
            continue
        if wanted.get(component.Name) == component.Type:
            existing[component.Name] = component
        else:
            project.VBComponents.Remove(component)
    project.Name = PROJECT
    project.Description = "VDOT LRS Linear Milepost (VBA)"
    names = [reference.Name for reference in project.References]
    if "MicroStationDGN" not in names:
        project.References.AddFromGuid(MSTN_TYPELIB, 10, 1)

    for name, kind, source in modules:
        component = existing.get(name)
        if component is None:
            component = project.VBComponents.Add(kind)
            component.Name = name
        if kind == FORM:
            designer = component.Designer
            present = {control.Name for control in designer.Controls}
            for prog_id, control_name in FORM_CONTROLS:
                if control_name not in present:
                    designer.Controls.Add(prog_id, control_name, True)
        code = component.CodeModule
        if code.CountOfLines:
            code.DeleteLines(1, code.CountOfLines)
        code.AddFromString(read_source(source))
        print(f"  {name}: {code.CountOfLines} lines")
    print("references:", [reference.Name for reference in project.References])
    save_project(vbe, project)
    print("saved:", bool(project.Saved), output, os.path.getsize(output), "bytes")

if __name__ == "__main__":
    pythoncom.CoInitialize()
    main()
