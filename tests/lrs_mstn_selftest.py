"""In-MicroStation check for VDOT_LRS_Linear_Milepost.py.

Run inside MicroStation with:  PYTHON LOAD "<this file>"

Opens the real dialog (with its own throw-away settings file), loads a jurisdiction and
routes through the background worker, exercises the opposite-direction, swap, full-route,
Linear/Point and pick-tool controls, and - only when the active file is the scratch copy
lrs_vba_scratch.dgn - creates three events. Results go to lrs_mstn_selftest_result.json
beside this file.
"""
import importlib.util
import json
import os
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(HERE, os.pardir, "package", "VDOT_LRS_Linear_Milepost",
                    "VDOT_LRS_Linear_Milepost.py")
RESULT = os.path.join(HERE, "lrs_mstn_selftest_result.json")
SCRATCH = "lrs_vba_scratch.dgn"

result = {"started": time.strftime("%Y-%m-%d %H:%M:%S"), "steps": [], "ok": False}


def step(name, **data):
    result["steps"].append(dict(step=name, **data))
    with open(RESULT, "w", encoding="utf-8") as handle:      # saved after every step
        json.dump(result, handle, indent=1, default=str)


try:
    spec = importlib.util.spec_from_file_location("vdot_lrs_tool", TOOL)
    tool = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(tool)
    step("import", in_mstn=tool.IN_MSTN, have_qt=tool.HAVE_QT, build=tool.BUILD_DATE)

    from PyQt5.QtCore import QSettings
    from PyQt5.QtWidgets import QApplication
    from MSPyMstnPlatform import ISessionMgr, PyCadInputQueue

    app = QApplication.instance() or QApplication(sys.argv)

    def pump(count=10):
        for _ in range(count):
            app.processEvents()
            PyCadInputQueue.PythonMainLoop()
            time.sleep(0.03)

    def wait(condition, seconds=300):
        started = time.time()
        while time.time() - started < seconds:
            pump(2)
            if condition():
                return round(time.time() - started, 1)
        raise RuntimeError("timed out waiting")

    def idle():
        return dialog._busy_text is None and dialog._on_result is None

    def choose(combo, text=None, data=None):
        index = combo.findText(text) if data is None else combo.findData(data)
        if index < 0:
            raise RuntimeError(f"'{text or data}' is not in the list")
        combo.setCurrentIndex(index)
        combo.activated.emit(index)
        pump(3)

    def measures(start, end):
        dialog.from_measure.setEditText(start)
        dialog.to_measure.setEditText(end)
        dialog._guard(dialog._measures_edited)
        pump(3)

    file_name = str(ISessionMgr.GetActiveDgnFile().GetFileName())
    is_scratch = os.path.basename(file_name).lower() == SCRATCH
    step("model", file=file_name, scratch=is_scratch, summary=tool.model_summary())

    ini = os.path.join(tempfile.gettempdir(), "vdot_lrs_mstn_selftest.ini")
    if os.path.exists(ini):
        os.remove(ini)
    dialog = tool.EventDialog(settings=QSettings(ini, QSettings.IniFormat))
    dialog.show()
    seconds = wait(lambda: dialog.jurisdiction.count() > 0 and idle())
    step("connected", seconds=seconds, status=dialog.status.text(),
         jurisdictions=dialog.jurisdiction.count(), selected=dialog.jurisdiction.currentIndex())

    choose(dialog.jurisdiction, "Washington County")
    busy = dialog._busy_text
    seconds = wait(lambda: dialog.route.count() > 0 and idle())
    step("routes", busy_while_loading=busy, seconds=seconds, count=dialog.route.count(),
         first=[dialog.route.itemText(i) for i in range(4)])

    choose(dialog.route, "I-81N")
    seconds = wait(lambda: dialog.route_info is not None and idle())
    pump(10)
    step("route", seconds=seconds, name_id=dialog.route_info.name_id, hint=dialog.route_hint.text(),
         converted=dialog.geometry_uor is not None, preview=len(dialog.route_preview),
         pick_active=tool.LrsPickTool.active is not None, pick_checked=dialog.pick_button.isChecked(),
         pick_text=dialog.pick_button.text(), from_label=dialog.from_label.text(),
         from_marked=dialog.from_label.property("next"),
         opposite_enabled=dialog.opposite_button.isEnabled())

    dialog.opposite_button.click()
    seconds = wait(lambda: dialog.route_info is not None and idle()
                   and dialog.route_info.name_id != "R-VA   IS00081NB")
    step("opposite", seconds=seconds, name_id=dialog.route_info.name_id,
         shown=dialog.route.currentText())

    choose(dialog.route, "SC-700E (Washington County)")
    seconds = wait(lambda: dialog.route_info is not None and idle()
                   and dialog.route_info.name_id == "R-VA095SC00700EB")
    pump(10)
    step("sc700", seconds=seconds, hint=dialog.route_hint.text(), event=dialog.event_hint.text(),
         opposite_enabled=dialog.opposite_button.isEnabled())

    # Picking: From at MP 12.6, then To behind it at MP 5.14 (must swap). The second pick
    # shows the "Draw it?" bar; answer No so nothing is created here.
    dialog.on_pick(*dialog.geometry_uor.point_at(12.6))
    pump(3)
    after_first = (dialog.next_pick, dialog.to_label.text(), dialog.to_label.property("next"),
                   dialog.offer_bar.isVisible())
    dialog.on_pick(*dialog.geometry_uor.point_at(5.14))
    pump(3)
    offer = (dialog.offer_bar.isVisible(), dialog.offer_text.text(),
             dialog.offer_yes.isVisible(), dialog.offer_no.isVisible())
    dialog.grab().save(os.path.join(HERE, "offer_bar.png"))
    dialog.offer_no.click()
    pump(3)
    step("pick", after_first=after_first, measures=[dialog.from_measure.currentText(),
                                                    dialog.to_measure.currentText()],
         event=dialog.event_hint.text(), offer=offer,
         hidden_after_no=not dialog.offer_bar.isVisible(), last_log=dialog._log_lines[-1:])
    result["offer_ok"] = (not after_first[3] and offer[0] and offer[2] and offer[3]
                          and not dialog.offer_bar.isVisible())

    measures("9", "4")
    invalid = (dialog.event_hint.text(), dialog.event_hint.property("error"))
    dialog.swap_button.click()
    pump(3)
    swapped = [dialog.from_measure.currentText(), dialog.to_measure.currentText()]
    dialog.full_button.click()
    pump(3)
    step("swap_full", invalid=invalid, swapped=swapped, full=dialog.event_hint.text(),
         preview=len(dialog.event_preview))

    dialog.point_button.click()
    pump(5)
    point = (dialog.to_measure.isVisible(), dialog.event_hint.text(), len(dialog.event_preview))
    dialog.linear_button.click()
    pump(5)
    step("point_toggle", point=point, to_visible_again=dialog.to_measure.isVisible())

    dialog.pick_button.click()
    pump(10)
    stopped = (tool.LrsPickTool.active is not None, dialog.pick_button.isChecked(),
               dialog.pick_button.text())
    dialog.pick_button.click()
    pump(10)
    step("pick_toggle", stopped=stopped, restarted=tool.LrsPickTool.active is not None,
         checked=dialog.pick_button.isChecked())

    if is_scratch:
        created = []
        for kind, start, end, mileposts in (("Linear", "2", "5", tool.MILEPOST_ENDS),
                                            ("Linear", "14", "17", 0.5),
                                            ("Point", "3.25", "3.25", tool.MILEPOST_ENDS)):
            (dialog.linear_button if kind == "Linear" else dialog.point_button).click()
            measures(start, end)
            choose(dialog.mileposts, data=mileposts)
            before = len(dialog._log_lines)
            dialog.create_button.click()
            pump(20)
            created.append(dialog._log_lines[before:])
            step("create", event=f"{kind} {start}-{end}", log=dialog._log_lines[before:])
        result["created_ok"] = all(lines and lines[-1].startswith("SUCCESS") for lines in created)
        # Pick From/To and answer Yes in the bar: the event is drawn.
        dialog.linear_button.click()
        pump(3)
        dialog.next_pick = "from"
        dialog.on_pick(*dialog.geometry_uor.point_at(6.0))
        dialog.on_pick(*dialog.geometry_uor.point_at(8.0))
        pump(3)
        before = len(dialog._log_lines)
        dialog.offer_yes.click()
        pump(20)
        yes_log = dialog._log_lines[before:]
        step("offer_yes", log=yes_log, hidden=not dialog.offer_bar.isVisible())
        result["offer_yes_ok"] = bool(yes_log) and yes_log[-1].startswith("SUCCESS")
        refs, counts = tool.find_created_elements()
        before = len(dialog._log_lines)
        dialog.delete_button.click()
        pump(5)
        asked = (dialog.offer_bar.isVisible(), dialog.offer_text.text(), dialog.offer_yes.text(),
                 QApplication.overrideCursor() is None)
        dialog.grab().save(os.path.join(HERE, "delete_bar.png"))
        dialog.offer_no.click()
        pump(3)
        kept, _ = tool.find_created_elements()
        dialog.delete_button.click()
        pump(5)
        dialog.offer_yes.click()
        pump(10)
        remaining, _ = tool.find_created_elements()
        step("delete", found=len(refs), counts=counts, asked=asked, kept_after_no=len(kept),
             log=dialog._log_lines[before:], remaining=len(remaining),
             cursor_clear=QApplication.overrideCursor() is None)
        result["deleted_ok"] = (bool(refs) and asked[0] and asked[3] and len(kept) == len(refs)
                                and not remaining and QApplication.overrideCursor() is None)
        dialog.linear_button.click()
        measures("2", "5")
        dialog._zoom()
        pump(10)

    step("log", errors=[line for line in dialog._log_lines if line.startswith(("ERROR", "WARNING"))],
         lines=len(dialog._log_lines))
    if "--keep-open" in os.environ.get("LRS_SELFTEST", ""):
        step("left_open")
    else:
        dialog.close()
        pump(10)
        step("closed", pick_tool_active=tool.LrsPickTool.active is not None,
             worker_alive=dialog._worker is not None and dialog._worker.is_alive())
    sys._vdot_lrs_event_dialog = None
    result["ok"] = True
except Exception:
    result["error"] = traceback.format_exc()
finally:
    result["finished"] = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(RESULT, "w", encoding="utf-8") as handle:
        json.dump(result, handle, indent=1, default=str)
    print("[LRS selftest] ok =", result["ok"], "->", RESULT)
