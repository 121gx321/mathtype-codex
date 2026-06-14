from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
import time
from pathlib import Path


WD_STORY = 6
WD_FIND_STOP = 0


def load_tool(tool_script: Path):
    spec = importlib.util.spec_from_file_location("mtww", tool_script)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load MathType-Word-WPS script: {tool_script}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["mtww"] = module
    spec.loader.exec_module(module)
    return module


def current_text(mtww, doc) -> str:
    return str(mtww.call_with_retry(lambda: doc.Content.Text))


def remaining_items(mtww, doc, mapping: list[dict[str, str]]) -> list[dict[str, str]]:
    text = current_text(mtww, doc)
    items = [item for item in mapping if item["placeholder"] in text]
    items.sort(key=lambda item: text.rfind(item["placeholder"]), reverse=True)
    return items


def select_placeholder(mtww, app, token: str) -> None:
    sel = mtww.call_with_retry(lambda: app.ActiveWindow.Selection)
    mtww.call_with_retry(lambda sel=sel: sel.EndKey(Unit=WD_STORY))
    find = mtww.call_with_retry(lambda sel=sel: sel.Find)
    mtww.call_with_retry(lambda find=find: find.ClearFormatting())
    find.Text = token
    find.Forward = False
    find.Wrap = WD_FIND_STOP
    find.MatchWildcards = False
    found = bool(mtww.call_with_retry(lambda find=find: find.Execute()))
    if not found:
        raise RuntimeError(f"Could not select placeholder: {token}")


def selected_paragraph_index(mtww, doc, app) -> int:
    start = int(mtww.call_with_retry(lambda: app.ActiveWindow.Selection.Range.Start))
    count = int(mtww.call_with_retry(lambda: doc.Range(0, start).Paragraphs.Count))
    return max(0, count - 1)


def insert_at_selection(mtww, app, doc, latex: str, *, font_pt: float, font_family: str, wait: float) -> int:
    import win32con
    import win32gui

    sel = mtww.call_with_retry(lambda: app.ActiveWindow.Selection)
    para_idx = selected_paragraph_index(mtww, doc, app)
    mtww.call_with_retry(lambda sel=sel: sel.Delete())

    mathml = mtww.normalize_mathml_for_mathtype_preview(mtww.latex_to_mathml(latex))
    before = set(mtww.list_mathtype_windows())
    shape = mtww.call_with_retry(
        lambda sel=sel: sel.InlineShapes.AddOLEObject(
            ClassType="Equation.DSMT4",
            FileName="",
            LinkToFile=False,
            DisplayAsIcon=False,
        )
    )
    try:
        hwnd = mtww.find_new_mathtype_window(before, timeout=20)
    except TimeoutError:
        mtww.call_with_retry(lambda shape=shape: shape.OLEFormat.DoVerb(0))
        hwnd = mtww.find_new_mathtype_window(before, timeout=20)

    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, mtww.MATHTYPE_SELECT_ALL_COMMAND_ID, 0)
    time.sleep(0.25)
    sized = mtww.apply_mathml_font_size(mathml, font_pt, font_family=font_family)
    mtww.set_mathml_clipboard(sized)
    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, mtww.MATHTYPE_PASTE_COMMAND_ID, 0)
    time.sleep(wait)
    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, mtww.MATHTYPE_UPDATE_HOST_COMMAND_ID, 0)
    time.sleep(wait)
    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, mtww.MATHTYPE_CLOSE_RETURN_COMMAND_ID, 0)

    deadline = time.time() + 20
    while time.time() < deadline and win32gui.IsWindow(hwnd):
        time.sleep(0.25)
    if win32gui.IsWindow(hwnd):
        raise RuntimeError("MathType editor did not close after update.")
    return para_idx


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--map-json", required=True, type=Path)
    parser.add_argument("--tool-script", required=True, type=Path)
    parser.add_argument("--limit", type=int, default=1)
    parser.add_argument("--backend", default="Word.Application")
    parser.add_argument("--font-pt", type=float, default=8.4)
    parser.add_argument("--font-family", default="Times New Roman")
    parser.add_argument("--wait", type=float, default=1.0)
    args = parser.parse_args()

    mtww = load_tool(args.tool_script.resolve())
    src = args.input.resolve()
    out = args.output.resolve()
    if src != out:
        shutil.copyfile(src, out)
    mtww.wait_for_exclusive_file(out, timeout=30)

    mapping = json.loads(args.map_json.resolve().read_text(encoding="utf-8"))
    pythoncom = None
    app = None
    doc = None
    try:
        pythoncom, app = mtww.open_word_app(args.backend, visible=False, allow_fallback=False)
        doc = mtww.open_document(app, out, read_only=False, com_progid=args.backend)
        mtww.call_with_retry(lambda: doc.Activate())
        items = remaining_items(mtww, doc, mapping)
        if args.limit:
            items = items[: args.limit]
        print(f"planned_replacements={len(items)}")
        for n, item in enumerate(items, 1):
            token = item["placeholder"]
            latex = item["latex"]
            select_placeholder(mtww, app, token)
            para_idx = insert_at_selection(
                mtww,
                app,
                doc,
                latex,
                font_pt=args.font_pt,
                font_family=args.font_family,
                wait=args.wait,
            )
            print(f"replace {n}/{len(items)} token={token} paragraph={para_idx} latex={latex}")
        mtww.call_with_retry(lambda: doc.Save())
    finally:
        try:
            if doc is not None:
                mtww.call_with_retry(lambda: doc.Close(False))
        finally:
            try:
                if app is not None:
                    mtww.call_with_retry(lambda: app.Quit())
            finally:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
    mtww.wait_for_exclusive_file(out, timeout=30)
    print(f"output={out}")


if __name__ == "__main__":
    main()
