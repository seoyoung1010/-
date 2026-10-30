# -*- coding: utf-8 -*-
"""선택한 CP CODE의 파생상품 조회 화면을 순서대로 출력한다.

설치: py -m pip install playwright openpyxl
      py -m playwright install chromium
실행: py kfb_print_selected.py --excel "은행연합회 자동화용 - 양그룹장님거.xlsx"

로그인과 MarkAny의 최종 인쇄는 사용자가 직접 수행한다. 사이트 UI가 바뀌면
아래 SCREEN 상수와 open_query_popup 함수를 조정해야 한다.
"""
import argparse
import datetime as dt
import pathlib
import re
import tkinter as tk
from tkinter import messagebox, ttk

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent
START_URL = "https://bizinfo.kfb.or.kr/biz/index.html"
POPUP_URL = ("https://bizinfo.kfb.or.kr/biz/popup.html"
             "?formname=BIZKI%3A%3ABIZKI_T0101.xfdl"
             "&framename=KI01_BIZKI_T0101.xfdl")

# 첨부 화면(1440 x 900)에서 확인한 위치. Nexacro가 DOM 입력칸을 만들지
# 않는 경우에도 이 좌표를 이용해 사용자가 하는 것처럼 클릭/입력한다.
SCREEN = {"first": (296, 132), "second": (402, 132),
          "search": (1195, 132), "print": (1285, 132)}


def load_targets(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["파생거래"]
        result, blanks = [], 0
        for rowno, (raw, _, name, cp, category) in enumerate(
                ws.iter_rows(min_row=2, max_col=5, values_only=True), 2):
            if all(x is None for x in (raw, name, cp, category)):
                blanks += 1
                # 일부 파일은 서식만 수십만 행에 남아 있다.
                if blanks >= 100:
                    break
                continue
            blanks = 0
            if not cp or raw is None:
                continue
            digits = re.sub(r"\D", "", str(raw))
            if len(digits) == 9:
                digits = digits.zfill(10)  # 엑셀이 버린 선행 0 복원
            if len(digits) != 10:
                raise ValueError(f"파생거래 {rowno}행: 사업자번호 10자리 아님: {raw!r}")
            result.append({"row": rowno, "cp": str(cp).strip(),
                           "category": str(category or "미분류").strip().upper(),
                           "name": str(name or "").strip(), "bizno": digits})
        return result
    finally:
        wb.close()


def select_targets(items, summary=""):
    root = tk.Tk()
    root.title("출력할 CP CODE 선택")
    root.geometry("800x700")
    style = ttk.Style(root)
    style.configure("CP.Treeview", rowheight=38, font=("맑은 고딕", 13))
    style.configure("CP.Treeview.Heading", font=("맑은 고딕", 12, "bold"))
    selected = []
    if summary:
        ttk.Label(root, text=summary, foreground="#17643a").pack(pady=(10, 0))
    tk.Label(root, text="기업 구분을 고른 뒤 행을 클릭하여 선택하세요.").pack(pady=10)
    category = tk.StringVar(value="전체")
    categories = ["전체"] + [x for x in ("JC", "GC")
                           if any(i["category"] == x for i in items)]
    categories += sorted({i["category"] for i in items} - {"JC", "GC"})
    tabs = ttk.Frame(root)
    tabs.pack(fill="x", padx=12, pady=(0, 8))
    for cat in categories:
        count = len(items) if cat == "전체" else sum(i["category"] == cat for i in items)
        ttk.Radiobutton(tabs, text=f"{cat} ({count})", variable=category,
                        value=cat).pack(side="left", padx=7)
    query = tk.StringVar()
    search = ttk.Frame(root)
    search.pack(fill="x", padx=12)
    ttk.Label(search, text="CP CODE 검색").pack(side="left", padx=(0, 8))
    ttk.Entry(search, textvariable=query).pack(side="left", fill="x", expand=True)
    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True, padx=12, pady=8)
    table = ttk.Treeview(frame, columns=("check", "cp", "name"), show="headings",
                         selectmode="none", style="CP.Treeview")
    table.heading("check", text="선택")
    table.heading("cp", text="CP CODE")
    table.heading("name", text="기업명")
    table.column("check", width=75, anchor="center", stretch=False)
    table.column("cp", width=125, anchor="center", stretch=False)
    table.column("name", width=475, anchor="w")
    scroll = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
    table.configure(yscrollcommand=scroll.set)
    table.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    checked = set()  # 엑셀 행 번호. 필터를 바꿔도 선택 상태를 유지한다.

    def update_count():
        count_label.config(text=f"선택 {len(checked)}개")

    def filter_list(*_):
        term = query.get().strip().casefold()
        table.delete(*table.get_children())
        for item in items:
            if (category.get() in ("전체", item["category"])
                    and term in item["cp"].casefold()):
                table.insert("", "end", iid=str(item["row"]),
                             values=("☑" if item["row"] in checked else "☐",
                                     item["cp"], item["name"]))
        table.yview_moveto(0)
        update_count()

    def toggle(event):
        row = table.identify_row(event.y)
        if not row:
            return
        key = int(row)
        if key in checked:
            checked.remove(key)
        else:
            checked.add(key)
        table.set(row, "check", "☑" if key in checked else "☐")
        update_count()

    table.bind("<Button-1>", toggle)

    def mousewheel(event):
        if event.delta:
            table.yview_scroll(-int(event.delta / 120), "units")
        elif event.num in (4, 5):
            table.yview_scroll(-1 if event.num == 4 else 1, "units")

    # 표와 스크롤바 어느 쪽에 커서가 있어도 휠로 목록을 움직인다.
    for widget in (table, scroll):
        widget.bind("<MouseWheel>", mousewheel)
        widget.bind("<Button-4>", mousewheel)
        widget.bind("<Button-5>", mousewheel)

    buttons = ttk.Frame(root)
    buttons.pack(fill="x", padx=12, pady=10)
    count_label = ttk.Label(buttons)

    def mark_visible(value):
        for row in table.get_children():
            key = int(row)
            if value:
                checked.add(key)
            else:
                checked.discard(key)
            table.set(row, "check", "☑" if value else "☐")
        update_count()

    ttk.Button(buttons, text="현재 목록 선택", command=lambda: mark_visible(True)).pack(side="left")
    ttk.Button(buttons, text="현재 목록 해제", command=lambda: mark_visible(False)).pack(side="left", padx=8)
    count_label.pack(side="left", padx=10)
    query.trace_add("write", filter_list)
    category.trace_add("write", filter_list)
    filter_list()
    update_count()

    def finish():
        selected.extend(item for item in items if item["row"] in checked)
        if not selected:
            messagebox.showinfo("선택 필요", "CP CODE를 하나 이상 선택해 주세요.", parent=root)
            return
        root.destroy()

    ttk.Button(buttons, text="출력 시작", command=finish).pack(side="right")
    ttk.Button(buttons, text="프로그램 종료", command=root.destroy).pack(side="right", padx=8)
    root.mainloop()
    return selected


def prompt(title, message, choices):
    """사이트 바깥의 확인창. 창을 닫거나 중단을 누르면 즉시 멈춘다."""
    root = tk.Tk()
    root.title(title)
    root.geometry("500x220")
    root.attributes("-topmost", True)
    ttk.Label(root, text=message, wraplength=460, justify="left").pack(padx=20, pady=26)
    answer = [None]
    bar = ttk.Frame(root)
    bar.pack(pady=8)
    for label, value in choices:
        def choose(v=value):
            answer[0] = v
            root.destroy()
        ttk.Button(bar, text=label, command=choose).pack(side="left", padx=6)
    root.mainloop()
    return answer[0]


def live_pages(ctx):
    return [p for p in ctx.pages if not p.is_closed()]


def open_query_popup(ctx, main):
    """로그인 후 메뉴를 열어본다. Nexacro 메뉴 탐지가 어려우면 팝업 URL 사용."""
    for label in ("파생상품",):
        try:
            main.get_by_text(label, exact=True).first.click(timeout=2500)
            main.wait_for_timeout(400)
        except Exception:
            pass
    for label in ("파생상품/FX거래/CRS거래/옵션거래", "파생상품"):
        try:
            main.get_by_text(label, exact=True).last.click(timeout=1800)
            main.wait_for_timeout(800)
        except Exception:
            pass
        for p in live_pages(ctx):
            if p != main and ("BIZKI_T0101" in p.url or "popup.html" in p.url):
                return p
    # 기존 프로그램의 검증된 화면 주소를 사용한다. 로그인 세션은 동일 브라우저다.
    p = ctx.new_page()
    p.set_viewport_size({"width": 1440, "height": 900})
    p.goto(POPUP_URL, wait_until="domcontentloaded")
    p.wait_for_timeout(1000)
    return p


def click_at(page, point):
    page.mouse.click(*point)


def clear_at(page, point):
    click_at(page, point)
    page.keyboard.press("ControlOrMeta+A")
    # Backspace는 빈 칸에서 이전 입력칸으로 포커스를 옮기는 화면이 있다.
    page.keyboard.press("Delete")


def dismiss_confirm(page):
    """웹 화면에 표시되는 조회 확인창만 처리한다."""
    for label in ("확인", "OK"):
        try:
            button = page.get_by_text(label, exact=True).last
            if button.is_visible(timeout=250):
                button.click(timeout=1000)
                return
        except Exception:
            pass


def handle_one(ctx, page, item):
    # 뒤 칸을 먼저 비운다. 앞 칸의 숫자를 지울 때 뒤 칸에서 역방향으로
    # 포커스가 이동하는 사이트라 앞 칸부터 지우면 숫자가 엉킬 수 있다.
    first = "999" + item["bizno"][:4]
    second = item["bizno"][4:]
    assert len(first) == 7 and len(second) == 6
    page.bring_to_front()
    clear_at(page, SCREEN["second"])
    clear_at(page, SCREEN["first"])
    click_at(page, SCREEN["first"])
    # 첫 칸 7자리가 차면 사이트가 자동으로 뒤 칸으로 이동한다.
    # 중간에 두 번째 칸을 다시 클릭하지 않고 연속으로 입력해야 한다.
    page.keyboard.type(first + second, delay=70)
    # 접근 가능한 실제 input이 있으면 입력 결과를 확인한다. Nexacro 마스킹
    # 필드는 값 노출이 제한될 수 있으므로 확인이 불가능하면 사용자에게 넘긴다.
    values = page.evaluate("""pts => pts.map(p => {
        let e = document.elementFromPoint(p[0], p[1]);
        if (e && e.tagName === 'INPUT') return e.value;
        return null;
    })""", [SCREEN["first"], SCREEN["second"]])
    if values[0] is not None and values[0].isdigit() and values[0] != first:
        raise RuntimeError(f"첫 입력칸 불일치: {values[0]!r}")
    if values[1] is not None and values[1].isdigit() and values[1] != second:
        raise RuntimeError("둘째 입력칸 불일치")

    # 실제 조회 요청의 응답을 기다린다. 입력만 되고 조회가 실행되지 않은
    # 상태에서 출력으로 넘어가는 일을 방지한다.
    try:
        with ctx.expect_event(
                "response",
                predicate=lambda r: "bizkiR0101SelectList.do" in r.url,
                timeout=25000) as response_event:
            # 사용자가 확인한 사이트 동작: 번호 입력칸에서 Enter로 조회.
            click_at(page, SCREEN["second"])
            page.keyboard.press("Enter")
            page.wait_for_timeout(450)
            dismiss_confirm(page)
        response = response_event.value
    except Exception as e:
        raise RuntimeError("Enter로 조회했지만 사이트 조회 응답이 확인되지 않았습니다. "
                           "조회 화면 또는 팝업 상태를 확인해 주세요.") from e
    if response.status != 200:
        raise RuntimeError(f"조회 요청 실패: HTTP {response.status}")

    page.wait_for_timeout(900)
    dismiss_confirm(page)
    page.bring_to_front()
    # Nexacro의 출력 글자는 일반 button 태그가 아닐 수도 있으므로
    # 역할과 정확한 텍스트를 순서대로 찾는다. 좌표 추측 클릭은 하지 않는다.
    printed = False
    for frame in page.frames:
        for locator in (frame.get_by_role("button", name="출력", exact=True),
                        frame.get_by_text("출력", exact=True)):
            matches = [locator.nth(i) for i in range(locator.count())
                       if locator.nth(i).is_visible()]
            if len(matches) == 1:
                matches[0].click(timeout=8000)
                printed = True
                break
        if printed:
            break
    if not printed:
        raise RuntimeError("조회는 끝났지만 화면에서 출력 버튼을 찾지 못했습니다. "
                           "표시된 버튼 이름을 확인해 주세요.")
    page.wait_for_timeout(1000)
    # 화면 안에서만 뜨는 '출력' 확인창은 처리한다. 별도 MarkAny 창의
    # 최종 인쇄는 사용자가 직접 누른다.
    dismiss_confirm(page)


def try_markany_print(page):
    """MarkAny 창의 식별 가능한 인쇄 버튼을 한 번만 누른다."""
    try:
        from pywinauto import Desktop
    except ImportError:
        return [], "자동 인쇄 준비 필요: CMD에서 py -m pip install pywinauto 실행"
    windows = []
    for _ in range(20):
        windows = [w for w in Desktop(backend="uia").windows()
                   if re.search(r"mark\s*any", w.window_text(), re.I)
                   and w.is_visible()]
        if windows:
            break
        page.wait_for_timeout(500)
    if len(windows) != 1:
        return [], "MarkAny 창을 하나로 식별하지 못했습니다. 직접 인쇄해 주세요."
    window = windows[0]
    labels = []
    candidates = []
    for button in window.descendants(control_type="Button"):
        name = button.window_text().strip()
        labels.append(name)
        normalized = re.sub(r"\(&.\)|&|\s", "", name).casefold()
        if normalized in ("인쇄", "출력", "print", "프린트") and button.is_enabled() and button.is_visible():
            candidates.append(button)
    (ROOT / "MarkAny_버튼확인.txt").write_text(
        "MarkAny 버튼 이름:\n" + "\n".join(labels), encoding="utf-8")
    if len(candidates) != 1:
        return windows, "인쇄 아이콘 이름을 식별하지 못했습니다. 직접 인쇄 후 버튼확인 파일을 알려주세요."
    try:
        window.set_focus()
        candidates[0].click_input()
    except Exception:
        return windows, "자동 클릭 결과를 확인할 수 없습니다. 중복 인쇄하지 않도록 화면을 확인해 주세요."
    # 아이콘 클릭 직후 열린 창이 키보드 포커스를 받는다.
    # 별도 팝업 검색/포커스 재설정을 하지 않고 사용자 동작 그대로 Enter.
    from pywinauto.keyboard import send_keys
    page.wait_for_timeout(2000)
    send_keys("{ENTER}", pause=0.1)
    # 출력 요청이 처리될 시간을 준 뒤 호출부에서 MarkAny 창을 닫는다.
    page.wait_for_timeout(3000)
    return windows, "인쇄 명령 전달 완료"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", type=pathlib.Path, default=ROOT / "은행연합회 자동화용 - 양그룹장님거.xlsx")
    args = ap.parse_args()
    items = load_targets(args.excel)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(ROOT / "browser_profile"), headless=False,
            viewport={"width": 1440, "height": 900}, locale="ko-KR",
            timezone_id="Asia/Seoul", accept_downloads=True,
            args=["--disable-popup-blocking"])
        ctx.on("page", lambda pg: pg.on("dialog", lambda d: d.accept()))
        main_page = ctx.pages[0] if ctx.pages else ctx.new_page()
        main_page.on("dialog", lambda d: d.accept())
        main_page.goto(START_URL, wait_until="domcontentloaded")
        if prompt("로그인", "사이트에 직접 로그인하고 추가 인증까지 마친 뒤 누르세요.",
                  [("로그인 완료", True), ("중단", False)]) is not True:
            ctx.close()
            return
        first_batch = True
        summary = ""
        while not main_page.is_closed():
            main_page.bring_to_front()
            targets = select_targets(items, summary)
            if not targets:
                break
            page = open_query_popup(ctx, main_page)
            if first_batch:
                if prompt("화면 확인", "파생상품 조회 화면이 열렸는지 확인하세요.",
                          [("조회 화면 준비 완료", True), ("중단", False)]) is not True:
                    break
                first_batch = False
            done = 0
            for index, item in enumerate(targets, 1):
                before = set(live_pages(ctx))
                markany_windows = []
                action = None
                try:
                    handle_one(ctx, page, item)
                    markany_windows, print_message = try_markany_print(page)
                    if print_message != "인쇄 명령 전달 완료":
                        raise RuntimeError(print_message)
                    for window in markany_windows:
                        try:
                            window.close()
                        except Exception:
                            pass
                    # UIA 창 객체의 visible 값으로 기다리면 이미 닫힌 창을
                    # 계속 열린 것으로 판단할 수 있다. 브라우저 출력창부터 정리.
                    for extra in list(live_pages(ctx)):
                        if extra != page and extra != main_page:
                            extra.close(run_before_unload=False)
                    page.wait_for_timeout(700)
                    leftovers = [p for p in live_pages(ctx)
                                 if p != page and p != main_page]
                    if leftovers:
                        raise RuntimeError("브라우저 출력 팝업을 닫지 못했습니다.")
                    if page.is_closed():
                        raise RuntimeError("조회창이 닫혀 다음 기업을 처리할 수 없습니다.")
                    page.bring_to_front()
                    markany_windows = []
                    action = "done"
                    print(f"[{index}/{len(targets)}] {item['cp']}: 인쇄 명령 전달 완료")
                    status = ("인쇄 명령 전달" if action == "done" else
                              "중단" if action in (None, "stop") else "실패")
                    done += action == "done"
                    if action is None or action == "stop":
                        break
                except Exception as e:
                    print(f"{item['cp']}: {e}")
                    prompt("작업 중단", f"{item['cp']}: {e}\n현재 건을 확인해 주세요. 자동 재인쇄하지 않습니다.",
                           [("닫기", True)])
                    break
                finally:
                    if action == "done":
                        for window in markany_windows:
                            try:
                                if window.exists():
                                    window.close()
                            except Exception:
                                pass
                    # 조회/출력 중 새로 열린 브라우저 창만 닫는다. MarkAny의
                    # Windows 네이티브 창은 사용자 인쇄/닫기 후 진행한다.
                    for extra in list(live_pages(ctx)):
                        if extra != page and extra != main_page:
                            try:
                                extra.close()
                            except Exception:
                                pass
            # 묶음 완료 후 조회 팝업도 닫고 로그인 메인창으로 돌아간다.
            if page != main_page and not page.is_closed():
                page.close(run_before_unload=False)
            summary = f"직전 묶음: {done}/{len(targets)}건 인쇄 명령 전달. 다음 기업을 선택하세요."
        ctx.close()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        error = traceback.format_exc()
        path = ROOT / "실행오류.txt"
        path.write_text(error, encoding="utf-8")
        print(error)
        prompt("실행 오류", "실행 중 오류가 발생했습니다.\n같은 폴더의 실행오류.txt 내용을 알려주세요.",
               [("닫기", True)])
