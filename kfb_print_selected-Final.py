# -*- coding: utf-8 -*-
"""CP CODE 선택 및 파생상품 자동 인쇄 (기업 목록 내장).
엑셀 파일 없이 실행합니다. 기업명·사업자번호·JC/GC 정보가 이 파일에 포함됩니다.
설치: py -m pip install playwright pywinauto
      py -m playwright install chromium
실행: py kfb_print_selected-15.py
목록 변경 시 아래 EMBEDDED_TARGETS를 수정하세요.
"""
import pathlib
import re
import xml.etree.ElementTree as ET
from urllib.parse import parse_qs
import tkinter as tk
from tkinter import messagebox, ttk

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


# 기업 목록 내장: 이 파일을 공유하면 아래 기업 정보도 함께 전달됩니다.
EMBEDDED_TARGETS = [{'row': 2, 'cp': 'AMKR', 'category': 'GC', 'name': '앰코테크놀로지코리아(주)', 'bizno': '4098146857'},
 {'row': 3, 'cp': 'BASF', 'category': 'GC', 'name': 'BASF COMPANY LT', 'bizno': '4178104858'},
 {'row': 4, 'cp': 'BIRL', 'category': 'GC', 'name': '비를라카본코리아(주)', 'bizno': '4178117746'},
 {'row': 5, 'cp': 'BKPL', 'category': 'GC', 'name': 'BAKE PLUS CO.,', 'bizno': '1208188819'},
 {'row': 6, 'cp': 'BPPK', 'category': 'GC', 'name': 'BASF PERFORMANC', 'bizno': '2148629575'},
 {'row': 7, 'cp': 'BTAG', 'category': 'GC', 'name': 'BRENNTAG KOREA', 'bizno': '1148177689'},
 {'row': 8, 'cp': 'CHEL', 'category': 'GC', 'name': 'SAMSUNG C&T COR', 'bizno': '2028145975'},
 {'row': 9, 'cp': 'CJED', 'category': 'GC', 'name': 'CJ CHEILJEDANG', 'bizno': '1048609535'},
 {'row': 10, 'cp': 'CJFW', 'category': 'GC', 'name': 'CJFOOD SYSTEM C', 'bizno': '6038111270'},
 {'row': 11, 'cp': 'CJLO', 'category': 'GC', 'name': 'CJ LOGISTICS CO', 'bizno': '1108105034'},
 {'row': 12, 'cp': 'DAER', 'category': 'GC', 'name': 'DAELIM CO. LTD.', 'bizno': '1108130979'},
 {'row': 13, 'cp': 'DWGS', 'category': 'GC', 'name': '(주)대우건설', 'bizno': '1048158180'},
 {'row': 14, 'cp': 'GSCX', 'category': 'GC', 'name': 'GS CALTEX CORPO', 'bizno': '1168119123'},
 {'row': 15, 'cp': 'GSGB', 'category': 'GC', 'name': '(주)GS글로벌', 'bizno': '2028133028'},
 {'row': 16, 'cp': 'HASO', 'category': 'GC', 'name': '한화솔루션(주)', 'bizno': '2028116577'},
 {'row': 17, 'cp': 'HDM', 'category': 'GC', 'name': 'HYUNDAI MOTOR S', 'bizno': '1018109147'},
 {'row': 18, 'cp': 'HMOB', 'category': 'GC', 'name': '현대모비스(주)', 'bizno': '1018116406'},
 {'row': 19, 'cp': 'HSFT', 'category': 'GC', 'name': 'HSHYOSUNG ADVAN', 'bizno': '1988700929'},
 {'row': 20, 'cp': 'HSTN', 'category': 'GC', 'name': '효성티앤씨(주)', 'bizno': '8808701070'},
 {'row': 21, 'cp': 'HWCP', 'category': 'GC', 'name': '(주)한화', 'bizno': '2028116825'},
 {'row': 22, 'cp': 'HYST', 'category': 'GC', 'name': 'HYUNDAI STEEL C', 'bizno': '1218110385'},
 {'row': 23, 'cp': 'JJAR', 'category': 'GC', 'name': '(주)제주항공', 'bizno': '6168150527'},
 {'row': 24, 'cp': 'KALC', 'category': 'GC', 'name': 'Korean Air Line', 'bizno': '1108114794'},
 {'row': 25, 'cp': 'KCCG', 'category': 'GC', 'name': '주식회사 케이씨씨글라스', 'bizno': '4018701412'},
 {'row': 26, 'cp': 'KIAM', 'category': 'GC', 'name': 'KIA CORPORATION', 'bizno': '1198102316'},
 {'row': 27, 'cp': 'LGC', 'category': 'GC', 'name': 'LG CHEM, LTD.', 'bizno': '1078198139'},
 {'row': 28, 'cp': 'LGE', 'category': 'GC', 'name': 'LG ELECTRONICS', 'bizno': '1078614075'},
 {'row': 29, 'cp': 'LGHS', 'category': 'GC', 'name': 'LX HAUSYS,LTD.', 'bizno': '1078718122'},
 {'row': 30, 'cp': 'LGI', 'category': 'GC', 'name': 'LX INTERNATIONA', 'bizno': '1168118686'},
 {'row': 31, 'cp': 'LGPH', 'category': 'GC', 'name': 'LG DISPLAY CO.,', 'bizno': '1208155297'},
 {'row': 32, 'cp': 'LOCC', 'category': 'GC', 'name': 'LOTTE CHEMICAL', 'bizno': '1188115012'},
 {'row': 33, 'cp': 'LSNI', 'category': 'GC', 'name': 'LS MNM INC.', 'bizno': '6108136565'},
 {'row': 34, 'cp': 'LVLS', 'category': 'GC', 'name': 'NOVELIS KOREA', 'bizno': '5128106559'},
 {'row': 35, 'cp': 'MDFS', 'category': 'GC', 'name': 'MAHLE THERMAL A', 'bizno': '1348106892'},
 {'row': 36, 'cp': 'MHLE', 'category': 'GC', 'name': 'MAHLE BEHR KORE', 'bizno': '1098607361'},
 {'row': 37, 'cp': 'NAVE', 'category': 'GC', 'name': 'NAVER CORPORATI', 'bizno': '2208162517'},
 {'row': 38, 'cp': 'NEXT', 'category': 'GC', 'name': 'NEXEN TIRE CORP', 'bizno': '6218110769'},
 {'row': 39, 'cp': 'POCH', 'category': 'GC', 'name': 'POSCO FUTURE M', 'bizno': '5068101452'},
 {'row': 40, 'cp': 'POEC', 'category': 'GC', 'name': '(주)포스코이앤씨', 'bizno': '5068102280'},
 {'row': 41, 'cp': 'POSD', 'category': 'GC', 'name': 'POSCO INTERNATI', 'bizno': '1048158195'},
 {'row': 42, 'cp': 'POSN', 'category': 'GC', 'name': 'POSCO CO LTD', 'bizno': '3018702315'},
 {'row': 43, 'cp': 'PSCS', 'category': 'GC', 'name': 'SeAH Changwon I', 'bizno': '6098129083'},
 {'row': 44, 'cp': 'SEAH', 'category': 'GC', 'name': 'SeAH Besteel Co', 'bizno': '7948102456'},
 {'row': 45, 'cp': 'SILW', 'category': 'GC', 'name': '(주)엘엑스세미콘', 'bizno': '3148129147'},
 {'row': 46, 'cp': 'SKC', 'category': 'GC', 'name': 'SK INNOVATION C', 'bizno': '1018632319'},
 {'row': 47, 'cp': 'SKCC', 'category': 'GC', 'name': 'SKC CO., LTD.', 'bizno': '1358101250'},
 {'row': 48, 'cp': 'SKE', 'category': 'GC', 'name': 'SK Energy Co.,', 'bizno': '1018660120'},
 {'row': 49, 'cp': 'SKGC', 'category': 'GC', 'name': 'SK GEO CENTRIC', 'bizno': '1018660116'},
 {'row': 50, 'cp': 'SKHY', 'category': 'GC', 'name': 'HYNIX SEMICONDU', 'bizno': '1268103725'},
 {'row': 51, 'cp': 'SKIE', 'category': 'GC', 'name': 'SK IE TECHNOLOG', 'bizno': '5598601539'},
 {'row': 52, 'cp': 'SKMT', 'category': 'GC', 'name': '에스케이스페셜티 주식회사', 'bizno': '4598802246'},
 {'row': 53, 'cp': 'SKNX', 'category': 'GC', 'name': '에스케이넥실리스 주식회사', 'bizno': '1018702452'},
 {'row': 54, 'cp': 'SKON', 'category': 'GC', 'name': 'SK On Co., Ltd.', 'bizno': '2428702258'},
 {'row': 55, 'cp': 'SOIL', 'category': 'GC', 'name': 'S - OIL CORPORA', 'bizno': '1168136743'},
 {'row': 56, 'cp': 'SSHY', 'category': 'GC', 'name': '삼성중공업(주)', 'bizno': '1208152780'},
 {'row': 57, 'cp': 'STPC', 'category': 'GC', 'name': '한화토탈에너지스 주식회사', 'bizno': '3168104477'},
 {'row': 58, 'cp': 'SYBP', 'category': 'GC', 'name': 'SAMYANG BIOPHAR', 'bizno': '4548703296'},
 {'row': 59, 'cp': 'SYCP', 'category': 'GC', 'name': 'SAMYANG CORP.', 'bizno': '1018666838'},
 {'row': 60, 'cp': 'TRNK', 'category': 'GC', 'name': 'TRANE KOREA INC', 'bizno': '2148185820'},
 {'row': 61, 'cp': 'ASHI', 'category': 'JC', 'name': 'ASHIMORI KOREA', 'bizno': '2248152122'},
 {'row': 62, 'cp': 'DSMM', 'category': 'JC', 'name': '롯데엠시시 주식회사', 'bizno': '3168108546'},
 {'row': 63, 'cp': 'DWFC', 'category': 'JC', 'name': 'DONGWOO FINE', 'bizno': '4038106691'},
 {'row': 64, 'cp': 'FJIU', 'category': 'JC', 'name': 'FUJITSU KOREA L', 'bizno': '1028117053'},
 {'row': 65, 'cp': 'FRLK', 'category': 'JC', 'name': 'FRL KOREA CO.,L', 'bizno': '1048191189'},
 {'row': 66, 'cp': 'FUXE', 'category': 'JC', 'name': 'FUJIFILM BUSINE', 'bizno': '1108101270'},
 {'row': 67, 'cp': 'HIHI', 'category': 'JC', 'name': 'HITACHI HIGH-TE', 'bizno': '1018613592'},
 {'row': 68, 'cp': 'HOBA', 'category': 'JC', 'name': 'HORIBA KOREA LT', 'bizno': '1308123448'},
 {'row': 69, 'cp': 'HRDC', 'category': 'JC', 'name': 'HARADA CORPORAT', 'bizno': '2208404522'},
 {'row': 70, 'cp': 'HTLG', 'category': 'JC', 'name': 'HITACHI-LG DATA', 'bizno': '1078193060'},
 {'row': 71, 'cp': 'HYUN', 'category': 'JC', 'name': 'HYUNDAM INDUSTR', 'bizno': '3128114139'},
 {'row': 72, 'cp': 'ITOK', 'category': 'JC', 'name': 'ITOCHU KOREA LT', 'bizno': '1048129161'},
 {'row': 73, 'cp': 'IWAI', 'category': 'JC', 'name': 'SOJITZ KOREA CO', 'bizno': '1048136624'},
 {'row': 74, 'cp': 'KOAS', 'category': 'JC', 'name': 'KOASA KOREA COR', 'bizno': '2478601150'},
 {'row': 75, 'cp': 'KOMC', 'category': 'JC', 'name': 'MITSUBISHI CORP', 'bizno': '2028119837'},
 {'row': 76, 'cp': 'MCRJ', 'category': 'JC', 'name': '', 'bizno': '0124011628'},
 {'row': 77, 'cp': 'MIDI', 'category': 'JC', 'name': 'MDI KOREA CO.,', 'bizno': '1228179844'},
 {'row': 78, 'cp': 'MITK', 'category': 'JC', 'name': 'MITSUI AND CO.,', 'bizno': '2208103274'},
 {'row': 79, 'cp': 'MTEV', 'category': 'JC', 'name': 'MITSUBISHI ELEV', 'bizno': '1078609043'},
 {'row': 80, 'cp': 'NAKO', 'category': 'JC', 'name': 'NAGASE KOREA CO', 'bizno': '1048160667'},
 {'row': 81, 'cp': 'NAMY', 'category': 'JC', 'name': 'NAMYANG METALS', 'bizno': '5148109042'},
 {'row': 82, 'cp': 'NDEN', 'category': 'JC', 'name': 'KOREA NITTO DEN', 'bizno': '1068187852'},
 {'row': 83, 'cp': 'NMBK', 'category': 'JC', 'name': '한국엔엠비 ( 주 )', 'bizno': '2208124245'},
 {'row': 84, 'cp': 'NTN', 'category': 'JC', 'name': 'NTN KOREA CO.,', 'bizno': '1048160522'},
 {'row': 85, 'cp': 'RYO', 'category': 'JC', 'name': 'KOREA RYOSAN CO', 'bizno': '2148676568'},
 {'row': 86, 'cp': 'SANR', 'category': 'JC', 'name': 'SANRIO KOREA CO', 'bizno': '2148624755'},
 {'row': 87, 'cp': 'SONY', 'category': 'JC', 'name': 'SONY KOREA CORP', 'bizno': '1068123810'},
 {'row': 88, 'cp': 'SOSC', 'category': 'JC', 'name': 'SONY SEMICONDUC', 'bizno': '5738802765'},
 {'row': 89, 'cp': 'SSIL', 'category': 'JC', 'name': 'SHIN-ETSU SILIC', 'bizno': '1048120501'},
 {'row': 90, 'cp': 'STEM', 'category': 'JC', 'name': 'STEMCO.LTD.', 'bizno': '3078107204'},
 {'row': 91, 'cp': 'SUMI', 'category': 'JC', 'name': 'SUMITOMO CORPOR', 'bizno': '1028134616'},
 {'row': 92, 'cp': 'TADA', 'category': 'JC', 'name': 'TADANO KOREA CO', 'bizno': '1058183663'},
 {'row': 93, 'cp': 'TAKT', 'category': 'JC', 'name': 'TAK TEXTILES KO', 'bizno': '5138173333'},
 {'row': 94, 'cp': 'TBSF', 'category': 'JC', 'name': 'TORAY BATTERY S', 'bizno': '1018637954'},
 {'row': 95, 'cp': 'TOSE', 'category': 'JC', 'name': 'TORAY ADVANCED', 'bizno': '5138116410'},
 {'row': 96, 'cp': 'TTSU', 'category': 'JC', 'name': 'TOYOTA TSUSHO K', 'bizno': '2018180945'},
 {'row': 97, 'cp': 'USHI', 'category': 'JC', 'name': 'USHIO KOREA.,IN', 'bizno': '2148602260'},
 {'row': 98, 'cp': 'YANM', 'category': 'JC', 'name': 'YANMAR AGRICULT', 'bizno': '1248634270'},
 {'row': 99, 'cp': 'YONG', 'category': 'JC', 'name': '영신쿼츠(주)', 'bizno': '1048125433'}]


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


def verify_query_identity(payload, expected):
    """조회 요청의 식별번호를 확인하지 못하면 출력하지 않는다."""
    found = []
    try:
        root = ET.fromstring(payload)
        found = [(node.text or "").strip() for node in root.iter()
                 if node.attrib.get("id") == "KIKO_ID_NO"
                 and node.tag.rsplit("}", 1)[-1] in ("Col", "Parameter")]
    except ET.ParseError:
        # Nexacro SSV: 열 이름과 값이 각각 레코드에 들어간다.
        columns = None
        for record in payload.split("\x1e"):
            cells = record.split("\x1f")
            if record.startswith("Dataset:"):
                columns = None
            names = [c.split(":", 1)[0] for c in cells]
            if "KIKO_ID_NO" in names:
                columns = names
            elif columns and cells and cells[0] in ("N", "U", "I", "O"):
                values = cells[1:]
                idx = columns.index("KIKO_ID_NO")
                if idx < len(values):
                    found.append(values[idx].strip())
        if not found:
            found = parse_qs(payload).get("KIKO_ID_NO", [])
    if not found or any(value != expected for value in found):
        raise RuntimeError("선택한 기업과 실제 조회 요청의 사업자번호가 일치하는지 "
                           "확인하지 못해 출력을 중단했습니다.")


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
                predicate=lambda r: ("bizkiR0101SelectList.do" in r.url
                                     and r.request.frame.page == page),
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

    verify_query_identity(response.request.post_data or "", first + second)
    response.finished()
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
    items = [dict(item) for item in EMBEDDED_TARGETS]
    initial_targets = select_targets(items)
    if not initial_targets:
        return

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
            if initial_targets is not None:
                targets = initial_targets
                initial_targets = None
            else:
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
                    if index > 1:
                        if page != main_page and not page.is_closed():
                            page.close(run_before_unload=False)
                        page = open_query_popup(ctx, main_page)
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
