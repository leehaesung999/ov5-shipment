"""OV5 Lock 재고 → 거래처 자동 매칭 GUI."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path
from tkinter import (
    BooleanVar, DoubleVar, StringVar, Tk, filedialog, messagebox,
    ttk, END, N, S, E, W,
)

sys.path.insert(0, str(Path(__file__).parent))
from core.stock_loader import load_locked_stock
from core.order_loader import load_orders
from core.master_loader import load_shelf_life_months, load_fs_ms_items, load_lot_assignments
from core.matcher import match
from core.writer import write_summary

ROOT = Path(__file__).parent
CONFIG_DIR = ROOT / "config"
SETTINGS_PATH = CONFIG_DIR / "settings.json"
MASTER_CACHE = CONFIG_DIR / "master_info.xlsx"
FS_MS_PATH = CONFIG_DIR / "fs_ms_items.xlsx"
LOT_PATH = CONFIG_DIR / "lot_assignments.xlsx"          # 수동 등록 (영구)
LOT_AUTO_PATH = CONFIG_DIR / "lot_assignments_auto.xlsx"  # 자동 누적 (재고 따라 정리)

DEFAULT_SETTINGS = {
    "threshold": 0.70,
    "ov_locations": ["OV5"],
    "nh_branch_keywords": {
        "포천": "농협포천", "평택": "농협평택", "횡성": "농협횡성",
        "군위": "농협군위", "장성": "농협장성", "경남": "농협경남",
    },
    "excluded_nh_keywords": ["제주"],
    "last_stock_dir": "", "last_orders_dir": "", "last_output_dir": "",
}


def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            s = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            return {**DEFAULT_SETTINGS, **s}
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()


def save_settings(s: dict) -> None:
    CONFIG_DIR.mkdir(exist_ok=True)
    SETTINGS_PATH.write_text(
        json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("OV5 지정출고 자동매칭")
        root.geometry("1180x760")
        self.settings = load_settings()
        self.last_result = None  # 매칭 후 결과 DataFrame

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_tab = ttk.Frame(nb)
        self.settings_tab = ttk.Frame(nb)
        nb.add(self.main_tab, text="  지정출고 매칭  ")
        nb.add(self.settings_tab, text="  기준정보 관리  ")

        self._build_main_tab()
        self._build_settings_tab()

    # ---------- 메인 탭 ----------
    def _build_main_tab(self):
        f = self.main_tab
        for i in range(3):
            f.columnconfigure(i, weight=1 if i == 1 else 0)
        f.rowconfigure(4, weight=1)

        self.stock_var = StringVar()
        self.orders_var = StringVar()
        self.out_var = StringVar()

        ttk.Label(f, text="재고 파일 (로케이션별 재고조회)").grid(
            row=0, column=0, sticky=W, padx=8, pady=(12, 2), columnspan=3)
        ttk.Entry(f, textvariable=self.stock_var).grid(
            row=1, column=0, columnspan=2, sticky=W + E, padx=(8, 4))
        ttk.Button(f, text="찾아보기...", command=self._pick_stock).grid(
            row=1, column=2, padx=(0, 8), sticky=W + E)

        ttk.Label(f, text="주문 파일 (출고진행현황 양식)").grid(
            row=2, column=0, sticky=W, padx=8, pady=(12, 2), columnspan=3)
        ttk.Entry(f, textvariable=self.orders_var).grid(
            row=3, column=0, columnspan=2, sticky=W + E, padx=(8, 4))
        ttk.Button(f, text="찾아보기...", command=self._pick_orders).grid(
            row=3, column=2, padx=(0, 8), sticky=W + E)

        ttk.Label(f, text="출력 폴더 (비워두면 재고 파일과 같은 위치)").grid(
            row=5, column=0, sticky=W, padx=8, pady=(12, 2), columnspan=3)
        ttk.Entry(f, textvariable=self.out_var).grid(
            row=6, column=0, columnspan=2, sticky=W + E, padx=(8, 4))
        ttk.Button(f, text="찾아보기...", command=self._pick_out).grid(
            row=6, column=2, padx=(0, 8), sticky=W + E)

        btn_frame = ttk.Frame(f)
        btn_frame.grid(row=7, column=0, columnspan=3, sticky=W + E, padx=8, pady=12)
        self.run_btn = ttk.Button(btn_frame, text="① 매칭 실행 (미리보기)", command=self._run)
        self.run_btn.pack(side="left", padx=(0, 8))
        self.save_btn = ttk.Button(btn_frame, text="② Excel로 저장", command=self._save_excel,
                                   state="disabled")
        self.save_btn.pack(side="left", padx=(0, 8))
        self.open_btn = ttk.Button(btn_frame, text="결과 폴더 열기",
                                   command=self._open_out, state="disabled")
        self.open_btn.pack(side="left")
        self.summary_var = StringVar(value="")
        ttk.Label(btn_frame, textvariable=self.summary_var, foreground="#0a6").pack(
            side="left", padx=(16, 0))

        ttk.Label(f, text="진행 로그").grid(
            row=8, column=0, sticky=W, padx=8, pady=(8, 2))
        log_frame = ttk.Frame(f)
        log_frame.grid(row=9, column=0, columnspan=3, sticky="nsew", padx=8, pady=(0, 8))
        from tkinter import Text, Scrollbar
        self.log_box = Text(log_frame, height=5, wrap="word")
        self.log_box.pack(side="left", fill="both", expand=True)
        sb = Scrollbar(log_frame, command=self.log_box.yview)
        sb.pack(side="right", fill="y")
        self.log_box.config(yscrollcommand=sb.set)

        preview_header = ttk.Frame(f)
        preview_header.grid(row=10, column=0, columnspan=3, sticky=W + E, padx=8, pady=(8, 2))
        ttk.Label(preview_header, text="결과 미리보기 (저장 전 검토)",
                  font=("Segoe UI", 9, "bold")).pack(side="left")
        self.matched_only_var = BooleanVar(value=False)
        ttk.Checkbutton(
            preview_header, text="매칭된 항목만 보기 / 저장",
            variable=self.matched_only_var,
            command=self._refresh_preview).pack(side="left", padx=(16, 0))

        tree_frame = ttk.Frame(f)
        tree_frame.grid(row=11, column=0, columnspan=3, sticky="nsew", padx=8, pady=(0, 12))
        f.rowconfigure(11, weight=1)
        self._build_tree(tree_frame)

        last_out = self.settings.get("last_output_dir") or ""
        if last_out: self.out_var.set(last_out)
        self._log(f"준비 완료. 잔존율 기준 {self.settings['threshold']*100:.0f}%, "
                  f"OV 로케이션 {','.join(self.settings['ov_locations'])}.")
        if not MASTER_CACHE.exists():
            self._log("⚠ 기준정보 캐시 없음 — [기준정보 관리] 탭에서 등록하세요.")

    def _build_tree(self, parent):
        from tkinter import Scrollbar
        from core.matcher import build_output_cols, DEFAULT_NH_COLS
        self.tree_parent = parent
        self.tree_cols = build_output_cols(DEFAULT_NH_COLS)
        self.tree = ttk.Treeview(parent, columns=self.tree_cols, show="headings", height=14)
        self._apply_tree_columns()
        vsb = Scrollbar(parent, orient="vertical", command=self.tree.yview)
        hsb = Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        self.tree.tag_configure("ok", background="#E7F4E4")
        self.tree.tag_configure("ng", background="#FAD9C9")
        self.tree.tag_configure("matched", background="#FFF7CC")

    def _apply_tree_columns(self):
        widths = {
            "Item ID": 70, "Item": 230, "유통기한": 75, "재고수량(BOX)": 70,
            "로케이션": 60, "잔존율": 65, "판정구분": 65, "매칭수량": 60,
            "거래처명": 220, "대리점": 50,
            "잔여량": 55, "유통기한(월)": 75, "제조일자": 85,
            "남은율": 65, "판정": 50,
        }
        self.tree.configure(columns=self.tree_cols)
        for c in self.tree_cols:
            self.tree.heading(c, text=c)
            anchor = "w" if c in ("Item", "거래처명") else "center"
            w = widths.get(c, 65 if c.startswith("농협") else 70)
            self.tree.column(c, width=w, anchor=anchor, stretch=False)

    def _rebuild_tree_columns_if_needed(self, df_cols):
        if list(df_cols) != list(self.tree_cols):
            self.tree_cols = list(df_cols)
            self._apply_tree_columns()

    def _refresh_preview(self):
        if self.last_result is not None:
            self._populate_tree(self.last_result)

    def _filter_df(self, df):
        if self.matched_only_var.get():
            return df[df["매칭수량"].fillna(0) > 0].reset_index(drop=True)
        return df

    def _populate_tree(self, df):
        self._rebuild_tree_columns_if_needed(df.columns)
        df = self._filter_df(df)
        for row in self.tree.get_children():
            self.tree.delete(row)
        for _, r in df.iterrows():
            values = []
            for c in self.tree_cols:
                v = r[c]
                if v is None or (isinstance(v, float) and v != v):  # NaN
                    values.append("")
                elif c in ("잔존율", "남은율") and isinstance(v, (int, float)):
                    values.append(f"{v:.4f}")
                else:
                    values.append(str(v))
            verdict = r.get("판정") or ""
            matched = r.get("매칭수량") or 0
            tag = "matched" if matched else ("ok" if verdict == "OK" else ("ng" if verdict == "NG" else ""))
            self.tree.insert("", "end", values=values, tags=(tag,) if tag else ())

    def _pick_stock(self):
        init = self.settings.get("last_stock_dir") or ""
        p = filedialog.askopenfilename(
            title="재고 파일 선택", initialdir=init,
            filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.stock_var.set(p)
            self.settings["last_stock_dir"] = str(Path(p).parent)

    def _pick_orders(self):
        init = self.settings.get("last_orders_dir") or ""
        p = filedialog.askopenfilename(
            title="주문 파일 선택", initialdir=init,
            filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.orders_var.set(p)
            self.settings["last_orders_dir"] = str(Path(p).parent)

    def _pick_out(self):
        p = filedialog.askdirectory(title="출력 폴더 선택")
        if p:
            self.out_var.set(p)

    def _open_out(self):
        out = self.out_var.get() or str(Path(self.stock_var.get()).parent)
        if os.path.exists(out):
            os.startfile(out)

    def _log(self, msg: str):
        self.log_box.insert(END, msg + "\n")
        self.log_box.see(END)
        self.root.update_idletasks()

    def _run(self):
        stock = self.stock_var.get().strip()
        orders = self.orders_var.get().strip()
        if not stock or not os.path.exists(stock):
            messagebox.showerror("오류", "재고 파일을 선택하세요."); return
        if not orders or not os.path.exists(orders):
            messagebox.showerror("오류", "주문 파일을 선택하세요."); return
        self.run_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.last_result = None
        threading.Thread(target=self._run_match,
                         args=(stock, orders), daemon=True).start()

    def _run_match(self, stock_path: str, orders_path: str):
        try:
            self._log("재고 파일 로드 중...")
            stocks = load_locked_stock(stock_path, self.settings["ov_locations"])
            self._log(f"  → Lock 재고 {len(stocks)}건")

            self._log("주문 파일 로드 중...")
            orders = load_orders(orders_path)
            self._log(f"  → 주문 {len(orders)}행")

            self._log("기준정보 로드 중...")
            shelf_map = load_shelf_life_months(str(MASTER_CACHE))
            fs_ms = load_fs_ms_items(str(FS_MS_PATH))
            self._prune_auto_lots(stocks)  # 오늘 재고에 없는 자동 lot 정리
            manual_lots = load_lot_assignments(str(LOT_PATH))
            auto_lots = load_lot_assignments(str(LOT_AUTO_PATH))
            lots = {**auto_lots, **manual_lots}  # 수동이 자동보다 우선
            self._log(f"  → 유통기한 {len(shelf_map)}품목 / FS·MS {len(fs_ms)}품목 "
                      f"/ 로트 수동 {len(manual_lots)} + 자동 {len(auto_lots)}건")

            excluded = self.settings.get("excluded_nh_keywords") or []
            if excluded:
                self._log(f"  → 농협 제외 키워드: {', '.join(excluded)}")
            self._log("매칭 계산 중...")
            result = match(
                stocks, orders,
                shelf_life_map=shelf_map, fs_ms_items=fs_ms,
                lot_assignments=lots,
                today=date.today(),
                threshold=float(self.settings["threshold"]),
                nh_branches=self.settings["nh_branch_keywords"],
                excluded_nh_keywords=excluded,
            )

            self.last_result = result
            self._populate_tree(result)

            ok = sum(1 for v in result["판정"] if v == "OK")
            ng = sum(1 for v in result["판정"] if v == "NG")
            matched = sum(1 for v in result["매칭수량"].fillna(0) if v)
            summary = f"총 {len(result)}건 · OK {ok} · NG {ng} · 매칭 {matched}"
            self.summary_var.set(summary)
            self._log(f"\n✅ 미리보기 완료 — {summary}")
            self._log("   결과를 검토한 뒤 [② Excel로 저장] 버튼을 누르세요.")
            self.save_btn.config(state="normal")
        except Exception as e:
            import traceback; traceback.print_exc()
            self._log(f"\n❌ 오류: {e}")
            messagebox.showerror("오류", str(e))
        finally:
            self.run_btn.config(state="normal")

    def _save_excel(self):
        if self.last_result is None or self.last_result.empty:
            messagebox.showwarning("알림", "먼저 [매칭 실행]을 눌러 결과를 만들어 주세요.")
            return
        df_to_save = self._filter_df(self.last_result)
        if df_to_save.empty:
            messagebox.showwarning("알림", "저장할 행이 없습니다.")
            return
        stock_path = self.stock_var.get().strip()
        out_dir = self.out_var.get().strip() or (
            str(Path(stock_path).parent) if stock_path else str(ROOT))
        try:
            self._auto_append_lots(df_to_save)
            prefix = ("지정출고_매칭결과_매칭만"
                      if self.matched_only_var.get() else "지정출고_매칭결과")
            path = write_summary(df_to_save, out_dir, prefix=prefix)
            self.settings["last_output_dir"] = out_dir
            save_settings(self.settings)
            self._log(f"\n💾 저장 완료: {path}")
            self.open_btn.config(state="normal")
            if messagebox.askyesno("저장 완료", f"저장됐습니다.\n\n{path}\n\n파일을 열어보시겠어요?"):
                os.startfile(path)
        except Exception as e:
            self._log(f"\n❌ 저장 오류: {e}")
            messagebox.showerror("저장 오류", str(e))

    # ---------- 설정 탭 ----------
    def _build_settings_tab(self):
        f = self.settings_tab
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="잔존율 기준 (예: 0.70 = 70%)", font=("Segoe UI", 9, "bold")
                  ).grid(row=0, column=0, sticky=W, padx=8, pady=(12, 4))
        self.thr_var = DoubleVar(value=self.settings["threshold"])
        ttk.Entry(f, textvariable=self.thr_var, width=10).grid(row=0, column=1, sticky=W)

        ttk.Label(f, text="OV 로케이션 (Lock 재고 대상)",
                  font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky=W, padx=8, pady=(12, 4))
        loc_frame = ttk.Frame(f)
        loc_frame.grid(row=1, column=1, sticky=W)
        self.loc_vars = {}
        for i, loc in enumerate(["OV1", "OV4", "OV5", "OV6"]):
            v = BooleanVar(value=loc in self.settings["ov_locations"])
            self.loc_vars[loc] = v
            ttk.Checkbutton(loc_frame, text=loc, variable=v).pack(side="left", padx=4)

        ttk.Label(f, text="기준정보 캐시 (유통기한월)",
                  font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, sticky=W, padx=8, pady=(16, 4))
        cache_frame = ttk.Frame(f)
        cache_frame.grid(row=2, column=1, sticky=W + E, padx=(0, 8))
        cache_frame.columnconfigure(0, weight=1)
        self.master_status_var = StringVar()
        self._refresh_master_status()
        ttk.Label(cache_frame, textvariable=self.master_status_var,
                  foreground="#444").grid(row=0, column=0, sticky=W)
        ttk.Button(cache_frame, text="기준정보 xlsx 업로드...",
                   command=self._upload_master).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(cache_frame, text="템플릿 다운로드",
                   command=self._master_template).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(f, text="FS/MS 품목 리스트",
                  font=("Segoe UI", 9, "bold")).grid(
            row=3, column=0, sticky=W, padx=8, pady=(16, 4))
        fsms_frame = ttk.Frame(f)
        fsms_frame.grid(row=3, column=1, sticky=W + E, padx=(0, 8))
        fsms_frame.columnconfigure(0, weight=1)
        self.fsms_status_var = StringVar()
        self._refresh_fsms_status()
        ttk.Label(fsms_frame, textvariable=self.fsms_status_var,
                  foreground="#444").grid(row=0, column=0, sticky=W)
        ttk.Button(fsms_frame, text="FS/MS xlsx 업로드...",
                   command=self._upload_fsms).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(fsms_frame, text="템플릿 다운로드",
                   command=self._fsms_template).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(f, text="로트 지정 거래처 (품목+유통기한 → 카테고리)",
                  font=("Segoe UI", 9, "bold")).grid(
            row=34, column=0, sticky=W, padx=8, pady=(16, 4))
        lot_frame = ttk.Frame(f)
        lot_frame.grid(row=34, column=1, sticky=W + E, padx=(0, 8))
        lot_frame.columnconfigure(0, weight=1)
        self.lot_status_var = StringVar()
        self._refresh_lot_status()
        ttk.Label(lot_frame, textvariable=self.lot_status_var,
                  foreground="#444").grid(row=0, column=0, sticky=W)
        ttk.Button(lot_frame, text="로트 지정 xlsx 업로드...",
                   command=self._upload_lot).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(lot_frame, text="템플릿 다운로드",
                   command=self._lot_template).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(f, text="농협 매칭 제외 키워드", font=("Segoe UI", 9, "bold")).grid(
            row=35, column=0, sticky=W, padx=8, pady=(16, 4))
        excl_frame = ttk.Frame(f)
        excl_frame.grid(row=35, column=1, sticky=W + E, padx=(0, 8))
        self.excluded_var = StringVar(
            value=", ".join(self.settings.get("excluded_nh_keywords") or []))
        ttk.Entry(excl_frame, textvariable=self.excluded_var, width=40).pack(side="left")
        ttk.Label(excl_frame, text="(쉼표로 구분, 예: 제주, 강원)",
                  foreground="#888").pack(side="left", padx=(8, 0))

        ttk.Label(f, text="농협 거래처 키워드 매핑",
                  font=("Segoe UI", 9, "bold")).grid(
            row=4, column=0, sticky=N + W, padx=8, pady=(16, 4))
        nh_frame = ttk.Frame(f)
        nh_frame.grid(row=4, column=1, sticky=W + E, padx=(0, 8))
        self.nh_entries: dict[str, tuple[StringVar, StringVar]] = {}
        ttk.Label(nh_frame, text="키워드", foreground="#666").grid(row=0, column=0, padx=4)
        ttk.Label(nh_frame, text="→ 컬럼명", foreground="#666").grid(row=0, column=1, padx=4)
        for i, (k, v) in enumerate(self.settings["nh_branch_keywords"].items(), start=1):
            k_var, v_var = StringVar(value=k), StringVar(value=v)
            ttk.Entry(nh_frame, textvariable=k_var, width=14).grid(row=i, column=0, padx=4, pady=2)
            ttk.Entry(nh_frame, textvariable=v_var, width=16).grid(row=i, column=1, padx=4, pady=2)
            self.nh_entries[k] = (k_var, v_var)

        ttk.Button(f, text="설정 저장", command=self._save_settings).grid(
            row=99, column=0, sticky=W, padx=8, pady=20)

    def _refresh_master_status(self):
        if MASTER_CACHE.exists():
            count = len(load_shelf_life_months(str(MASTER_CACHE)))
            self.master_status_var.set(f"✅ 등록됨 ({count}품목)  |  {MASTER_CACHE.name}")
        else:
            self.master_status_var.set("⚠ 기준정보 없음 — 24개월 기본값 사용")

    def _prune_auto_lots(self, stocks):
        """매칭 실행 직전: 오늘 재고에 없는 자동 lot 항목을 제거.
        수동 등록(LOT_PATH)은 절대 건드리지 않음."""
        import pandas as pd
        if not LOT_AUTO_PATH.exists():
            return
        auto = load_lot_assignments(str(LOT_AUTO_PATH))
        if not auto:
            return
        # 오늘 재고의 (item_code, 유통기한) 집합
        stock_keys = set()
        for _, s in stocks.iterrows():
            try:
                code = str(int(s["제품코드"]))
                ymd_val = s.get("소비기한")
                ymd = int(ymd_val) if pd.notna(ymd_val) else None
            except (ValueError, TypeError):
                continue
            stock_keys.add((code, ymd))
        kept = {k: v for k, v in auto.items() if k in stock_keys}
        removed = len(auto) - len(kept)
        if removed == 0:
            return
        if kept:
            rows = [{"Item code": int(c), "유통기한": y if y else "", "카테고리": cat}
                    for (c, y), cat in kept.items()]
            pd.DataFrame(rows).to_excel(str(LOT_AUTO_PATH), index=False,
                                         sheet_name="로트지정_자동")
        else:
            LOT_AUTO_PATH.unlink(missing_ok=True)
        self._log(f"   🧹 자동 로트 정리: {removed}건 제거 (오늘 재고에 없음)")

    def _auto_append_lots(self, df):
        """매칭된 행의 (품목+유통기한 → 카테고리)를 lot_assignments_auto.xlsx에 누적.
        수동 등록과 자동 등록 모두에 이미 있으면 덮어쓰지 않음."""
        import pandas as pd
        from core.matcher import determine_lot_category
        nh_cols = [c for c in df.columns if c.startswith("농협")]
        manual = load_lot_assignments(str(LOT_PATH))
        auto = load_lot_assignments(str(LOT_AUTO_PATH))

        new_rows = []
        seen = set(manual.keys()) | set(auto.keys())
        for _, r in df.iterrows():
            qty = r.get("매칭수량")
            if qty is None or pd.isna(qty) or qty == 0:
                continue
            try:
                item_id = str(int(r["Item ID"]))
                ymd_int = int(r["유통기한"]) if pd.notna(r.get("유통기한")) else None
            except (ValueError, TypeError):
                continue
            if (item_id, ymd_int) in seen:
                continue
            cat = determine_lot_category(r.to_dict(), nh_cols)
            if not cat:
                continue
            seen.add((item_id, ymd_int))
            new_rows.append({
                "Item code": int(item_id),
                "Item": r.get("Item"),
                "유통기한": ymd_int if ymd_int else "",
                "카테고리": cat,
            })

        if not new_rows:
            return
        if LOT_AUTO_PATH.exists():
            try:
                existing_df = pd.read_excel(str(LOT_AUTO_PATH))
            except Exception:
                existing_df = pd.DataFrame(columns=["Item code", "Item", "유통기한", "카테고리"])
        else:
            existing_df = pd.DataFrame(columns=["Item code", "Item", "유통기한", "카테고리"])
        merged = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)
        CONFIG_DIR.mkdir(exist_ok=True)
        merged.to_excel(str(LOT_AUTO_PATH), index=False, sheet_name="로트지정_자동")
        if hasattr(self, "lot_status_var"):
            self._refresh_lot_status()
        self._log(f"   📌 자동 로트 누적: {len(new_rows)}건 → {LOT_AUTO_PATH.name}")

    def _refresh_lot_status(self):
        manual = len(load_lot_assignments(str(LOT_PATH))) if LOT_PATH.exists() else 0
        auto = len(load_lot_assignments(str(LOT_AUTO_PATH))) if LOT_AUTO_PATH.exists() else 0
        if manual or auto:
            self.lot_status_var.set(f"✅ 수동 {manual}건 / 자동 {auto}건 (저장 시 누적, 재고 사라지면 자동 제거)")
        else:
            self.lot_status_var.set("(등록된 로트 지정 없음 — 기존 매칭 룰대로 동작)")

    def _upload_lot(self):
        p = filedialog.askopenfilename(
            title="로트 지정 xlsx 선택 (Item code · 카테고리 컬럼 필수)",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        try:
            m = load_lot_assignments(p)
            if not m:
                messagebox.showerror(
                    "오류",
                    "'Item code' · '카테고리' 컬럼이 있고 유효 값(농협/대리점/FS/MS/급식)이 "
                    "있는 시트를 찾지 못했습니다.")
                return
            import shutil
            CONFIG_DIR.mkdir(exist_ok=True)
            shutil.copy(p, LOT_PATH)
            self._refresh_lot_status()
            messagebox.showinfo("완료", f"로트 지정 {len(m)}건 등록 완료.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _lot_template(self):
        p = filedialog.asksaveasfilename(
            title="로트 지정 템플릿 저장 위치",
            defaultextension=".xlsx",
            initialfile="로트지정_템플릿.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not p: return
        import pandas as pd
        pd.DataFrame({
            "Item code": [1010422, 2061502, 1015028, 1014854, 2032260],
            "Item": ["진간장 금S 15L", "폰타나 라치오 로마노치즈 1kg(학교급식)",
                     "맛간장 조림볶음용 1.7L", "유산균발효양조간장 15L(학교급식)",
                     "쓱쓱싹싹 양념깻잎"],
            "유통기한": [20280309, 20270917, "", 20280116, ""],
            "카테고리": ["농협", "MS", "대리점", "FS", "급식"],
        }).to_excel(p, index=False, sheet_name="로트지정")
        messagebox.showinfo(
            "저장 완료",
            f"{p}\n\n"
            "【필수 컬럼】 Item code · 카테고리\n"
            "【선택 컬럼】 유통기한(YYYYMMDD) — 비우면 그 품목 전체에 적용\n\n"
            "【카테고리 5종】\n"
            "  • 농협   → [NH]농협경제지주(매핑된 6개 지점)\n"
            "  • 대리점 → Customer code가 GT로 시작\n"
            "  • FS    → 거래처명에 [FS]\n"
            "  • MS    → 거래처명에 [MS]\n"
            "  • 급식   → 거래처명에 '급식' 포함\n\n"
            "지정된 재고는 해당 카테고리 주문에만 매칭되고 나머지는 모두 차단됩니다.\n"
            "수정 후 [로트 지정 xlsx 업로드] 버튼으로 등록하세요.")

    def _refresh_fsms_status(self):
        if FS_MS_PATH.exists():
            count = len(load_fs_ms_items(str(FS_MS_PATH)))
            self.fsms_status_var.set(f"✅ 등록됨 ({count}품목)  |  {FS_MS_PATH.name}")
        else:
            self.fsms_status_var.set("(등록된 FS/MS 품목 없음)")

    def _upload_master(self):
        p = filedialog.askopenfilename(
            title="기준정보 xlsx 선택 (Item code · 유통기한(월) 컬럼 포함)",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        try:
            m = load_shelf_life_months(p)
            if not m:
                messagebox.showerror("오류",
                    "'Item code' / '유통기한(월)' 컬럼을 가진 시트를 찾지 못했습니다.")
                return
            import shutil
            CONFIG_DIR.mkdir(exist_ok=True)
            shutil.copy(p, MASTER_CACHE)
            self._refresh_master_status()
            messagebox.showinfo("완료", f"기준정보 {len(m)}품목 등록 완료.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _upload_fsms(self):
        p = filedialog.askopenfilename(
            title="FS/MS 품목 xlsx 선택 (Item code · 구분 컬럼 포함)",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        try:
            m = load_fs_ms_items(p)
            if not m:
                messagebox.showerror("오류",
                    "'Item code' / '구분' 컬럼을 가진 시트를 찾지 못했습니다.")
                return
            import shutil
            CONFIG_DIR.mkdir(exist_ok=True)
            shutil.copy(p, FS_MS_PATH)
            self._refresh_fsms_status()
            messagebox.showinfo("완료", f"FS/MS 품목 {len(m)}건 등록 완료.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _master_template(self):
        p = filedialog.asksaveasfilename(
            title="기준정보 템플릿 저장 위치",
            defaultextension=".xlsx",
            initialfile="기준정보_템플릿.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not p: return
        import pandas as pd
        pd.DataFrame({
            "Item code": [1010122, 1010422, 2061502],
            "Item": ["양조간장701 500ml(R_09)", "진간장 금S 15L(R_12)_용기변경",
                     "폰타나 라치오 로마노치즈 1kg (학교급식용)"],
            "유통기한(월)": [24, 24, 18],
        }).to_excel(p, index=False, sheet_name="기준정보")
        messagebox.showinfo(
            "저장 완료",
            f"{p}\n\n[Item code · 유통기한(월)] 두 컬럼이 필수입니다.\n"
            "Item 컬럼은 선택사항(보기용)입니다.\n\n"
            "이 파일에 품목을 추가/수정한 뒤\n"
            "[기준정보 xlsx 업로드] 버튼으로 등록하세요.")

    def _fsms_template(self):
        p = filedialog.asksaveasfilename(
            title="FS/MS 템플릿 저장 위치",
            defaultextension=".xlsx",
            initialfile="FS_MS_품목_템플릿.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not p: return
        import pandas as pd
        pd.DataFrame({
            "Item code": [2061502, 1012859, 1014854, 1014854],
            "Item": ["폰타나 라치오 로마노치즈 1kg (학교급식용)",
                     "양조간장501 14L(학교급식전용)",
                     "유산균발효로깔끔한양조간장 15L(학교급식전용)",
                     "유산균발효로깔끔한양조간장 15L(학교급식전용)"],
            "유통기한": [20270917, 20280104, 20280116, ""],
            "구분": ["MS", "FS", "FS", "FS"],
        }).to_excel(p, index=False, sheet_name="FS_MS")
        messagebox.showinfo(
            "저장 완료",
            f"{p}\n\n"
            "[Item code · 구분(FS/MS)] 필수, [유통기한] 선택입니다.\n"
            "  • 유통기한(YYYYMMDD)을 적으면 그 소비기한 재고만 매칭\n"
            "  • 유통기한 비우면 해당 품목코드 전체에 적용\n\n"
            "예: 같은 품목이라도 20270917 재고만 MS로 보내고,\n"
            "    20280104 재고는 FS로 보낼 수 있습니다.\n\n"
            "수정 후 [FS/MS xlsx 업로드] 버튼으로 등록하세요.")

    def _save_settings(self):
        try:
            thr = float(self.thr_var.get())
        except Exception:
            messagebox.showerror("오류", "잔존율 기준은 숫자여야 합니다 (예: 0.70)")
            return
        self.settings["threshold"] = thr
        self.settings["ov_locations"] = [
            loc for loc, v in self.loc_vars.items() if v.get()]
        if not self.settings["ov_locations"]:
            messagebox.showerror("오류", "OV 로케이션을 최소 1개 선택하세요.")
            return
        nh = {}
        for (k_var, v_var) in self.nh_entries.values():
            k = k_var.get().strip()
            v = v_var.get().strip()
            if k and v:
                nh[k] = v
        self.settings["nh_branch_keywords"] = nh
        self.settings["excluded_nh_keywords"] = [
            x.strip() for x in self.excluded_var.get().split(",") if x.strip()]
        save_settings(self.settings)
        messagebox.showinfo("저장됨", "설정을 저장했습니다.")


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
