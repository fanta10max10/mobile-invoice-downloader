"""キャリアタブウィジェット（SB/YM ・ au/UQ ・ docomo）

回線管理スプシから取得した電話番号一覧をチェックボックスで表示し、
PDFの種類コンボボックスで選択できる。

認証情報シートへの保存機能を内包する。
"""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QVBoxLayout, QWidget,
)

from .styles import (
    ACCENT, ACCENT_SOFT, ACCENT_CHIP, ACCENT_DARK,
    BG_SURFACE, BG_HOVER, BG_CARD, BG_MAIN, BG_ELEVATED,
    BORDER, BORDER_LIGHT, TEXT, TEXT_DIM, TEXT_MUTED, TEXT_FAINT,
    GREEN, AMBER, RED,
)
from .workers import get_default_pdf_type, get_pdf_types_for_carrier

# タブ定義: (ラベル, キャリアファミリー, 属するキャリア名リスト)
TAB_DEFS = [
    ("SoftBank / Y!mobile", "softbank", ["SoftBank", "Ymobile"]),
    ("au / UQ mobile",       "au",      ["au", "UQmobile"]),
    ("docomo",               "docomo",  ["docomo"]),
]

CARRIER_ICONS = {
    "SoftBank": "SB",
    "Ymobile":  "YM",
    "au":       "au",
    "UQmobile": "UQ",
    "docomo":   "dc",
}

# キャリアバッジカラー（背景 / テキスト）
CARRIER_COLORS = {
    "SoftBank": ("#1a1a30", "#a5b4fc"),
    "Ymobile":  ("#1a2030", "#93c5fd"),
    "au":       ("#1a2a1a", "#6ee7b7"),
    "UQmobile": ("#1a2820", "#34d399"),
    "docomo":   ("#2a1a1a", "#fca5a5"),
}

# キャリアファミリー → 実行ボタンテキスト
RUN_LABEL = {
    "softbank": "▶  SB / YM を実行",
    "au":       "▶  au / UQ を実行",
    "docomo":   "▶  docomo を実行",
}


class _PhoneRow(QWidget):
    """電話番号1行ウィジェット（チェックボックス + PDFの種類コンボ付き）。"""

    changed = Signal()  # チェック状態またはPDF種類が変わったとき

    def __init__(self, carrier: str, info: dict, docomo_rep: str, parent=None):
        super().__init__(parent)
        self._carrier    = carrier
        self._phone      = info.get("phone", "")
        self._cancelled  = info.get("cancelled", False)
        self._docomo_rep = docomo_rep
        self._is_rep     = (carrier == "docomo" and self._phone == docomo_rep)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(10)

        # チェックボックス
        self._cb = QCheckBox()
        self._cb.setChecked(not self._cancelled)
        self._cb.stateChanged.connect(lambda _: self.changed.emit())
        layout.addWidget(self._cb)

        # キャリアバッジ（アルファベット略称）
        badge_bg, badge_fg = CARRIER_COLORS.get(carrier, ("#1a1a30", "#a5b4fc"))
        icon = CARRIER_ICONS.get(carrier, "?")
        c_badge = QLabel(icon)
        c_badge.setAlignment(Qt.AlignCenter)
        c_badge.setFixedSize(30, 20)
        c_badge.setStyleSheet(f"""
            background: {badge_bg};
            color: {badge_fg};
            border-radius: 5px;
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 0.5px;
        """)
        layout.addWidget(c_badge)

        # 電話番号（モノスペース）
        phone_disp = _format_phone(self._phone)
        phone_color = TEXT_MUTED if self._cancelled else TEXT
        deco = "text-decoration: line-through;" if self._cancelled else ""
        self._phone_label = QLabel(phone_disp)
        self._phone_label.setStyleSheet(f"""
            color: {phone_color};
            {deco}
            font-size: 13px;
            font-family: "Menlo", "Monaco";
            font-weight: 500;
        """)
        layout.addWidget(self._phone_label)

        # 端末名
        device = info.get("device", "")
        if device:
            dev_label = QLabel(device)
            dev_label.setStyleSheet(f"""
                color: {TEXT_FAINT};
                font-size: 11px;
                padding: 1px 6px;
            """)
            layout.addWidget(dev_label)

        layout.addStretch()

        # docomo代表回線バッジ
        if self._is_rep:
            rep_badge = QLabel("代表")
            rep_badge.setStyleSheet(f"""
                background: {ACCENT_CHIP};
                color: #c4b5fd;
                border: 1px solid rgba(139,92,246,0.25);
                border-radius: 5px;
                padding: 2px 7px;
                font-size: 10px;
                font-weight: 600;
            """)
            layout.addWidget(rep_badge)

        # 解約済みバッジ
        if self._cancelled:
            badge = QLabel("解約済")
            badge.setStyleSheet("""
                background: rgba(100,116,139,0.10);
                color: #475569;
                border: 1px solid rgba(100,116,139,0.15);
                border-radius: 5px;
                padding: 2px 7px;
                font-size: 10px;
            """)
            layout.addWidget(badge)

        # PDFの種類コンボボックス
        self._combo = QComboBox()
        self._combo.setFixedWidth(186)
        pdf_types = get_pdf_types_for_carrier(carrier, self._phone, docomo_rep)
        for t in pdf_types:
            self._combo.addItem(t)

        self._combo.currentTextChanged.connect(lambda _: self.changed.emit())
        layout.addWidget(self._combo)

        # 行スタイル（hover時に左ボーダーアクセント）
        self.setAutoFillBackground(True)
        self._set_row_style(False)

    def enterEvent(self, event):
        self._set_row_style(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._set_row_style(False)
        super().leaveEvent(event)

    def _set_row_style(self, hovered: bool):
        if self._cancelled:
            bg = "rgba(0,0,0,0)"
            border_left = "rgba(100,116,139,0.15)"
        elif hovered:
            bg = "rgba(139,92,246,0.05)"
            border_left = ACCENT
        else:
            bg = "rgba(0,0,0,0)"
            border_left = "transparent"
        self.setStyleSheet(f"""
            QWidget {{
                background: {bg};
                border-left: 2px solid {border_left};
                border-top: none;
                border-right: none;
                border-bottom: 1px solid rgba(255,255,255,0.03);
                border-radius: 0px;
            }}
        """)

    # ── プロパティ ──

    @property
    def is_checked(self) -> bool:
        return self._cb.isChecked()

    @property
    def phone(self) -> str:
        return self._phone

    @property
    def pdf_type(self) -> str:
        return self._combo.currentText()

    @property
    def carrier(self) -> str:
        return self._carrier

    # ── 外部からの設定 ──

    def set_checked(self, checked: bool):
        self._cb.setChecked(checked)

    def set_pdf_type(self, pdf_type: str):
        idx = self._combo.findText(pdf_type)
        if idx >= 0:
            self._combo.setCurrentIndex(idx)

    def set_default_pdf_type(self):
        default = get_default_pdf_type(self._carrier, self._phone, self._docomo_rep)
        self.set_pdf_type(default)


def _format_phone(phone: str) -> str:
    """11桁の電話番号をハイフン付きに変換する。"""
    digits = phone.replace("-", "").replace(" ", "")
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    return phone


class CarrierTabs(QWidget):
    """3タブのキャリア選択ウィジェット。

    Signals:
        run_requested(list, list):
            (carriers: list[CarrierConfig], selected_phones: list[str])
        save_requested(dict, dict, str):
            (phones_data, selections, docomo_rep)
        save_and_run_requested(dict, dict, str, list):
            (phones_data, selections, docomo_rep, all_carrier_configs)
    """

    run_requested          = Signal(list, list)         # (carriers, selected_phones)
    save_requested         = Signal(dict, dict, str)    # (phones_data, selections, docomo_rep)
    save_and_run_requested = Signal(dict, dict, str, list)  # (phones_data, selections, docomo_rep, carriers)

    def __init__(self, parent=None):
        super().__init__(parent)

        # キャリア別の行リスト: carrier_name -> [_PhoneRow]
        self._phone_rows: dict = {}
        # 全キャリアのphonesデータ（PhoneManagerLoaderから受け取る）
        self._phones_data: dict = {}
        self._docomo_rep: str = ""
        self._current_tab = 0

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # タブバー（ピル型タブ）
        tab_bar = QWidget()
        tab_bar.setStyleSheet(f"""
            QWidget {{
                background: {BG_CARD};
                border-radius: 12px 12px 0 0;
                border: 1px solid rgba(255,255,255,0.06);
                border-bottom: none;
                padding: 6px 8px 0 8px;
            }}
        """)
        tab_bar_layout = QHBoxLayout(tab_bar)
        tab_bar_layout.setContentsMargins(8, 6, 8, 0)
        tab_bar_layout.setSpacing(4)

        self._tab_btns: list = []
        for i, (label, _, _) in enumerate(TAB_DEFS):
            btn = QPushButton(label)
            btn.setObjectName("tabBtn")
            btn.setCheckable(False)
            btn.clicked.connect(lambda _, idx=i: self._switch_tab(idx))
            tab_bar_layout.addWidget(btn)
            self._tab_btns.append(btn)

        outer.addWidget(tab_bar)

        # タブコンテンツ（スタック）
        self._pages: list = []
        self._stack = QWidget()
        stack_layout = QVBoxLayout(self._stack)
        stack_layout.setContentsMargins(0, 0, 0, 0)
        stack_layout.setSpacing(0)

        for i, (_, family, carriers) in enumerate(TAB_DEFS):
            page = self._build_page(family, carriers)
            page.setVisible(i == 0)
            self._pages.append(page)
            stack_layout.addWidget(page)

        outer.addWidget(self._stack)

        self._switch_tab(0)

    def _build_page(self, family: str, carriers: list) -> QWidget:
        """タブ1ページ分のウィジェットを作る。"""
        page = QWidget()
        page.setStyleSheet(f"""
            QWidget {{
                background: {BG_CARD};
                border: 1px solid rgba(255,255,255,0.06);
                border-top: none;
                border-radius: 0 0 12px 12px;
            }}
        """)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 12)
        layout.setSpacing(0)

        # 上部ツールバー
        toolbar_wrapper = QWidget()
        toolbar_wrapper.setStyleSheet(f"""
            QWidget {{
                background: rgba(139,92,246,0.04);
                border-bottom: 1px solid rgba(255,255,255,0.04);
                border-top: none;
                border-left: none;
                border-right: none;
                border-radius: 0;
            }}
        """)
        toolbar = QHBoxLayout(toolbar_wrapper)
        toolbar.setContentsMargins(14, 8, 14, 8)
        toolbar.setSpacing(6)

        for label, checked in [("全選択", True), ("全解除", False)]:
            btn = QPushButton(label)
            btn.setFixedHeight(26)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    color: {TEXT_MUTED};
                    border: 1px solid rgba(255,255,255,0.08);
                    border-radius: 6px;
                    font-size: 11px;
                    padding: 0 10px;
                }}
                QPushButton:hover {{
                    background: rgba(255,255,255,0.04);
                    color: {TEXT_DIM};
                    border-color: rgba(255,255,255,0.14);
                }}
            """)
            btn.clicked.connect(lambda _, f=family, c=checked: self._select_all(f, c))
            toolbar.addWidget(btn)

        toolbar.addStretch()

        summary_label = QLabel("")
        summary_label.setObjectName(f"summary_{family}")
        summary_label.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 11px;
            font-family: "Menlo";
        """)
        toolbar.addWidget(summary_label)

        layout.addWidget(toolbar_wrapper)

        # スクロールエリア（電話番号リスト）
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFixedHeight(196)
        scroll.setStyleSheet(f"""
            QScrollArea {{
                background: transparent;
                border: none;
            }}
            QScrollArea > QWidget > QWidget {{
                background: transparent;
            }}
        """)

        list_widget = QWidget()
        list_widget.setStyleSheet("background: transparent;")
        list_layout = QVBoxLayout(list_widget)
        list_layout.setContentsMargins(0, 4, 0, 4)
        list_layout.setSpacing(0)

        loading = QLabel("  ⏳  読み込み中...")
        loading.setObjectName(f"loading_{family}")
        loading.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px; padding: 16px 14px;")
        list_layout.addWidget(loading)
        list_layout.addStretch()

        scroll.setWidget(list_widget)
        layout.addWidget(scroll)

        # ボタン行（保存 / 保存して実行 / 実行）
        btn_wrapper = QWidget()
        btn_wrapper.setStyleSheet("background: transparent; border: none;")
        btn_row = QHBoxLayout(btn_wrapper)
        btn_row.setContentsMargins(14, 8, 14, 0)
        btn_row.setSpacing(8)

        save_btn = QPushButton("💾  保存")
        save_btn.setObjectName("saveBtn")
        save_btn.setEnabled(False)
        save_btn.setFixedHeight(34)
        save_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {TEXT_MUTED};
                border: 1px solid rgba(255,255,255,0.10);
                border-radius: 8px;
                font-size: 12px;
                padding: 0 16px;
            }}
            QPushButton:hover {{
                background: rgba(255,255,255,0.04);
                color: {TEXT_DIM};
                border-color: rgba(255,255,255,0.18);
            }}
            QPushButton:disabled {{
                color: {TEXT_FAINT};
                border-color: rgba(255,255,255,0.05);
            }}
        """)
        save_btn.clicked.connect(lambda: self._on_save(family))
        btn_row.addWidget(save_btn)

        save_run_btn = QPushButton("💾▶  保存して実行")
        save_run_btn.setObjectName("saveRunBtn")
        save_run_btn.setEnabled(False)
        save_run_btn.setFixedHeight(34)
        save_run_btn.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT_CHIP};
                color: #c4b5fd;
                border: 1px solid rgba(139,92,246,0.28);
                border-radius: 8px;
                font-size: 12px;
                padding: 0 16px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background: rgba(139,92,246,0.22);
                border-color: rgba(139,92,246,0.45);
            }}
            QPushButton:disabled {{
                color: {TEXT_FAINT};
                background: transparent;
                border-color: rgba(255,255,255,0.05);
            }}
        """)
        save_run_btn.clicked.connect(lambda: self._on_save_and_run(family))
        btn_row.addWidget(save_run_btn)

        btn_row.addStretch()

        run_btn = QPushButton(RUN_LABEL[family])
        run_btn.setObjectName("runBtn")
        run_btn.setEnabled(False)
        run_btn.setFixedHeight(34)
        run_btn.clicked.connect(lambda: self._on_run(family))
        btn_row.addWidget(run_btn)

        layout.addWidget(btn_wrapper)

        page.setProperty("family", family)
        page.setProperty("list_layout", list_layout)
        page.setProperty("run_btn", run_btn)
        page.setProperty("save_btn", save_btn)
        page.setProperty("save_run_btn", save_run_btn)
        page.setProperty("loading_label", loading)
        page.setProperty("summary_label", summary_label)

        return page

    def _switch_tab(self, idx: int):
        self._current_tab = idx
        for i, (btn, page) in enumerate(zip(self._tab_btns, self._pages)):
            active = (i == idx)
            btn.setProperty("active", "true" if active else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            page.setVisible(active)

    # ─── データ読み込み ───

    def load_data(self, data: dict):
        """PhoneManagerLoader から受け取ったデータをセットする。

        data = {
            "phones": {carrier: [{"phone", "cancelled", "device", "name", "loginId"}]},
            "selections": {carrier: {phone: {"pdfType": str}}},
            "docomo_rep": str,
            "target_month": str,
        }
        """
        phones     = data.get("phones", {})
        selections = data.get("selections", {})
        docomo_rep = data.get("docomo_rep", "")

        self._phones_data = phones
        self._docomo_rep  = docomo_rep

        for i, (_, family, carriers) in enumerate(TAB_DEFS):
            page          = self._pages[i]
            list_layout   = page.property("list_layout")
            run_btn       = page.property("run_btn")
            save_btn      = page.property("save_btn")
            save_run_btn  = page.property("save_run_btn")
            loading_label = page.property("loading_label")
            summary_label = page.property("summary_label")

            # ローディングラベルを削除
            if loading_label:
                loading_label.setParent(None)

            # 既存行をクリア
            for carrier in carriers:
                self._phone_rows[carrier] = []

            # ストレッチ行を取り除く（既存のものを削除）
            _clear_stretch(list_layout)

            # 各キャリアの行を追加
            total_active = 0
            for carrier in carriers:
                phone_list = phones.get(carrier, [])
                carrier_sel = selections.get(carrier, {})

                for info in phone_list:
                    row = _PhoneRow(carrier, info, docomo_rep)
                    # 既存の選択状態を反映
                    phone = info.get("phone", "")
                    if phone in carrier_sel:
                        sel_info = carrier_sel[phone]
                        row.set_checked(True)
                        row.set_pdf_type(sel_info.get("pdfType", ""))
                    else:
                        row.set_checked(not info.get("cancelled", False))
                        row.set_default_pdf_type()

                    row.changed.connect(lambda f=family: self._update_summary(f))
                    self._phone_rows.setdefault(carrier, []).append(row)
                    list_layout.addWidget(row)
                    if not info.get("cancelled", False):
                        total_active += 1

            if not any(phones.get(c) for c in carriers):
                empty = QLabel("この認証グループの回線はありません")
                empty.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px; padding: 8px;")
                list_layout.addWidget(empty)

            list_layout.addStretch()

            run_btn.setEnabled(total_active > 0)
            save_btn.setEnabled(True)
            save_run_btn.setEnabled(total_active > 0)

            self._update_summary(family)

    def load_phones(self, phones_by_family: dict):
        """PhoneListLoader（旧API）との互換性のために残す。"""
        # 旧形式 {family: [{"phone", "carrier", "status", "device"}]} を新形式に変換
        carrier_map = {
            "softbank": ["SoftBank", "Ymobile"],
            "au":       ["au", "UQmobile"],
            "docomo":   ["docomo"],
        }
        phones_new = {k: [] for k in ("SoftBank", "Ymobile", "au", "UQmobile", "docomo")}
        for family, phone_list in phones_by_family.items():
            for item in phone_list:
                carrier = item.get("carrier", "")
                if carrier in phones_new:
                    phones_new[carrier].append({
                        "phone":     item.get("phone", ""),
                        "cancelled": item.get("status", "契約中") != "契約中",
                        "device":    item.get("device", ""),
                        "name":      "",
                        "loginId":   "",
                    })
        self.load_data({
            "phones":       phones_new,
            "selections":   {},
            "docomo_rep":   "",
            "target_month": "",
        })

    # ─── サマリー更新 ───

    def _update_summary(self, family: str):
        """選択サマリーラベルを更新する。"""
        carriers = next(c for _, f, c in TAB_DEFS if f == family)
        page = next(p for p in self._pages if p.property("family") == family)
        summary_label = page.property("summary_label")
        if summary_label is None:
            return

        parts = []
        for carrier in carriers:
            rows = self._phone_rows.get(carrier, [])
            checked = sum(1 for r in rows if r.is_checked)
            total   = len(rows)
            if total > 0:
                icon = CARRIER_ICONS.get(carrier, "📱")
                parts.append(f"{icon} {checked}/{total}")

        summary_label.setText("  ".join(parts) if parts else "")

    # ─── 全選択/全解除 ───

    def _select_all(self, family: str, checked: bool):
        carriers = next(c for _, f, c in TAB_DEFS if f == family)
        for carrier in carriers:
            for row in self._phone_rows.get(carrier, []):
                row.set_checked(checked)
        self._update_summary(family)

    # ─── 選択状態の収集 ───

    def _collect_selections(self, family: str) -> dict:
        """指定ファミリーのチェック済み回線を {carrier: {phone: {pdfType}}} 形式で返す。"""
        carriers = next(c for _, f, c in TAB_DEFS if f == family)
        selections = {}
        for carrier in carriers:
            sel = {}
            for row in self._phone_rows.get(carrier, []):
                if row.is_checked:
                    sel[row.phone] = {"pdfType": row.pdf_type}
            if sel:
                selections[carrier] = sel
        return selections

    def _collect_all_selections(self) -> dict:
        """全ファミリーのチェック済み回線を {carrier: {phone: {pdfType}}} で返す。"""
        result = {}
        for _, family, carriers in TAB_DEFS:
            for carrier in carriers:
                sel = {}
                for row in self._phone_rows.get(carrier, []):
                    if row.is_checked:
                        sel[row.phone] = {"pdfType": row.pdf_type}
                if sel:
                    result[carrier] = sel
        return result

    # ─── ボタンイベント ───

    def _on_save(self, family: str):
        """保存ボタン押下 → save_requested シグナルを発火。"""
        selections = self._collect_selections(family)
        self.save_requested.emit(self._phones_data, selections, self._docomo_rep)

    def _on_save_and_run(self, family: str):
        """保存して実行ボタン押下 → save_and_run_requested シグナルを発火。"""
        from download import (
            SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG, DOCOMO_CONFIG
        )
        family_carriers = {
            "softbank": [SOFTBANK_CONFIG, YMOBILE_CONFIG],
            "au":       [AU_CONFIG, UQ_CONFIG],
            "docomo":   [DOCOMO_CONFIG],
        }
        carrier_configs = family_carriers.get(family, [])
        selections = self._collect_selections(family)
        self.save_and_run_requested.emit(
            self._phones_data, selections, self._docomo_rep, carrier_configs
        )

    def _on_run(self, family: str):
        """実行ボタン押下 → run_requested シグナルを発火。"""
        from download import (
            SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG, DOCOMO_CONFIG
        )
        family_carriers = {
            "softbank": [SOFTBANK_CONFIG, YMOBILE_CONFIG],
            "au":       [AU_CONFIG, UQ_CONFIG],
            "docomo":   [DOCOMO_CONFIG],
        }
        carriers = family_carriers.get(family, [])
        rows_all = []
        for carrier in next(c for _, f, c in TAB_DEFS if f == family):
            rows_all.extend(self._phone_rows.get(carrier, []))

        all_phones = [r.phone for r in rows_all if r.phone]
        selected   = [r.phone for r in rows_all if r.is_checked and r.phone]

        # 全件チェックなら RETRY_PHONES を渡さない（全件実行）
        if set(selected) == set(all_phones):
            selected = []

        self.run_requested.emit(carriers, selected)

    def get_all_selected(self) -> tuple:
        """全タブの選択済み (carriers, phones) を返す（全キャリア一括実行用）。"""
        from download import (
            SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG, DOCOMO_CONFIG
        )
        all_configs = [SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG, DOCOMO_CONFIG]

        all_phones_total = []
        selected_total   = []
        for _, family, carriers in TAB_DEFS:
            for carrier in carriers:
                rows = self._phone_rows.get(carrier, [])
                for r in rows:
                    if r.phone:
                        all_phones_total.append(r.phone)
                    if r.is_checked and r.phone:
                        selected_total.append(r.phone)

        # 全件チェックなら RETRY_PHONES を渡さない
        if set(selected_total) == set(all_phones_total):
            selected_total = []

        return all_configs, selected_total

    def set_enabled(self, enabled: bool):
        """実行中はすべての操作ボタンを無効化する。"""
        for page in self._pages:
            for prop in ("run_btn", "save_btn", "save_run_btn"):
                btn = page.property(prop)
                if btn:
                    btn.setEnabled(enabled)


def _clear_stretch(layout: QVBoxLayout):
    """VBoxLayout からストレッチアイテムを取り除く。"""
    for i in reversed(range(layout.count())):
        item = layout.itemAt(i)
        if item and item.spacerItem():
            layout.takeAt(i)
