from PyQt6 import QtWidgets, QtCore
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QDateEdit, QTableWidget, QTableWidgetItem, QMessageBox,
    QCheckBox, QTabWidget, QScrollArea, QGridLayout, QFileDialog,
    QSizePolicy, QLineEdit, QHeaderView, QTableView, QAbstractItemView
)
from PyQt6.QtCore import QDate, QTimer, QEvent, Qt  # <- Qt ergänzt
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import datetime, os, math

from controller.main_controller import MainController   # <- LOGIC!
from util.paths import resource_path


class MainWindow(QMainWindow):
    def __init__(self, controller: MainController):
        super().__init__()
        self.ctrl = controller
        self.setWindowTitle("LabStats – Befund-Statistik")
        self.resize(1250, 800)
        self._apply_style()

        self._cols = 8  # Analyten-Gitter-Spalten

        tabs = QTabWidget()
        tabs.addTab(self._build_tab_counts(), "Zählungen")
        tabs.addTab(self._build_tab_open(), "Offene Anforderungen")
        tabs.addTab(self._build_tab_suspected(), "Nicht entnommen?")
        tabs.addTab(self._build_tab_singlets(), "Singlets")
        tabs.addTab(self._build_tab_settings(), "Einstellungen")
        self.setCentralWidget(tabs)

        try:
            self.ctrl.monthly_export_if_first()
        except Exception as ex:
            print("Monthly export failed:", ex)

    # ---------- Styling/Helpers
    def _apply_style(self):
        try:
            qss_path = resource_path("resources/style.qss")
            with open(qss_path, "r", encoding="utf-8") as f:
                self.setStyleSheet(f.read())
        except Exception:
            pass

    def _make_checkbox_grid(self, analytes, cols: int, *, show_all: bool = False, cap_ratio: float = 0.25):
        """
        show_all=True  -> ScrollArea passt sich der gesamten Inhaltshöhe an (keine Scrollbar, nichts abgeschnitten).
        show_all=False -> Höhe wird auf cap_ratio * Fensterhöhe begrenzt (Scrollbars möglich).
        """
        search_line = QLineEdit()
        search_line.setPlaceholderText("Analyt suchen …")
        search_line.setObjectName("analyteSearch")

        scroll = QScrollArea()
        scroll.setWidgetResizable(False)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        inner = QWidget()
        grid = QGridLayout(inner)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(0)
        grid.setVerticalSpacing(0)

        chk_list = []
        for i, a in enumerate(analytes):
            cb = QCheckBox(a)
            cb.setObjectName("analyteCheck")
            cb.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
            r, c = divmod(i, cols)
            grid.addWidget(cb, r, c)
            chk_list.append(cb)
        inner.setLayout(grid)
        scroll.setWidget(inner)

        # Scrollbar-Policy je nach Modus
        scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff if show_all else Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        wrapper = QWidget()
        wrapper.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        v = QVBoxLayout(wrapper)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)
        v.addWidget(search_line)
        v.addWidget(scroll)

        # Flags für Layout-Anpassung merken
        wrapper._scroll = scroll
        wrapper._search = search_line
        wrapper._checks = chk_list
        wrapper._cols = cols
        wrapper._show_all = show_all
        wrapper._cap_ratio = cap_ratio

        def _filter(t: str):
            t = t.strip().lower()
            for cb in chk_list:
                cb.setVisible((t in cb.text().lower()) if t else True)
            self._fit_checkbox_wrapper(wrapper)

        search_line.textChanged.connect(_filter)
        wrapper.installEventFilter(self)
        QTimer.singleShot(0, lambda: self._fit_checkbox_wrapper(wrapper))
        return wrapper, scroll, chk_list

    def eventFilter(self, obj: QtCore.QObject, ev: QtCore.QEvent) -> bool:
        if hasattr(obj, "_scroll") and ev.type() in (QEvent.Type.Show, QEvent.Type.ShowToParent, QEvent.Type.Resize):
            QTimer.singleShot(0, lambda o=obj: self._fit_checkbox_wrapper(o))
        return super().eventFilter(obj, ev)

    def _fit_checkbox_wrapper(self, wrapper: QWidget):
        if not hasattr(wrapper, "_scroll"):
            return
        scroll = wrapper._scroll
        search = wrapper._search
        checks = wrapper._checks
        cols = wrapper._cols
        show_all = getattr(wrapper, "_show_all", False)
        cap_ratio = getattr(wrapper, "_cap_ratio", 0.25)

        # sichtbare Elemente bestimmen
        visible = [cb for cb in checks if cb.isVisible()] or checks
        row_h = max(18, visible[0].sizeHint().height())

        # Layout/Abstände exakt berücksichtigen
        inner = scroll.widget()
        grid = inner.layout()
        m = grid.contentsMargins()
        vsp = max(0, grid.verticalSpacing())
        rows = (len(visible) + cols - 1) // cols

        grid_h = rows * row_h + max(0, rows - 1) * vsp + m.top() + m.bottom()
        frame = scroll.frameWidth() * 2
        fudge = 10  # kleiner Puffer
        desired_h = grid_h + frame + fudge

        if show_all:
            h = desired_h
        else:
            # auf Anteil der Fensterhöhe begrenzen
            limit = int(self.height() * cap_ratio)
            h = min(desired_h, max(60, limit))

        scroll.setFixedHeight(h)
        wrapper.setFixedHeight(search.sizeHint().height() + 4 + h)

    # ---------- Tab: Zählungen (hier: show_all=True -> nichts abgeschnitten)
    def _build_tab_counts(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        top = QGroupBox("Zeitraum & Filter")
        tl = QHBoxLayout(top)

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        today = datetime.date.today()
        first = today.replace(day=1)
        self.start_date.setDate(QDate(first.year, first.month, first.day))
        self.end_date.setDate(QDate(today.year, today.month, today.day))
        tl.addWidget(QLabel("Start:"))
        tl.addWidget(self.start_date)
        tl.addWidget(QLabel("Ende:"))
        tl.addWidget(self.end_date)

        self.status_only_open = QCheckBox("Nur offene (für Wochentage)")
        tl.addWidget(self.status_only_open)

        btn_reload = QPushButton("Analyten aktualisieren")
        btn_reload.clicked.connect(self._reload_analyte_controls)
        tl.addWidget(btn_reload)

        btn_run = QPushButton("Berechnen")
        btn_run.clicked.connect(self._run_counts)
        tl.addWidget(btn_run)
        tl.addStretch(1)

        layout.addWidget(top)

        analytes = self.ctrl.list_included_analytes() or self.ctrl.list_all_analytes()
        # >>> alle Analyten vollständig anzeigen (ohne Abschneiden)
        self.wrap_counts, self.scroll_counts, self.chk_analytes_counts = self._make_checkbox_grid(
            analytes, self._cols, show_all=True
        )
        layout.addWidget(self.wrap_counts)

        btn_all = QPushButton("Alle auswählen")
        def _sel_all():
            any_unchecked = any(not cb.isChecked() and cb.isVisible() for cb in self.chk_analytes_counts)
            for cb in self.chk_analytes_counts:
                if cb.isVisible():
                    cb.setChecked(any_unchecked)
        btn_all.clicked.connect(_sel_all)
        layout.addWidget(btn_all)

        self.model_counts = QStandardItemModel(0, 4, self)
        self.model_counts.setHorizontalHeaderLabels(["Kategorie", "Wert", "Hinweis", "Details"])
        self.view_counts = QTableView()
        self.view_counts.setModel(self.model_counts)
        hdr = self.view_counts.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        self.view_counts.setAlternatingRowColors(True)
        self.view_counts.setSortingEnabled(False)
        layout.addWidget(self.view_counts)
        return w

    def _run_counts(self):
        start = datetime.datetime(self.start_date.date().year(), self.start_date.date().month(), self.start_date.date().day())
        end   = datetime.datetime(self.end_date.date().year(),   self.end_date.date().month(),   self.end_date.date().day(), 23,59,59)
        analytes = [cb.text() for cb in self.chk_analytes_counts if cb.isChecked() and cb.isVisible()]
        try:
            rows = self.ctrl.build_counts_rows_multi(start, end, analytes, self.status_only_open.isChecked())
        except Exception as ex:
            QMessageBox.critical(self, "Fehler beim Berechnen", str(ex)); return

        new_model = QStandardItemModel(0, 4, self)
        new_model.setHorizontalHeaderLabels(["Kategorie", "Wert", "Hinweis", "Details"])
        for a,b,c,d in rows:
            items = [QStandardItem(str(x)) for x in (a,b,c,d)]
            for it in items: it.setEditable(False)
            new_model.appendRow(items)
        old = self.view_counts.model(); self.view_counts.setModel(None); self.view_counts.setModel(new_model)
        if old is not None and old is not new_model: old.deleteLater()
        self.model_counts = new_model

    # ---------- Tab: Offene Anforderungen
    def _build_tab_open(self) -> QWidget:
        w = QWidget(); layout = QVBoxLayout(w)
        line = QHBoxLayout()
        line.addWidget(QLabel("Offene Anforderungen ab Datum:"))
        self.since_date = QDateEdit(); self.since_date.setCalendarPopup(True)
        today = datetime.date.today(); first = today.replace(day=1)
        self.since_date.setDate(QDate(first.year, first.month, first.day))
        line.addWidget(self.since_date)
        self.btn_open_count = QPushButton("Zählen"); self.btn_open_count.clicked.connect(self._run_open); line.addWidget(self.btn_open_count)
        btn_reload = QPushButton("Analyten aktualisieren"); btn_reload.clicked.connect(self._reload_analyte_controls); line.addWidget(btn_reload)
        line.addStretch(1); layout.addLayout(line)

        analytes = self.ctrl.list_included_analytes() or self.ctrl.list_all_analytes()
        self.wrap_open, self.scroll_analytes, self.chk_analytes = self._make_checkbox_grid(analytes, self._cols)
        layout.addWidget(self.wrap_open)

        btn_row = QHBoxLayout()
        btn_all = QPushButton("Alle auswählen")
        def _select_all():
            any_unchecked = any(not cb.isChecked() and cb.isVisible() for cb in self.chk_analytes)
            for cb in self.chk_analytes:
                if cb.isVisible(): cb.setChecked(any_unchecked)
        btn_all.clicked.connect(_select_all)
        btn_row.addWidget(btn_all); btn_row.addStretch(1); layout.addLayout(btn_row)

        self.table_open = QTableWidget(0, 4)
        self.table_open.setHorizontalHeaderLabels(["Analyt (1)", "Offene (1)", "Analyt (2)", "Offene (2)"])
        self.table_open.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_open.setAlternatingRowColors(True); self.table_open.setSortingEnabled(False)
        layout.addWidget(self.table_open)
        return w

    def _run_open(self):
        analytes = [cb.text() for cb in self.chk_analytes if cb.isChecked() and cb.isVisible()]
        if not analytes:
            QMessageBox.warning(self, "Hinweis", "Bitte mindestens einen Analyt auswählen."); return
        since = datetime.datetime(self.since_date.date().year(), self.since_date.date().month(), self.since_date.date().day())
        rows = self.ctrl.build_open_counts_since(analytes, since)

        self.table_open.setUpdatesEnabled(False)
        try:
            n_pairs = math.ceil(len(rows)/2); self.table_open.clearContents(); self.table_open.setRowCount(n_pairs)
            for i in range(n_pairs):
                code1, cnt1 = rows[2*i]; self.table_open.setItem(i, 0, QTableWidgetItem(code1)); self.table_open.setItem(i, 1, QTableWidgetItem(str(cnt1)))
                if 2*i+1 < len(rows):
                    code2, cnt2 = rows[2*i+1]; self.table_open.setItem(i, 2, QTableWidgetItem(code2)); self.table_open.setItem(i, 3, QTableWidgetItem(str(cnt2)))
                else:
                    self.table_open.setItem(i, 2, QTableWidgetItem("")); self.table_open.setItem(i, 3, QTableWidgetItem(""))
        finally:
            self.table_open.setUpdatesEnabled(True)

    # ---------- Tab: Nicht entnommen?
    def _build_tab_suspected(self) -> QWidget:
        w = QWidget(); layout = QVBoxLayout(w)

        btn_refresh = QPushButton("Liste aktualisieren")
        btn_refresh.clicked.connect(self._refresh_suspected)
        layout.addWidget(btn_refresh)

        self.table_susp = QTableWidget(0, 4)
        self.table_susp.setHorizontalHeaderLabels(["ProbenNr", "Order-Zeit", "Anzahl Anforderungen", "Analyte"])
        hdr = self.table_susp.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        self.table_susp.setAlternatingRowColors(True)
        self.table_susp.setSortingEnabled(False)
        self.table_susp.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_susp.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout.addWidget(self.table_susp)

        btn_delete = QPushButton("Ausgewählte Probe(n) löschen (mit Audit)")
        btn_delete.clicked.connect(self._delete_selected_samples)
        layout.addWidget(btn_delete)
        return w

    def _refresh_suspected(self):
        try:
            rows = self.ctrl.suspected_missing_blood_draw()
        except Exception as ex:
            QMessageBox.critical(self, "Fehler", str(ex)); return
        self.table_susp.setUpdatesEnabled(False)
        try:
            self.table_susp.clearContents(); self.table_susp.setRowCount(len(rows))
            for r, row in enumerate(rows):
                self.table_susp.setItem(r, 0, QTableWidgetItem(str(row.get("ProbenNr",""))))
                self.table_susp.setItem(r, 1, QTableWidgetItem(str(row.get("OrderTime",""))))
                self.table_susp.setItem(r, 2, QTableWidgetItem(str(row.get("NumReq",""))))
                self.table_susp.setItem(r, 3, QTableWidgetItem(str(row.get("Analytes",""))))
        finally:
            self.table_susp.setUpdatesEnabled(True)

    def _delete_selected_samples(self):
        sel = self.table_susp.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst mindestens eine Zeile auswählen.")
            return

        proben = []
        for idx in sel:
            item = self.table_susp.item(idx.row(), 0)
            if item:
                p = (item.text() or "").strip()
                if p:
                    proben.append(p)

        if not proben:
            return

        if QMessageBox.question(
                self, "Löschung bestätigen",
                f"{len(proben)} Probe(n) wirklich löschen? Diese Aktion ist irreversibel."
        ) != QMessageBox.StandardButton.Yes:
            return

        try:
            deleted = self.ctrl.delete_samples_with_audit(proben)
            QMessageBox.information(self, "Gelöscht", f"Gelöschte DB-Zeilen: {deleted}")
            self._refresh_suspected()
        except Exception as ex:
            QMessageBox.critical(self, "Fehler", str(ex))

    # ---------- Tab: Singlets
    def _build_tab_singlets(self) -> QWidget:
        w = QWidget(); layout = QVBoxLayout(w)
        line = QHBoxLayout(); line.addWidget(QLabel("Analyse ab Datum:"))
        self.sing_since = QDateEdit(); self.sing_since.setCalendarPopup(True)
        today = datetime.date.today(); first = today.replace(day=1)
        self.sing_since.setDate(QDate(first.year, first.month, first.day))
        line.addWidget(self.sing_since); btn = QPushButton("Analysieren"); btn.clicked.connect(self._run_singlets); line.addWidget(btn); line.addStretch(1)
        layout.addLayout(line)

        self.tbl_sing = QTableWidget(0,2); self.tbl_sing.setHorizontalHeaderLabels(["Analyt (Singlet)","Anzahl"])
        hdr1 = self.tbl_sing.horizontalHeader(); hdr1.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); hdr1.setStretchLastSection(True)
        self.tbl_sing.setAlternatingRowColors(True); self.tbl_sing.setSortingEnabled(False)
        layout.addWidget(QLabel("Einzelne offene Analyten (Singlets)")); layout.addWidget(self.tbl_sing)

        self.tbl_pairs = QTableWidget(0,2); self.tbl_pairs.setHorizontalHeaderLabels(["Analyt-Kombination (2)","Anzahl"])
        hdr2 = self.tbl_pairs.horizontalHeader(); hdr2.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); hdr2.setStretchLastSection(True)
        self.tbl_pairs.setAlternatingRowColors(True); self.tbl_pairs.setSortingEnabled(False)
        layout.addWidget(QLabel("Häufige 2er-Kombinationen offener Analyten")); layout.addWidget(self.tbl_pairs)

        self.tbl_trips = QTableWidget(0,2); self.tbl_trips.setHorizontalHeaderLabels(["Analyt-Kombination (3)","Anzahl"])
        hdr3 = self.tbl_trips.horizontalHeader(); hdr3.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); hdr3.setStretchLastSection(True)
        self.tbl_trips.setAlternatingRowColors(True); self.tbl_trips.setSortingEnabled(False)
        layout.addWidget(QLabel("Häufige 3er-Kombinationen offener Analyten")); layout.addWidget(self.tbl_trips)
        return w

    def _run_singlets(self):
        since = datetime.datetime(self.sing_since.date().year(), self.sing_since.date().month(), self.sing_since.date().day())
        sing, pairs, trips = self.ctrl.combo_stats_since(since)
        def fill(tbl, rows):
            tbl.setUpdatesEnabled(False)
            try:
                tbl.clearContents(); tbl.setRowCount(0)
                for k,v in rows:
                    r = tbl.rowCount(); tbl.insertRow(r)
                    tbl.setItem(r,0,QTableWidgetItem(str(k))); tbl.setItem(r,1,QTableWidgetItem(str(v)))
            finally:
                tbl.setUpdatesEnabled(True)
        fill(self.tbl_sing, sing); fill(self.tbl_pairs, pairs); fill(self.tbl_trips, trips)

    # ---------- Settings
    def _build_tab_settings(self) -> QWidget:
        w = QWidget(); layout = QVBoxLayout(w); layout.setSpacing(8)

        gb = QGroupBox("Pfade"); layout.addWidget(gb); gl = QVBoxLayout(gb); gl.setContentsMargins(10,10,10,10); gl.setSpacing(6)
        def add_row(lbl, line, browse_dir=False, browse_file=False):
            row = QHBoxLayout(); row.setContentsMargins(0,0,0,0); row.setSpacing(8)
            row.addWidget(QLabel(lbl)); row.addWidget(line); btn = QPushButton("…")
            def choose():
                if browse_dir: p = QFileDialog.getExistingDirectory(self,"Ordner wählen",os.getcwd())
                elif browse_file: p, _ = QFileDialog.getOpenFileName(self,"Datei wählen",os.getcwd())
                else: p = ""
                if p: line.setText(p)
            btn.clicked.connect(choose); row.addWidget(btn); gl.addLayout(row)

        self.le_db = QtWidgets.QLineEdit(self.ctrl.paths.get("database_path",""))
        self.le_excel = QtWidgets.QLineEdit(self.ctrl.paths.get("excel_file",""))
        self.le_export = QtWidgets.QLineEdit(self.ctrl.paths.get("export_dir",""))
        add_row("Datenbank:", self.le_db, browse_file=True)
        add_row("Excel:", self.le_excel, browse_file=True)
        add_row("Export-Ordner:", self.le_export, browse_dir=True)

        gf = QGroupBox("Analyten-Filter (Häkchen = aus Suche ausschließen)"); layout.addWidget(gf)
        v = QVBoxLayout(gf); v.setContentsMargins(10,10,10,10); v.setSpacing(6); self.filter_layout = v

        info = QHBoxLayout(); info.setContentsMargins(0,0,0,0); info.setSpacing(8)
        info.addWidget(QLabel("Liste aus Datenbank (BefTag.TestKB)."))
        btn_reload = QPushButton("Neu laden"); btn_reload.clicked.connect(self._reload_filter_list)
        info.addWidget(btn_reload); info.addStretch(1); v.addLayout(info)

        analytes = self.ctrl.list_all_analytes()
        self.wrap_filter, self.scroll_filter, self.chk_filter = self._make_checkbox_grid(analytes, self._cols)
        v.addWidget(self.wrap_filter)

        excluded = set(self.ctrl.get_excluded_analytes())
        for cb in self.chk_filter: cb.setChecked(cb.text() in excluded)

        btns = QHBoxLayout(); btns.setContentsMargins(0,0,0,0); btns.setSpacing(8)
        btn_all = QPushButton("Alle ausschließen"); btn_none = QPushButton("Alle einschließen")
        btn_all.clicked.connect(lambda: [cb.setChecked(True) for cb in self.chk_filter if cb.isVisible()])
        btn_none.clicked.connect(lambda: [cb.setChecked(False) for cb in self.chk_filter if cb.isVisible()])
        btns.addWidget(btn_all); btns.addWidget(btn_none); btns.addStretch(1); v.addLayout(btns)

        btn_save = QPushButton("Einstellungen speichern"); btn_save.clicked.connect(self._save_settings)
        layout.addWidget(btn_save)
        return w

    def _reload_filter_list(self):
        analytes = self.ctrl.list_all_analytes()
        self.filter_layout.removeWidget(self.wrap_filter); self.wrap_filter.setParent(None)
        self.wrap_filter, self.scroll_filter, self.chk_filter = self._make_checkbox_grid(analytes, self._cols)
        self.filter_layout.insertWidget(1, self.wrap_filter)
        excluded = set(self.ctrl.get_excluded_analytes())
        for cb in self.chk_filter: cb.setChecked(cb.text() in excluded)

    def _reload_analyte_controls(self):
        analytes = self.ctrl.list_included_analytes() or self.ctrl.list_all_analytes()

        # Zählungen -> show_all=True
        parent_counts = self.wrap_counts.parent(); self.wrap_counts.setParent(None)
        self.wrap_counts, self.scroll_counts, self.chk_analytes_counts = self._make_checkbox_grid(
            analytes, self._cols, show_all=True
        )
        parent_counts.layout().insertWidget(1, self.wrap_counts)

        # Offene Anforderungen -> begrenzt (Standard)
        parent_open = self.wrap_open.parent(); self.wrap_open.setParent(None)
        self.wrap_open, self.scroll_analytes, self.chk_analytes = self._make_checkbox_grid(analytes, self._cols)
        parent_open.layout().insertWidget(1, self.wrap_open)

    def _save_settings(self):
        self.ctrl.paths["database_path"] = self.le_db.text()
        self.ctrl.paths["excel_file"]   = self.le_excel.text()
        self.ctrl.paths["export_dir"]   = self.le_export.text()
        excluded = [cb.text() for cb in self.chk_filter if cb.isChecked()]
        self.ctrl.update_excluded_analytes(excluded)
        self.ctrl.save_settings()
        QtWidgets.QMessageBox.information(self, "Gespeichert", "Einstellungen gespeichert. Über „Analyten aktualisieren“ die Listen neu laden.")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        for wrap in (getattr(self,"wrap_counts",None), getattr(self,"wrap_open",None), getattr(self,"wrap_filter",None)):
            if wrap is not None:
                self._fit_checkbox_wrapper(wrap)
