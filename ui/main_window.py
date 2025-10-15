from PyQt6 import QtWidgets, QtCore
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QComboBox, QDateEdit, QTableWidget, QTableWidgetItem, QMessageBox,
    QCheckBox, QTabWidget, QScrollArea, QGridLayout, QFrame, QFileDialog
)
from PyQt6.QtCore import QDate
import datetime, os, math

from controller.main_controller import MainController
from util.paths import resource_path


class MainWindow(QMainWindow):
    def __init__(self, controller: MainController):
        super().__init__()
        self.ctrl = controller
        self.setWindowTitle("LabStats – Befund-Statistik")
        self.resize(1200, 780)
        self._apply_style()

        tabs = QTabWidget()
        tabs.addTab(self._build_tab_counts(), "Zählungen")
        tabs.addTab(self._build_tab_open(), "Offene Anforderungen")
        tabs.addTab(self._build_tab_suspected(), "Nicht entnommen?")
        tabs.addTab(self._build_tab_settings(), "Einstellungen")

        self.setCentralWidget(tabs)

        try:
            self.ctrl.monthly_export_if_first()
        except Exception as ex:
            print("Monthly export failed:", ex)

    from util.paths import resource_path
    ...

    def _apply_style(self):
        try:
            qss_path = resource_path("resources/style.qss")  # <-- funktioniert auch in OneFile
            with open(qss_path, "r", encoding="utf-8") as f:
                self.setStyleSheet(f.read())
        except Exception:
            pass
    # --- Tab 1: Zählungen ---
    def _build_tab_counts(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        gb = QGroupBox("Zeitraum & Filter")
        gbl = QHBoxLayout(gb)

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)

        today = datetime.date.today()
        first = today.replace(day=1)
        self.start_date.setDate(QDate(first.year, first.month, first.day))
        self.end_date.setDate(QDate(today.year, today.month, today.day))

        gbl.addWidget(QLabel("Start:"))
        gbl.addWidget(self.start_date)
        gbl.addWidget(QLabel("Ende:"))
        gbl.addWidget(self.end_date)

        self.combo_analyte = QComboBox()
        self.combo_analyte.addItem("— Analyt wählen —", userData=None)
        for a in self.ctrl.analytes:
            self.combo_analyte.addItem(a, userData=a)

        gbl.addWidget(QLabel("Analyt (TestKB):"))
        gbl.addWidget(self.combo_analyte)

        self.status_only_open = QCheckBox("Nur offene")
        self.status_only_open.setChecked(False)
        gbl.addWidget(self.status_only_open)

        btn_run = QPushButton("Berechnen")
        btn_run.clicked.connect(self._run_counts)
        gbl.addWidget(btn_run)

        layout.addWidget(gb)

        self.table_counts = QTableWidget(0, 4)
        self.table_counts.setHorizontalHeaderLabels(["Kategorie", "Wert", "Hinweis", "Details"])
        # FIX: PyQt6 enum
        self.table_counts.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        layout.addWidget(self.table_counts)
        return w

    def _run_counts(self):
        start = datetime.datetime(self.start_date.date().year(), self.start_date.date().month(), self.start_date.date().day())
        end = datetime.datetime(self.end_date.date().year(), self.end_date.date().month(), self.end_date.date().day(), 23, 59, 59)
        analyte = self.combo_analyte.currentData()
        only_open = self.status_only_open.isChecked()

        rows = self.ctrl.build_counts_rows(start, end, analyte, only_open)
        self.table_counts.setRowCount(len(rows))
        for r, (a, b, c, d) in enumerate(rows):
            self.table_counts.setItem(r, 0, QTableWidgetItem(a))
            self.table_counts.setItem(r, 1, QTableWidgetItem(b))
            self.table_counts.setItem(r, 2, QTableWidgetItem(c))
            self.table_counts.setItem(r, 3, QTableWidgetItem(d))

    # --- Tab 2: Offene Anforderungen ---
    def _build_tab_open(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        # 1) Einzeilige Zeitfilter-Leiste (ab Datum X)
        line = QHBoxLayout()
        line.addWidget(QLabel("Offene Anforderungen ab Datum:"))
        self.since_date = QDateEdit()
        self.since_date.setCalendarPopup(True)
        today = datetime.date.today()
        first = today.replace(day=1)
        self.since_date.setDate(QDate(first.year, first.month, first.day))
        line.addWidget(self.since_date)
        self.btn_open_count = QPushButton("Zählen")
        self.btn_open_count.clicked.connect(self._run_open)
        line.addWidget(self.btn_open_count)
        line.addStretch(1)
        layout.addLayout(line)

        # 2) Checkbox-Grid (8 Spalten), max. 1/3 Fensterhöhe
        self.chk_analytes: list[QCheckBox] = []
        self.scroll_analytes = QScrollArea()
        self.scroll_analytes.setWidgetResizable(True)

        inner = QWidget()
        grid = QGridLayout(inner)
        grid.setContentsMargins(6, 6, 6, 6)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(4)

        cols = 8
        for idx, a in enumerate(self.ctrl.analytes):
            chip = QFrame()
            chip.setObjectName("analyteChip")
            chip.setFrameShape(QFrame.Shape.NoFrame)
            chip_layout = QHBoxLayout(chip)
            chip_layout.setContentsMargins(6, 2, 6, 2)
            chip_layout.setSpacing(4)

            cb = QCheckBox(a)
            cb.setChecked(False)
            self.chk_analytes.append(cb)
            chip_layout.addWidget(cb)

            r, c = divmod(idx, cols)
            grid.addWidget(chip, r, c)

        inner.setLayout(grid)
        self.scroll_analytes.setWidget(inner)
        layout.addWidget(self.scroll_analytes)

        # Maximalhöhe dynamisch auf 1/3 der Fensterhöhe begrenzen
        self._update_analyte_scroll_max_height()

        # „Alle auswählen“-Button
        btn_row = QHBoxLayout()
        btn_all = QPushButton("Alle auswählen")
        def _select_all():
            any_unchecked = any(not cb.isChecked() for cb in self.chk_analytes)
            for cb in self.chk_analytes:
                cb.setChecked(any_unchecked)
        btn_all.clicked.connect(_select_all)
        btn_row.addWidget(btn_all)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        # 3) Ergebnis-Tabelle: Zweier-Tupel pro Zeile (gleich breite Spalten!)
        self.table_open = QTableWidget(0, 4)
        self.table_open.setHorizontalHeaderLabels(["Analyt (1)", "Offene (1)", "Analyt (2)", "Offene (2)"])
        # FIX: PyQt6 enum
        self.table_open.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.table_open)

        return w

    def _update_analyte_scroll_max_height(self):
        max_h = int(self.height() * 0.33)
        self.scroll_analytes.setMaximumHeight(max(220, max_h))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "scroll_analytes"):
            self._update_analyte_scroll_max_height()

    def _run_open(self):
        analytes = [cb.text() for cb in self.chk_analytes if cb.isChecked()]
        if not analytes:
            QMessageBox.warning(self, "Hinweis", "Bitte mindestens einen Analyt (TestKB) auswählen.")
            return

        since = datetime.datetime(self.since_date.date().year(), self.since_date.date().month(), self.since_date.date().day())
        rows = self.ctrl.build_open_counts_since(analytes, since)  # List[(code, count)]

        # In 2er-Tupel auf Zeilen verteilen: (A1, N1, A2, N2)
        n_pairs = math.ceil(len(rows) / 2)
        self.table_open.setRowCount(n_pairs)

        for i in range(n_pairs):
            code1, cnt1 = rows[2*i]
            self.table_open.setItem(i, 0, QTableWidgetItem(code1))
            self.table_open.setItem(i, 1, QTableWidgetItem(str(cnt1)))
            if 2*i + 1 < len(rows):
                code2, cnt2 = rows[2*i + 1]
                self.table_open.setItem(i, 2, QTableWidgetItem(code2))
                self.table_open.setItem(i, 3, QTableWidgetItem(str(cnt2)))
            else:
                self.table_open.setItem(i, 2, QTableWidgetItem(""))
                self.table_open.setItem(i, 3, QTableWidgetItem(""))

    # --- Tab 3: Verdacht „nicht entnommen?“ ---
    def _build_tab_suspected(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        btn_refresh = QPushButton("Liste aktualisieren")
        btn_refresh.clicked.connect(self._refresh_suspected)
        layout.addWidget(btn_refresh)

        self.table_susp = QTableWidget(0, 4)
        self.table_susp.setHorizontalHeaderLabels(["ProbenNr", "Order-Zeit", "Anzahl Anforderungen", "Analyte"])
        # FIX: PyQt6 enum
        self.table_susp.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        layout.addWidget(self.table_susp)

        btn_delete = QPushButton("Ausgewählte Probe löschen")
        btn_delete.clicked.connect(self._delete_selected_sample)
        layout.addWidget(btn_delete)

        return w

    def _refresh_suspected(self):
        try:
            rows = self.ctrl.suspected_missing_blood_draw()
        except Exception as ex:
            QMessageBox.critical(self, "Fehler", str(ex))
            return
        self.table_susp.setRowCount(len(rows))
        for r, row in enumerate(rows):
            self.table_susp.setItem(r, 0, QTableWidgetItem(str(row.get("ProbenNr", ""))))
            self.table_susp.setItem(r, 1, QTableWidgetItem(str(row.get("OrderTime", ""))))
            self.table_susp.setItem(r, 2, QTableWidgetItem(str(row.get("NumReq", ""))))
            self.table_susp.setItem(r, 3, QTableWidgetItem(str(row.get("Analytes", ""))))

    def _delete_selected_sample(self):
        row = self.table_susp.currentRow()
        if row < 0:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst eine Zeile auswählen.")
            return
        proben_nr = self.table_susp.item(row, 0).text()
        if not proben_nr:
            return
        if QMessageBox.question(
            self, "Löschen bestätigen",
            f"Probe {proben_nr} wirklich löschen? Diese Aktion ist irreversibel."
        ) != QMessageBox.StandardButton.Yes:
            return
        try:
            deleted = self.ctrl.delete_sample(proben_nr)
            QMessageBox.information(self, "Gelöscht", f"Gelöschte Zeilen: {deleted}")
            self._refresh_suspected()
        except Exception as ex:
            QMessageBox.critical(self, "Fehler", str(ex))

    # --- Tab 4: Einstellungen ---
    def _build_tab_settings(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        gb = QGroupBox("Pfade")
        gbl = QHBoxLayout(gb)

        self.le_db = QtWidgets.QLineEdit(self.ctrl.paths.get("database_path", ""))
        self.le_excel = QtWidgets.QLineEdit(self.ctrl.paths.get("excel_file", ""))
        self.le_export = QtWidgets.QLineEdit(self.ctrl.paths.get("export_dir", ""))
        self.le_analyt = QtWidgets.QLineEdit(self.ctrl.paths.get("analytes_file", ""))

        def add_row(lbl, line, browse_dir=False, browse_file=False):
            box = QHBoxLayout()
            box.addWidget(QLabel(lbl))
            box.addWidget(line)
            btn = QPushButton("…")
            def choose():
                if browse_dir:
                    p = QFileDialog.getExistingDirectory(self, "Ordner wählen", os.getcwd())
                elif browse_file:
                    p, _ = QFileDialog.getOpenFileName(self, "Datei wählen", os.getcwd())
                else:
                    p = ""
                if p:
                    line.setText(p)
            btn.clicked.connect(choose)
            box.addWidget(btn)
            layout.addLayout(box)

        layout.addWidget(gb)
        add_row("Datenbank:", self.le_db, browse_file=True)
        add_row("Excel:", self.le_excel, browse_file=True)
        add_row("Export-Ordner:", self.le_export, browse_dir=True)
        add_row("Analyte.txt:", self.le_analyt, browse_file=True)

        btn_save = QPushButton("Speichern")
        btn_save.clicked.connect(self._save_settings)
        layout.addWidget(btn_save)

        return w

    def _save_settings(self):
        import configparser
        cfg = configparser.ConfigParser()
        cfg["paths"] = {
            "database_path": self.le_db.text(),
            "excel_file": self.le_excel.text(),
            "export_dir": self.le_export.text(),
            "analytes_file": self.le_analyt.text()
        }
        with open("config/settings.ini", "w", encoding="utf-8") as f:
            cfg.write(f)
        QMessageBox.information(self, "Gespeichert", "Einstellungen gespeichert. Bitte App neu starten.")
