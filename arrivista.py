# coding: utf-8
import sys
import contextlib
import numpy as np
import pandas as pd
import xlsxwriter
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QDesktopWidget, QApplication, 
    QFileDialog, QPushButton, QGridLayout, QLabel, QLineEdit, QComboBox,
    QListView, QGroupBox, QTableView, QHBoxLayout, QVBoxLayout,
    QFormLayout, QCheckBox)
from PyQt5.QtCore import pyqtSlot, QAbstractTableModel, Qt, QModelIndex
from PyQt5.QtGui import QPalette
from arrivista_db import Base, Magazine, Issue, Numbering
from sqlalchemy import create_engine, MetaData, desc
from sqlalchemy.orm import sessionmaker
from collections import namedtuple

APPLICATION_TITLE = "L'Arrivista"
DB_FILENAME = "arrivista.db"
DEFAULT_NUMBER_PREFIX = "n° "
ControlGroup = namedtuple('ControlGroup', 'group tasks')

def getattr_rec(obj, attr):
    dot_pos = attr.find('.')
    if dot_pos > -1:
        return getattr_rec(getattr(obj, attr[:dot_pos]), attr[dot_pos+1:])
    return getattr(obj, attr)

def setattr_rec(obj, attr, val):
    dot_pos = attr.find('.')
    if dot_pos > -1:
        setattr_rec(getattr(obj, attr[:dot_pos]), attr[dot_pos+1:], val)
    setattr(obj, attr, val)

def ignore_exception(ignore_exception=Exception, default_value=None):
    def dec(function):
        def _dec(*args, **kwargs):
            try:
                return function(*args, **kwargs)
            except ignore_exception:
                return default_value
        return _dec
    return dec

def try_parse(type_func, value, default_value=None):
    return ignore_exception(default_value=None)(type_func)(value)

class ArrivistaTableModel(QAbstractTableModel):

    def __init__(self, manager, model, parent=None, filter=None, add_empty_row=False, sort_column=None, sort_order=Qt.AscendingOrder):
        super().__init__(parent)
        self.manager = manager
        self.session = manager.Session()
        self.model = model
        self.sort_column = sort_column
        self.sort_order = sort_order
        self.filter = filter
        self.add_filter = None
        self.current_filter = filter
        self.add_empty_row = add_empty_row
        self.raw_data = []
        self.refresh()

    def _update_current_filter(self):
        if self.add_filter is None:
            self.current_filter = self.filter
        elif self.filter is None:
            self.current_filter = self.add_filter
        else:
            self.current_filter = lambda x: self.filter(x) and self.add_filter(x)

    def setFilter(self, filter):
        self.add_filter = filter
        self._update_current_filter()
        self.refresh(True)

    def resetFilter(self):
        self.setFilter(None)

    def refresh(self, notify=False):
        if notify:
            old_count = self.rowCount()
        if self.sort_column is None:
            if self.current_filter is None:
                self.raw_data = list(self.session.query(self.model))
            else:
                self.raw_data = list(d for d in self.session.query(self.model) if self.current_filter(d))
        else:
            q = self.session.query(self.model)
            joined_table = self.model.joined_tables[self.sort_column]
            if joined_table is not None:
                q = q.join(joined_table)
            col = self.model.sort_fields[self.sort_column]
            q = q.order_by(col if self.sort_order == Qt.AscendingOrder else desc(col))
            if self.current_filter is None:
                self.raw_data = list(q)
            else:
                self.raw_data = list(d for d in q if self.current_filter(d))
        if self.add_empty_row:
            self.raw_data = [self.model()] + self.raw_data
        if notify:
            new_count = self.rowCount()
            if old_count < new_count:
                self.beginInsertRows(QModelIndex(), old_count, new_count-1)
                self.endInsertRows()
            elif old_count > new_count:
                self.beginRemoveRows(QModelIndex(), new_count, old_count-1)
                self.endRemoveRows()
            self.dataChanged.emit(self.createIndex(0, 0), self.createIndex(self.rowCount()-1, self.columnCount()-1))

    def rowCount(self, parent=None):
        return len(self.raw_data)

    def columnCount(self, parent=None):
        return len(self.model.show_columns)

    def data(self, index, role):
        if index.row() >= self.rowCount():
            return None
        if role == Qt.DisplayRole:
            return getattr_rec(self.raw_data[index.row()], self.model.show_columns[index.column()])
        return None

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if section >= self.columnCount():
                return None
            return self.model.column_names[section]
        if orientation == Qt.Vertical:
            if section >= self.rowCount():
                return None
            return section+1
        return None

    def flags(self, index):
        f = Qt.ItemIsEnabled + Qt.ItemIsSelectable
        if self.model.edit_columns[index.column()]:
            f += Qt.ItemIsEditable
        return Qt.ItemFlags(f)

    def setData(self, index, value, role):
        if role != Qt.EditRole:
            return False
        if index.column() >= self.columnCount() or index.row() >= self.rowCount():
            return False
        try:
            setattr_rec(self.raw_data[index.row()], self.model.show_columns[index.column()], value)
            self.session.commit()
            self.dataChanged.emit(index, index)
            return True
        except:
            self.session.rollback()
            return False

    def sort(self, column, order):
        self.sort_column = column
        self.sort_order = order
        self.refresh(True)

    def cloneData(self):
        return list(self.raw_data)

    def getRawData(self):
        return self.raw_data

    def resetConnection(self):
        self.session = self.manager.Session()
        self.refresh(True)
        

class MissingNumbersTableModel(QAbstractTableModel):

    def __init__(self, manager, parent=None):
        super().__init__(parent)
        self.manager = manager
        self.session = manager.Session()
        self.magazine_id = None
        self.magazine = None
        self.raw_data = []
        self.refresh()

    def refresh(self, notify=False):
        if notify:
            old_count = self.rowCount()
        if self.magazine_id is None:
            self.raw_data = []
            self.magazine = None
        else:
            self.magazine = self.session.query(Magazine).get(self.magazine_id)
            self.raw_data = self.magazine.get_missing_numbers()
        if notify:
            new_count = self.rowCount()
            if old_count < new_count:
                self.beginInsertRows(QModelIndex(), old_count, new_count-1)
                self.endInsertRows()
            elif old_count > new_count:
                self.beginRemoveRows(QModelIndex(), new_count, old_count-1)
                self.endRemoveRows()
            self.dataChanged.emit(self.createIndex(0, 0), self.createIndex(self.rowCount()-1, self.columnCount()-1))        

    def rowCount(self, parent=None):
        return len(self.raw_data)

    def columnCount(self, parent=None):
        return 2

    def data(self, index, role):
        if index.row() >= self.rowCount():
            return None
        if role == Qt.DisplayRole:
            return self.raw_data[index.row()][index.column()]
        return None

    def headerData(self, section, orientation, role):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if section >= self.columnCount():
                return None
            return "anno" if section == 0 else "numero"
        if orientation == Qt.Vertical:
            if section >= self.rowCount():
                return None
            return section+1
        return None

    def setMagazineId(self, magazine_id):
        self.magazine_id = magazine_id
        self.refresh(notify=True)
        
    def resetMagazineId(self):
        self.setMagazineId(None)
        
    def cloneData(self):
        return list(self.raw_data)

    def getRawData(self):
        return self.raw_data

    def resetConnection(self):
        self.session = self.manager.Session()
        self.refresh(True)
        

class ArchiveManager:

    def __init__(self, filename):
        self.engine = create_engine('sqlite:///' + filename)
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.meta = MetaData(bind=self.engine)
        self.meta.reflect()

    def update_archive_from_csv(self, csv_path):
        m = pd.read_csv(csv_path).as_matrix()

        # first retrieve all new magazines
        magazine_names = sorted(set(m[:,0]))
        magazines = self.load_magazine_dict()
        new_magazines = [Magazine(name=n) for n in magazine_names if magazines.get(n) is None]

        session = self.Session()
        try:
            session.add_all(new_magazines)
            magazines = self.load_magazine_dict(session=session)
            issues = self.load_issue_dict(session=session, add_selection=True)
            num_magazines = len(new_magazines)
            issues_to_save = []
            num_new_issues = 0
            num_updated_issues = 0
            for mag, y, num in m:
                num = str(num)
                current_magazine = magazines[mag]
                if issues.get(current_magazine.id) is None:
                    issues[current_magazine.id] = {}
                if issues[current_magazine.id].get(y) is None:
                    issues[current_magazine.id][y] = {}
                if issues[current_magazine.id][y].get(num) is None:
                    issue = Issue(magazine=current_magazine, year=y, issue_number=str(num), is_new=False, copies=1)
                    issue.populate_issue_numbers()
                    issues[current_magazine.id][y][num] = [issue, True]
                    issues_to_save.append(issue)
                    num_new_issues += 1
                else:
                    issue_tuple = issues[current_magazine.id][y][num]
                    issue_tuple[1] = True # issue is still there
                    if issue_tuple[0].is_new:
                        issue_tuple[0].is_new = False
                        issues_to_save.append(issue_tuple[0])
                        num_updated_issues += 1
            session.add_all(issues_to_save)
            issues_to_delete = [issue_tuple[0] for mag_id, mag in issues.items() 
                for y, y_list in mag.items() 
                for _, issue_tuple in y_list.items() if issue_tuple[1] is False]
            num_deleted_issues = len(issues_to_delete)

            for x in issues_to_delete:
                session.delete(x)

            session.commit()
            return num_magazines, num_new_issues, num_updated_issues, num_deleted_issues
        except:
            session.rollback()
            raise

    def delete_all_data(self):
        with contextlib.closing(self.engine.connect()) as con:
            trans = con.begin()
            try:
                for table in reversed(self.meta.sorted_tables):
                    con.execute(table.delete())
                trans.commit()
            except:
                trans.rollback()
                raise

    def delete_new_issues(self):
        session = self.Session()
        try:
            for issue in session.query(Issue).filter_by(is_new = True):
                session.delete(issue)
            session.commit()
        except:
            session.rollback()
            raise

    def insert(self, obj):
        session = self.Session()
        try:
            session.add(obj)
            session.commit()
        except:
            session.rollback()
            raise
        
    def insert_all(self, obj_list):
        session = self.Session()
        try:
            session.add_all(obj_list)
            session.commit()
        except:
            session.rollback()
            raise

    def load_magazine_list(self, session=None):
        if session is None:
            session = self.Session()
        return list(session.query(Magazine))
        
    def load_magazine_dict(self, session=None):
        return {mag.name: mag for mag in self.load_magazine_list(session=session)}

    def load_issue_list(self, session=None):
        if session is None:
            session = self.Session()
        return list(session.query(Issue))
        
    def load_issue_dict(self, session=None, add_selection=False):
        issue_list = self.load_issue_list(session=session)
        magazine_ids = set([issue.magazine_id for issue in issue_list])
        issues_for_magazines = {magazine_id: [issue for issue in issue_list
            if issue.magazine_id==magazine_id] for magazine_id in magazine_ids}
        years_for_magazines = {magazine_id: set([issue.year for issue in issues_for_magazines[magazine_id]]) 
            for magazine_id in magazine_ids}
        if add_selection:
            return {magazine_id: {year: {issue.issue_number: [issue, False] for issue in issues_for_magazines[magazine_id] 
                if issue.year==year} 
                for year in years_for_magazines[magazine_id]} for magazine_id in magazine_ids}
        return {magazine_id: {year: {issue.issue_number: issue for issue in issues_for_magazines[magazine_id] 
            if issue.year==year} 
            for year in years_for_magazines[magazine_id]} for magazine_id in magazine_ids}        


class Arrivista(QWidget):

    def __init__(self, manager, title):
        super().__init__()
        self.manager = manager
        self.groups = {}
        self._initUI(title)

    def _showGroup(self, group_to_show, button_to_highlight=None):
        for controlGroup in self.groups:
            if controlGroup.group == group_to_show:
                controlGroup.group.setVisible(True)
                for task in controlGroup.tasks:
                    task()
            else:
                controlGroup.group.setVisible(False)
        for sidebarButton in self.sidebarButtons:
            if sidebarButton == button_to_highlight:
                sidebarButton.setStyleSheet("background-color: #55BCDD; border: 1px solid; padding: 5px; margin: 2px;")
            else:
                sidebarButton.setStyleSheet("background-color: #AAAAAA; border: 1px solid; padding: 5px; margin: 2px;")
            sidebarButton.update()

    def _showGroupSignal(self, group_to_show, button_to_highlight=None):
        def dummy():
            self._showGroup(group_to_show, button_to_highlight)
        return dummy

    def _getExportFilePath(self, default_filename):
        file_path, _ = QFileDialog.getSaveFileName(parent=self, directory='~/Documents/'+default_filename.format(datetime.now().date()),
            filter="Excel files (*.xls *.xlsx)", initialFilter="Excel files (*.xls *.xlsx)", caption="Salva file")
        return file_path

    def _messageBox(self, message):
        print(message)

    def _exportData(self, model_to_export, tree_structure=False, default_filename=''):
        file_path = self._getExportFilePath(default_filename)
        if file_path:
            try:

                data = model_to_export.cloneData()
                data.sort(key=lambda x: (x.magazine.name, 0 if x.year is None else x.year, x.issue_number))
                
                if len(data) == 0:
                    self._messageBox('Nessun dato da esportare')
                    return

                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet()

                bold = workbook.add_format({'bold': True})
                italic = workbook.add_format({'italic': True})

                cur_mag, cur_year = None, None

                for (issue, row) in zip(data, range(len(data))):
                    if not tree_structure or cur_mag is None or issue.magazine_id != cur_mag:
                        worksheet.write(row, 0, issue.magazine.name, bold)
                        cur_mag = issue.magazine_id
                        cur_year = None
                    if not tree_structure or cur_year is None or issue.year != cur_year:
                        cur_year = issue.year
                        worksheet.write(row, 1, '-' if issue.year is None else str(issue.year), italic)
                    worksheet.write(row, 2, str(issue.issue_number))
                    worksheet.write(row, 3, issue.copies)

                workbook.close()
            except Exception as e:
                print(e)
                pass

    def _exportDataSignal(self, model_to_export, export_method, **kwargs):
        def dummy():
            export_method(model_to_export, **kwargs)
        return dummy

    def _exportMissingNumbers(self, model_to_export, default_filename=''):
        default_filename = model_to_export.magazine.name + default_filename
        file_path = self._getExportFilePath(default_filename)
        if file_path:
            try:

                data = model_to_export.cloneData()
                
                if len(data) == 0:
                    self._messageBox('Nessun dato da esportare')
                    return

                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet()

                title = workbook.add_format({'bold': True, 'font_size': 18})
                bold = workbook.add_format({'bold': True})
                italic = workbook.add_format({'italic': True})

                show_year = any(item[0] is not None for item in data)

                worksheet.write(0, 0, model_to_export.magazine.name + ' - numeri mancanti', title)

                cur_year, row, number_col = None, 3, 1 if show_year else 0
                
                if show_year:
                    worksheet.write(2, 0, 'anno', bold)

                worksheet.write(2, number_col, 'numero', bold)

                for (year, number) in data:
                    if cur_year is None or year != cur_year:
                        cur_year = year
                        worksheet.write(row, 0, '-' if year is None else str(year), italic)
                    worksheet.write(row, number_col, str(number))
                    row += 1

                workbook.close()
            except Exception as e:
                print(e)
                pass

    def _exportMissingNumbersSignal(self, model_to_export, export_method, **kwargs):
        def dummy():
            export_method(model_to_export, **kwargs)
        return dummy

    def _resetImportMessage(self):
        self.importMessage.setText("Trascina qui il file .csv contenente il catalogo aggiornato.")

    def _generateIssueFilter(self, magazine_combo=None, year_edit=None, number_edit=None):
        magazine_id = None if magazine_combo is None else magazine_combo.model().getRawData()[magazine_combo.currentIndex()].id
        year = None if year_edit is None else try_parse(int, year_edit.text())
        number = None if number_edit is None or len(number_edit.text()) == 0 else number_edit.text()
        return lambda x: (magazine_id is None or x.magazine_id == magazine_id) \
            and (year is None or x.year == year) \
            and (number is None or x.issue_number.find(number) >= 0)

    def _applyFilter(self, model_to_filter, filter_generator, **kwargs):
        model_to_filter.setFilter(filter_generator(**kwargs))

    def _resetFilter(self, model_to_filter, magazine_combo=None, year_edit=None, number_edit=None):
        model_to_filter.resetFilter()
        if magazine_combo is not None:
            magazine_combo.setCurrentIndex(0)
        if year_edit is not None:
            year_edit.setText('')
        if number_edit is not None:
            number_edit.setText('')

    def _insertIssue(self, model_to_refresh, magazine_combo=None, year_edit=None, number_edit=None):
        try:
            if magazine_combo is None or year_edit is None or number_edit is None:
                raise Exception('Invalid arguments: magazine, year and number widgets must be specified')
            magazine_id = magazine_combo.model().getRawData()[magazine_combo.currentIndex()].id
            year = try_parse(int, year_edit.text())
            issue_number = number_edit.text()
            self.manager.insert(Issue(magazine_id=magazine_id, year=year, issue_number=issue_number, is_new=True, copies=1))
            model_to_refresh.resetConnection()
        except Exception as e:
            self._showError(e)

    def _insertNumbering(self, model_to_refresh, magazine_combo=None, 
        year_from_edit=None, year_to_edit=None, is_yearly_check=None, 
        number_from_edit=None, number_to_edit=None):
        try:
            if magazine_combo is None or year_from_edit is None or year_to_edit is None or is_yearly_check is None or number_from_edit is None or number_to_edit is None:
                raise Exception('Invalid arguments: magazine, year and number widgets must be specified')
            magazine_id = magazine_combo.model().getRawData()[magazine_combo.currentIndex()].id
            year_from = try_parse(int, year_from_edit.text())
            year_to = try_parse(int, year_to_edit.text())
            is_yearly = is_yearly_check.isChecked()
            number_from = try_parse(int, number_from_edit.text())
            number_to = try_parse(int, number_to_edit.text())
            self.manager.insert(Numbering(magazine_id=magazine_id, from_year=year_from, to_year=year_to, 
                is_yearly=is_yearly, from_number=number_from, to_number=number_to))
            model_to_refresh.resetConnection()
        except Exception as e:
            self._showError(e)

    def _filterSignal(self, model_to_filter, filter_generator, **kwargs):
        def dummy():
            self._applyFilter(model_to_filter, filter_generator, **kwargs)
        return dummy

    def _resetFilterSignal(self, model_to_filter, **kwargs):
        def dummy():
            self._resetFilter(model_to_filter, **kwargs)
        return dummy

    def _insertIssueSignal(self, model_to_refresh, **kwargs):
        def dummy():
            self._insertIssue(model_to_refresh, **kwargs)
        return dummy

    def _insertNumberingSignal(self, model_to_refresh, **kwargs):
        def dummy():
            self._insertNumbering(model_to_refresh, **kwargs)
        return dummy

    def _setVisibilitySignal(self, widget, visible):
        def dummy():
            widget.setVisible(visible)
        return dummy

    def _showError(error_message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)

        msg.setText(error_message)
        msg.setWindowTitle("Errore")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    @pyqtSlot(int)
    def _selectedMagazineChanged(self, index):
        if index > 0:
            self.missingNumbersModel.setMagazineId(self.missingNumbersMagazineModel.getRawData()[index].id)
            self.missingNumbersTable.setVisible(True)
        else:
            self.missingNumbersModel.resetMagazineId()
            self.missingNumbersTable.setVisible(False)

    def _createGroups(self, main_grid, row, col, rowspan, colspan):

        # initialize group list
        self.groups = []

        #
        # helper methods
        #
        def getFilterGroup(parent_model, allow_no_selection=True, filter_caption='Filtra', reset_insert_caption='Reimposta', 
            add_filter=None, add_reset_insert=None, reset_insert_signal=self._resetFilterSignal):
            formLayout = QFormLayout()
            formLayout.setSpacing(10)

            magazineModel = ArrivistaTableModel(self.manager, Magazine, add_empty_row=allow_no_selection, sort_column=0)

            cmbMagazine = QComboBox()
            cmbMagazine.setModel(magazineModel)
            txtYear = QLineEdit()
            txtNumber = QLineEdit()
            btnFilter = QPushButton(filter_caption)
            btnResetInsert = QPushButton(reset_insert_caption)

            formLayout.addRow('Testata', cmbMagazine)
            formLayout.addRow('Anno', txtYear)
            formLayout.addRow('Numero', txtNumber)
            formLayout.addRow(btnFilter, btnResetInsert)

            formGroup = QGroupBox()
            formGroup.setLayout(formLayout)

            if add_filter is not None:
                btnFilter.clicked.connect(add_filter)

            btnFilter.clicked.connect(self._filterSignal(parent_model, 
                self._generateIssueFilter, magazine_combo=cmbMagazine, year_edit=txtYear, number_edit=txtNumber))

            if add_reset_insert is not None:
                btnResetInsert.clicked.connect(add_reset_insert)

            btnResetInsert.clicked.connect(reset_insert_signal(parent_model, 
                magazine_combo=cmbMagazine, year_edit=txtYear, number_edit=txtNumber))

            return formGroup, magazineModel

        def getNumberingsFilterGroup(parent_model):
            formLayout = QFormLayout()
            formLayout.setSpacing(10)

            magazineModel = ArrivistaTableModel(self.manager, Magazine, add_empty_row=True, sort_column=0)

            cmbMagazine = QComboBox()
            cmbMagazine.setModel(magazineModel)
            txtYearFrom = QLineEdit()
            txtYearTo = QLineEdit()
            chkIsYearly = QCheckBox()
            txtNumberFrom = QLineEdit()
            txtNumberTo = QLineEdit()
            btnFilter = QPushButton('Filtra')
            btnInsert = QPushButton('Inserisci')

            formLayout.addRow('Testata', cmbMagazine)
            formLayout.addRow('Anno (da)', txtYearFrom)
            formLayout.addRow('Anno (a)', txtYearTo)
            formLayout.addRow('Numerazione annuale', chkIsYearly)
            formLayout.addRow('Numero (da)', txtNumberFrom)
            formLayout.addRow('Numero (a)', txtNumberTo)
            formLayout.addRow(btnFilter, btnInsert)

            formGroup = QGroupBox()
            formGroup.setLayout(formLayout)

            btnFilter.setVisible(False)
            #btnFilter.clicked.connect(self._filterSignal(parent_model, 
            #    self._generateIssueFilter, magazine_combo=cmbMagazine, year_edit=txtYear, number_edit=txtNumber))

            btnInsert.clicked.connect(self._insertNumberingSignal(parent_model, 
                magazine_combo=cmbMagazine, year_from_edit=txtYearFrom, year_to_edit=txtYearTo,
                is_yearly_check=chkIsYearly, number_from_edit=txtNumberFrom, number_to_edit=txtNumberTo))

            return formGroup, magazineModel

        def getMissingNumbersFilterGroup(parent_model):
            formLayout = QFormLayout()
            formLayout.setSpacing(10)

            magazineModel = ArrivistaTableModel(self.manager, Magazine, add_empty_row=True, sort_column=0, filter=lambda x: len(x.numberings) > 0)

            cmbMagazine = QComboBox()
            cmbMagazine.setModel(magazineModel)
            cmbMagazine.currentIndexChanged.connect(self._selectedMagazineChanged)

            formLayout.addRow('Testata', cmbMagazine)

            formGroup = QGroupBox()
            formGroup.setLayout(formLayout)

            return formGroup, magazineModel
        #
        #
        #


        #
        # welcome group
        #
        welcomeLayout = QVBoxLayout()
        welcomeLayout.setSpacing(10)
        welcomeLayout.setAlignment(Qt.AlignTop)

        welcomeLabel = QLabel("Benvenut* nell'Arrivista! Seleziona a lato l'operazione da eseguire.")
        welcomeLayout.addWidget(welcomeLabel, alignment=Qt.AlignHCenter)

        self.welcomeGroup = QGroupBox()
        self.welcomeGroup.setLayout(welcomeLayout)
        self.groups.append(ControlGroup(group=self.welcomeGroup, tasks=[]))
        #
        #
        #


        #
        # import group
        #
        importLayout = QVBoxLayout()
        importLayout.setSpacing(10)
        importLayout.setAlignment(Qt.AlignTop)

        self.importMessage = QLabel('Trascina qui il file .csv contenente il catalogo aggiornato')
        importLayout.addWidget(self.importMessage, alignment=Qt.AlignHCenter)

        self.importGroup = QGroupBox()
        self.importGroup.setLayout(importLayout)
        self.groups.append(ControlGroup(group=self.importGroup, tasks=[self._resetImportMessage]))
        #
        #
        #


        #
        # view all group
        #
        viewAllLayout = QGridLayout()
        viewAllLayout.setSpacing(10)

        viewAllModel = ArrivistaTableModel(self.manager, Issue)

        self.viewAllTable = QTableView()
        self.viewAllTable.setSortingEnabled(True)
        self.viewAllTable.setModel(viewAllModel)
        self.viewAllTable.resizeColumnsToContents()

        btnExport = QPushButton('Esporta')
        btnExport.clicked.connect(self._exportDataSignal(viewAllModel, self._exportData, default_filename='Lista completa {}.xlsx'))

        viewAllFormGroup, viewAllMagazineModel = getFilterGroup(viewAllModel)

        viewAllLayout.addWidget(viewAllFormGroup, 0, 0, 1, 10)
        viewAllLayout.addWidget(self.viewAllTable, 1, 0, 13, 10)
        viewAllLayout.addWidget(btnExport, 19, 8, 1, 2)

        self.viewAllGroup = QGroupBox()
        self.viewAllGroup.setLayout(viewAllLayout)
        self.groups.append(ControlGroup(group=self.viewAllGroup, tasks=[viewAllModel.resetConnection, viewAllMagazineModel.resetConnection]))
        #
        #
        #


        #
        # view duplicates group
        #
        viewDuplicatesLayout = QGridLayout()
        viewDuplicatesLayout.setSpacing(10)

        viewDuplicatesModel = ArrivistaTableModel(self.manager, Issue, filter=lambda x: x.copies > 1)

        self.viewDuplicatesTable = QTableView()
        self.viewDuplicatesTable.setSortingEnabled(True)
        self.viewDuplicatesTable.setModel(viewDuplicatesModel)
        self.viewDuplicatesTable.resizeColumnsToContents()

        btnExport = QPushButton('Esporta')
        btnExport.clicked.connect(self._exportDataSignal(viewDuplicatesModel, self._exportData, default_filename='Duplicati {}.xlsx'))

        viewDuplicatesFormGroup, viewDuplicatesMagazineModel = getFilterGroup(viewDuplicatesModel)

        viewDuplicatesLayout.addWidget(viewDuplicatesFormGroup, 0, 0, 1, 10)
        viewDuplicatesLayout.addWidget(self.viewDuplicatesTable, 1, 0, 13, 10)
        viewDuplicatesLayout.addWidget(btnExport, 19, 8, 1, 2)

        self.viewDuplicatesGroup = QGroupBox()
        self.viewDuplicatesGroup.setLayout(viewDuplicatesLayout)
        self.groups.append(ControlGroup(group=self.viewDuplicatesGroup, tasks=[viewDuplicatesModel.resetConnection, viewDuplicatesMagazineModel.resetConnection]))
        #
        #
        #


        #
        # view new group
        #
        viewNewLayout = QGridLayout()
        viewNewLayout.setSpacing(10)

        viewNewModel = ArrivistaTableModel(self.manager, Issue, filter=lambda x: x.is_new)

        self.viewNewTable = QTableView()
        self.viewNewTable.setSortingEnabled(True)
        self.viewNewTable.setModel(viewNewModel)
        self.viewNewTable.resizeColumnsToContents()

        btnExport = QPushButton('Esporta')
        btnExport.clicked.connect(self._exportDataSignal(viewNewModel, self._exportData, default_filename='Nuovi {}.xlsx'))

        viewNewFormGroup, viewNewMagazineModel = getFilterGroup(viewNewModel)

        viewNewLayout.addWidget(viewNewFormGroup, 0, 0, 1, 10)
        viewNewLayout.addWidget(self.viewNewTable, 1, 0, 13, 10)
        viewNewLayout.addWidget(btnExport, 19, 8, 1, 2)

        self.viewNewGroup = QGroupBox()
        self.viewNewGroup.setLayout(viewNewLayout)
        self.groups.append(ControlGroup(group=self.viewNewGroup, tasks=[viewNewModel.resetConnection, viewNewMagazineModel.resetConnection]))
        #
        #
        #


        #
        # insert issues group
        #
        insertIssuesLayout = QGridLayout()
        insertIssuesLayout.setSpacing(10)

        insertIssuesModel = ArrivistaTableModel(self.manager, Issue)

        self.insertIssuesTable = QTableView()
        self.insertIssuesTable.setSortingEnabled(False)
        self.insertIssuesTable.setModel(insertIssuesModel)
        self.insertIssuesTable.resizeColumnsToContents()
        self.insertIssuesTable.setVisible(False)

        insertIssuesFormGroup, insertIssuesMagazineModel = getFilterGroup(insertIssuesModel,
            allow_no_selection=False, filter_caption='Controlla', reset_insert_caption='Inserisci',
            add_filter=self._setVisibilitySignal(self.insertIssuesTable, True),
            reset_insert_signal=self._insertIssueSignal)

        insertIssuesLayout.addWidget(insertIssuesFormGroup, 0, 0, 1, 10)
        insertIssuesLayout.addWidget(self.insertIssuesTable, 1, 0, 13, 10)

        self.insertIssuesGroup = QGroupBox()
        self.insertIssuesGroup.setLayout(insertIssuesLayout)
        self.groups.append(ControlGroup(group=self.insertIssuesGroup, tasks=[insertIssuesModel.resetConnection, insertIssuesMagazineModel.resetConnection]))
        #
        #
        #


        #
        # numberings group
        #
        numberingsLayout = QGridLayout()
        numberingsLayout.setSpacing(10)

        numberingsModel = ArrivistaTableModel(self.manager, Numbering)

        self.numberingsTable = QTableView()
        self.numberingsTable.setSortingEnabled(True)
        self.numberingsTable.setModel(numberingsModel)
        self.numberingsTable.resizeColumnsToContents()

        numberingsFormGroup, numberingsMagazineModel = getNumberingsFilterGroup(numberingsModel)

        numberingsLayout.addWidget(numberingsFormGroup, 0, 0, 1, 10)
        numberingsLayout.addWidget(self.numberingsTable, 1, 0, 13, 10)

        self.numberingsGroup = QGroupBox()
        self.numberingsGroup.setLayout(numberingsLayout)
        self.groups.append(ControlGroup(group=self.numberingsGroup, tasks=[numberingsModel.resetConnection, numberingsMagazineModel.resetConnection]))
        #
        #
        #


        #
        # numberings group
        #
        numberingsLayout = QGridLayout()
        numberingsLayout.setSpacing(10)

        numberingsModel = ArrivistaTableModel(self.manager, Numbering)

        self.numberingsTable = QTableView()
        self.numberingsTable.setSortingEnabled(True)
        self.numberingsTable.setModel(numberingsModel)
        self.numberingsTable.resizeColumnsToContents()

        numberingsFormGroup, numberingsMagazineModel = getNumberingsFilterGroup(numberingsModel)

        numberingsLayout.addWidget(numberingsFormGroup, 0, 0, 1, 10)
        numberingsLayout.addWidget(self.numberingsTable, 1, 0, 13, 10)

        self.numberingsGroup = QGroupBox()
        self.numberingsGroup.setLayout(numberingsLayout)
        self.groups.append(ControlGroup(group=self.numberingsGroup, tasks=[numberingsModel.resetConnection, numberingsMagazineModel.resetConnection]))
        #
        #
        #


        #
        # missing numbers group
        #
        missingNumbersLayout = QGridLayout()
        missingNumbersLayout.setSpacing(10)

        self.missingNumbersModel = MissingNumbersTableModel(self.manager)

        self.missingNumbersTable = QTableView()
        self.missingNumbersTable.setSortingEnabled(False)
        self.missingNumbersTable.setModel(self.missingNumbersModel)
        self.missingNumbersTable.resizeColumnsToContents()
        self.missingNumbersTable.setVisible(False)

        missingNumbersFormGroup, self.missingNumbersMagazineModel = getMissingNumbersFilterGroup(self.missingNumbersModel)

        btnExport = QPushButton('Esporta')
        btnExport.clicked.connect(self._exportMissingNumbersSignal(self.missingNumbersModel, self._exportMissingNumbers, default_filename=' numeri mancanti {}.xlsx'))

        missingNumbersLayout.addWidget(missingNumbersFormGroup, 0, 0, 1, 10)
        missingNumbersLayout.addWidget(self.missingNumbersTable, 1, 0, 13, 10)
        missingNumbersLayout.addWidget(btnExport, 19, 8, 1, 2)

        self.missingNumbersGroup = QGroupBox()
        self.missingNumbersGroup.setLayout(missingNumbersLayout)
        self.groups.append(ControlGroup(group=self.missingNumbersGroup, tasks=[self.missingNumbersModel.resetConnection, self.missingNumbersMagazineModel.resetConnection]))
        #
        #
        #


        # add all groups in the same grid position
        for controlGroup in self.groups:
            main_grid.addWidget(controlGroup.group, row, col, rowspan, colspan)
                
    def _initUI(self, title):

        # set main window properties
        grid = QGridLayout()
        grid.setSpacing(0)
        self.setLayout(grid)          
        self.resize(1280, 768)
        self.center()
        self.setWindowTitle(title)
        self.setAcceptDrops(True)

        # create groups
        self._createGroups(grid, 0, 3, 10, 10)

        # create sidebar
        menuLayout = QVBoxLayout()
        menuLayout.setSpacing(0)
        menuLayout.setAlignment(Qt.AlignTop)
        menu = QGroupBox()
        menu.setLayout(menuLayout)
        btnImport = QPushButton('Importa catalogo')
        btnViewAll = QPushButton('Tutti i numeri')
        btnViewDuplicates = QPushButton('Numeri doppi')
        btnViewMissing = QPushButton('Numeri mancanti')
        btnViewNew = QPushButton('Numeri nuovi')
        btnInsertIssues = QPushButton('Inserisci numeri')
        btnNumberings = QPushButton('Numerazioni')
        menuLayout.addWidget(btnImport)
        menuLayout.addWidget(btnViewAll)
        menuLayout.addWidget(btnViewDuplicates)
        menuLayout.addWidget(btnViewMissing)
        menuLayout.addWidget(btnViewNew)
        menuLayout.addWidget(btnInsertIssues)
        menuLayout.addWidget(btnNumberings)
        grid.addWidget(menu, 0, 0, 10, 3)

        btnImport.clicked.connect(self._showGroupSignal(self.importGroup, btnImport))
        btnViewAll.clicked.connect(self._showGroupSignal(self.viewAllGroup, btnViewAll))
        btnViewDuplicates.clicked.connect(self._showGroupSignal(self.viewDuplicatesGroup, btnViewDuplicates))
        btnViewNew.clicked.connect(self._showGroupSignal(self.viewNewGroup, btnViewNew))
        btnViewMissing.clicked.connect(self._showGroupSignal(self.missingNumbersGroup, btnViewMissing))
        btnInsertIssues.clicked.connect(self._showGroupSignal(self.insertIssuesGroup, btnInsertIssues))
        btnNumberings.clicked.connect(self._showGroupSignal(self.numberingsGroup, btnNumberings))

        # add buttons to list and modify palette
        buttonPalette = QPalette(btnImport.palette())
        self.sidebarButtons = [btnImport, btnViewAll, btnViewDuplicates, btnViewNew, btnNumberings, btnViewMissing, btnInsertIssues]
        for btn in self.sidebarButtons:
            btn.setAutoFillBackground(True)
            btn.setFlat(True)
            btn.setPalette(buttonPalette)
            btn.update()

        # show welcome group
        self._showGroup(self.welcomeGroup)

        # show main window
        self.show()
               
    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def dragEnterEvent(self, e):
        if self.importGroup.isVisible():
            e.accept()

    def dropEvent(self, e):
        if e.mimeData().hasText():
            filePath = e.mimeData().text()[7:]
            self.importMessage.setText("Importazione in corso dal file {}, attendi...".format(filePath.split('/')[-1]))
            self.update()
            m, ni, ui, di = self.manager.update_archive_from_csv(filePath)
            self.importMessage.setText("Importazione conclusa:\n{} nuove testate, {} nuovi numeri, {} numeri aggiornati e {} numeri eliminati.".format(m, ni, ui, di))

    def _old_NOUSE(self):
        # create all widgets for main grid
        repoLabel = QLabel('File riviste')
        repoEdit = QLineEdit()
        browseButton = QPushButton('...')
        loadButton = QPushButton('Carica')
        self.groupBox = QGroupBox('Verifica duplicati')

        # create all widgets for sub grid
        magazineLabel = QLabel('Rivista')
        yearLabel = QLabel('Anno')
        issueLabel = QLabel('Numero')
        magazineCombo = QComboBox()
        yearCombo = QComboBox()
        issueEdit = QLineEdit()
        searchButton = QPushButton('Cerca')
        searchResults = QTableView()

        # add widgets to grid
        grid.addWidget(repoLabel, 1, 0, 1, 2)
        grid.addWidget(repoEdit, 1, 2, 1, 6)
        grid.addWidget(browseButton, 1, 8)
        grid.addWidget(loadButton, 2, 0)

        # add widgets to sub grid
        subGrid.addWidget(magazineLabel, 1, 0, 1, 2)
        subGrid.addWidget(magazineCombo, 1, 2, 1, 6)
        subGrid.addWidget(yearLabel, 2, 0, 1, 2)
        subGrid.addWidget(yearCombo, 2, 2, 1, 2)
        subGrid.addWidget(issueLabel, 3, 0, 1, 2)
        subGrid.addWidget(issueEdit, 3, 2, 1, 6)
        subGrid.addWidget(searchButton, 4, 0, 1, 2)
        subGrid.addWidget(searchResults, 5, 0, 5, 9)

        # set layout for group box
        self.groupBox.setLayout(subGrid)

        # set table views
        searchResults.setSortingEnabled(True)
        searchResults.setModel(ArrivistaTableModel(self.manager, Issue))
        searchResults.resizeColumnsToContents()

        # hide group by default
        self.groupBox.setVisible(False)

        # add group box to main grid
        grid.addWidget(self.groupBox, 3, 0, 9, 9)

        # add events
        loadButton.clicked.connect(self.loadButtonClicked)
        
        
if __name__ == '__main__':
    try:
        manager = ArchiveManager(DB_FILENAME)
        app = QApplication(sys.argv)
        ex = Arrivista(manager, APPLICATION_TITLE)
        sys.exit(app.exec_())
    except Exception as e:
        print ("ERROR", e)
        sys.exit(1)
