import wx
import wx.grid as gridlib
import pandas as pd
import os
from pathlib import Path
import numpy as np
import re

# --- Funciones utilitarias ---
def excel_column_to_number(col_str):
    col_str = col_str.upper().strip()
    num = 0
    for c in col_str:
        if c < 'A' or c > 'Z':
            return 0
        num = num * 26 + (ord(c) - ord('A'))
    return num

def number_to_excel_column(n):
    n += 1  # Convertir a 1-indexed
    col_str = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str

def deduplicate_columns(columns):
    seen = {}
    deduped = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            deduped.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            deduped.append(col)
    return deduped

# --- Pestaña de carga de Excel ---
class ExcelLoaderTab(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.full_df = None
        self.df = None
        self.file_path = None
        self.sheet_names = []
        self.col_checkboxes = []
        self.range_selectors = []
        # ...existing code for layout and controls...
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        ctrl_panel = wx.Panel(self)
        ctrl_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.load_btn = wx.Button(ctrl_panel, label="Cargar Excel")
        self.load_btn.Bind(wx.EVT_BUTTON, self.on_load_excel)
        ctrl_sizer.Add(self.load_btn, 0, wx.ALL | wx.EXPAND, 5)
        ctrl_sizer.Add(wx.StaticText(ctrl_panel, label="Hoja:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.sheet_choice = wx.Choice(ctrl_panel, choices=[])
        self.sheet_choice.Bind(wx.EVT_CHOICE, self.on_sheet_selected)
        ctrl_sizer.Add(self.sheet_choice, 1, wx.ALL | wx.EXPAND, 5)
        self.export_btn = wx.Button(ctrl_panel, label="Exportar CSV")
        self.export_btn.Bind(wx.EVT_BUTTON, self.on_export_csv)
        self.export_btn.Disable()
        ctrl_sizer.Add(self.export_btn, 0, wx.ALL | wx.EXPAND, 5)
        ctrl_panel.SetSizer(ctrl_sizer)
        main_sizer.Add(ctrl_panel, 0, wx.EXPAND | wx.ALL, 5)
        splitter = wx.SplitterWindow(self, style=wx.SP_LIVE_UPDATE | wx.SP_BORDER)
        self.table_panel = wx.Panel(splitter)
        table_sizer = wx.BoxSizer(wx.VERTICAL)
        self.grid = gridlib.Grid(self.table_panel)
        self.grid.CreateGrid(0, 0)
        self.grid.EnableEditing(False)
        self.grid.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.on_cell_select)
        self.grid.Bind(wx.grid.EVT_GRID_RANGE_SELECT, self.on_range_select)
        table_sizer.Add(self.grid, 1, wx.EXPAND)
        self.table_panel.SetSizer(table_sizer)
        self.selection_panel = wx.Panel(splitter)
        selection_sizer = wx.BoxSizer(wx.VERTICAL)
        range_panel = wx.StaticBoxSizer(wx.VERTICAL, self.selection_panel, "Definir rango de la tabla")
        range_sizer = wx.BoxSizer(wx.VERTICAL)
        self.info_label_original = wx.StaticText(range_panel.GetStaticBox(), label="")
        self.info_label_actual = wx.StaticText(range_panel.GetStaticBox(), label="")
        range_sizer.Add(self.info_label_original, 0, wx.ALL, 2)
        range_sizer.Add(self.info_label_actual, 0, wx.ALL, 2)
        range_types = ["Fila Inicial", "Fila Final", "Columna Inicial", "Columna Final"]
        self.range_controls = {}
        range_grid = wx.FlexGridSizer(4, 2, 5, 5)
        for label in range_types:
            range_grid.Add(wx.StaticText(range_panel.GetStaticBox(), label=f"{label}:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
            if "Fila" in label:
                ctrl = wx.SpinCtrl(range_panel.GetStaticBox(), min=0, max=10000, initial=0)
                ctrl.Bind(wx.EVT_SPINCTRL, self.on_range_changed)
            else:
                ctrl = wx.TextCtrl(range_panel.GetStaticBox(), value="A")
                ctrl.Bind(wx.EVT_TEXT, self.on_range_changed)
            self.range_controls[label] = ctrl
            range_grid.Add(ctrl, 0, wx.EXPAND | wx.ALL, 5)
        range_sizer.Add(range_grid, 0, wx.EXPAND | wx.ALL, 5)
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        set_range_sizer = wx.BoxSizer(wx.HORIZONTAL)
        range_sizer.Add(set_range_sizer, 0, wx.EXPAND | wx.TOP, 10)
        apply_sizer = wx.BoxSizer(wx.HORIZONTAL)
        apply_btn = wx.Button(range_panel.GetStaticBox(), label="Aplicar Rango")
        apply_btn.Bind(wx.EVT_BUTTON, self.apply_range)
        apply_sizer.Add(apply_btn, 1, wx.EXPAND | wx.RIGHT, 5)
        auto_btn = wx.Button(range_panel.GetStaticBox(), label="Detectar Auto")
        auto_btn.Bind(wx.EVT_BUTTON, self.auto_detect_range)
        apply_sizer.Add(auto_btn, 1, wx.EXPAND | wx.LEFT, 5)
        range_sizer.Add(apply_sizer, 0, wx.EXPAND | wx.TOP, 10)
        range_panel.Add(range_sizer, 0, wx.EXPAND | wx.ALL, 5)
        selection_sizer.Add(range_panel, 0, wx.EXPAND | wx.ALL, 5)
        col_panel = wx.StaticBoxSizer(wx.VERTICAL, self.selection_panel, "Columnas a exportar")
        self.col_scroller = wx.ScrolledWindow(col_panel.GetStaticBox())
        self.col_scroller.SetScrollRate(10, 10)
        self.col_sizer = wx.BoxSizer(wx.VERTICAL)
        self.col_scroller.SetSizer(self.col_sizer)
        col_panel.Add(self.col_scroller, 1, wx.EXPAND | wx.ALL, 5)
        selection_sizer.Add(col_panel, 1, wx.EXPAND | wx.ALL, 5)
        row_panel = wx.StaticBoxSizer(wx.VERTICAL, self.selection_panel, "Rango de filas para exportar")
        row_grid = wx.FlexGridSizer(2, 2, 5, 5)
        row_grid.Add(wx.StaticText(row_panel.GetStaticBox(), label="Desde:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.row_start = wx.SpinCtrl(row_panel.GetStaticBox(), min=0, max=10000, initial=0)
        self.row_start.Bind(wx.EVT_SPINCTRL, self.update_preview)
        row_grid.Add(self.row_start, 0, wx.EXPAND | wx.ALL, 5)
        row_grid.Add(wx.StaticText(row_panel.GetStaticBox(), label="Hasta:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.row_end = wx.SpinCtrl(row_panel.GetStaticBox(), min=0, max=10000, initial=0)
        self.row_end.Bind(wx.EVT_SPINCTRL, self.update_preview)
        row_grid.Add(self.row_end, 0, wx.EXPAND | wx.ALL, 5)
        row_panel.Add(row_grid, 0, wx.EXPAND | wx.ALL, 5)
        selection_sizer.Add(row_panel, 0, wx.EXPAND | wx.ALL, 5)
        self.selection_panel.SetSizer(selection_sizer)
        splitter.SplitVertically(self.table_panel, self.selection_panel, sashPosition=600)
        splitter.SetMinimumPaneSize(200)
        main_sizer.Add(splitter, 1, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(main_sizer)
    # ...resto de métodos igual que antes, usando las funciones utilitarias locales...
    def on_load_excel(self, event):
        with wx.FileDialog(self, "Seleccionar archivo Excel", 
                          wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.file_path = dlg.GetPath()
                self.load_excel_sheets()
    def load_excel_sheets(self):
        try:
            xl = pd.ExcelFile(self.file_path)
            self.sheet_names = xl.sheet_names
            self.sheet_choice.SetItems(self.sheet_names)
            if self.sheet_names:
                self.sheet_choice.SetSelection(0)
                self.on_sheet_selected(None)
        except Exception as e:
            wx.MessageBox(f"Error al leer el archivo:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)
    def on_sheet_selected(self, event):
        if not self.file_path or not self.sheet_names:
            return
        sheet_name = self.sheet_names[self.sheet_choice.GetSelection()]
        try:
            self.full_df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
            self.auto_detect_range()
            self.export_btn.Enable()
        except Exception as e:
            wx.MessageBox(f"Error al cargar la hoja:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)
    def auto_detect_range(self):
        if self.full_df is None or self.full_df.empty:
            return
        start_row, start_col = None, None
        for r in range(len(self.full_df)):
            for c in range(len(self.full_df.columns)):
                if not pd.isna(self.full_df.iloc[r, c]):
                    start_row, start_col = r, c
                    break
            if start_row is not None:
                break
        if start_row is None:
            return
        end_row, end_col = start_row, start_col
        for r in range(len(self.full_df)-1, start_row-1, -1):
            for c in range(len(self.full_df.columns)-1, start_col-1, -1):
                if not pd.isna(self.full_df.iloc[r, c]):
                    end_row, end_col = r, c
                    break
            if end_row is not None:
                break
        self.range_controls["Fila Inicial"].SetValue(start_row)
        self.range_controls["Fila Final"].SetValue(end_row)
        self.range_controls["Columna Inicial"].SetValue(number_to_excel_column(start_col))
        self.range_controls["Columna Final"].SetValue(number_to_excel_column(end_col))
        self.apply_range()
    def on_range_changed(self, event):
        if self.full_df is not None:
            max_row = len(self.full_df) - 1
            for ctrl in [self.range_controls["Fila Inicial"], self.range_controls["Fila Final"]]:
                ctrl.SetRange(0, max_row)
    def apply_range(self, event=None):
        if self.full_df is None:
            return
        try:
            start_row = self.range_controls["Fila Inicial"].GetValue()
            end_row = self.range_controls["Fila Final"].GetValue()
            start_col = excel_column_to_number(self.range_controls["Columna Inicial"].GetValue())
            end_col = excel_column_to_number(self.range_controls["Columna Final"].GetValue())
            if start_row > end_row or start_col > end_col:
                wx.MessageBox("Rango inválido: El inicio debe ser menor que el final", "Error", wx.OK | wx.ICON_WARNING)
                return
            self.df = self.full_df.iloc[start_row:end_row+1, start_col:end_col+1].copy()
            if not self.df.empty:
                new_columns = []
                for i, col in enumerate(self.df.iloc[0]):
                    if not pd.isna(col) and col != "":
                        new_columns.append(str(col))
                    else:
                        new_columns.append(f"Columna_{i+1}")
                new_columns = deduplicate_columns(new_columns)
                self.df.columns = new_columns
                self.df = self.df[1:].reset_index(drop=True)
            self.update_grid()
            self.update_column_selection()
            self.update_row_range()
        except Exception as e:
            wx.MessageBox(f"Error al aplicar el rango:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)
    def update_grid(self):
        if self.grid.GetNumberRows() > 0:
            self.grid.DeleteRows(0, self.grid.GetNumberRows())
        if self.grid.GetNumberCols() > 0:
            self.grid.DeleteCols(0, self.grid.GetNumberCols())
        if self.df is None or self.df.empty:
            return
        self.grid.AppendCols(len(self.df.columns))
        self.grid.AppendRows(len(self.df))
        for col_idx, col_name in enumerate(self.df.columns):
            self.grid.SetColLabelValue(col_idx, col_name)
        max_rows = min(100, len(self.df))
        for row_idx in range(max_rows):
            for col_idx in range(len(self.df.columns)):
                value = str(self.df.iloc[row_idx, col_idx])
                self.grid.SetCellValue(row_idx, col_idx, value)
        self.grid.AutoSizeColumns()
    def update_column_selection(self):
        for child in self.col_scroller.GetChildren():
            child.Destroy()
        self.col_checkboxes = []
        if self.df is None:
            return
        for col in self.df.columns:
            cb = wx.CheckBox(self.col_scroller, label=col)
            cb.SetValue(True)
            cb.Bind(wx.EVT_CHECKBOX, self.update_preview)
            self.col_sizer.Add(cb, 0, wx.ALL, 3)
            self.col_checkboxes.append(cb)
        self.col_scroller.Layout()
        self.col_scroller.FitInside()
    def update_row_range(self):
        if self.df is None or self.df.empty:
            max_row = 0
        else:
            max_row = len(self.df) - 1
        self.row_start.SetRange(0, max_row)
        self.row_end.SetRange(0, max_row)
        self.row_start.SetValue(0)
        self.row_end.SetValue(max_row)
        self.update_preview()
    def update_preview(self, event=None):
        if self.df is None or self.df.empty:
            return
        start_row = self.row_start.GetValue()
        end_row = self.row_end.GetValue() + 1
        selected_cols = self.get_selected_columns()
        preview_df = self.df[selected_cols].iloc[start_row:end_row]
        self.update_grid_with_data(preview_df)
    def update_grid_with_data(self, df):
        if self.grid.GetNumberRows() > 0:
            self.grid.DeleteRows(0, self.grid.GetNumberRows())
        if self.grid.GetNumberCols() > 0:
            self.grid.DeleteCols(0, self.grid.GetNumberCols())
        if df is None or df.empty:
            return
        start_row = self.range_controls["Fila Inicial"].GetValue()
        end_row = self.range_controls["Fila Final"].GetValue()
        row_originales = list(range(start_row, end_row + 1))
        start_col = excel_column_to_number(self.range_controls["Columna Inicial"].GetValue())
        end_col = excel_column_to_number(self.range_controls["Columna Final"].GetValue())
        col_originales = [number_to_excel_column(i) for i in range(start_col, end_col + 1)]
        col_nuevas = list(df.columns)
        num_cols = len(df.columns)
        num_rows = len(df)
        self.grid.AppendCols(num_cols)
        self.grid.AppendRows(num_rows)
        for col_idx, col_name in enumerate(col_nuevas):
            label = f"{col_name} ({col_originales[col_idx]})"
            self.grid.SetColLabelValue(col_idx, label)
        max_rows = min(100, num_rows)
        for row_idx in range(max_rows):
            row_label = f"{row_idx} ({row_originales[row_idx]})"
            self.grid.SetRowLabelValue(row_idx, row_label)
            for col_idx in range(num_cols):
                value = str(df.iloc[row_idx, col_idx])
                self.grid.SetCellValue(row_idx, col_idx, value)
        self.grid.AutoSizeColumns()
    def on_cell_select(self, event):
        event.Skip()
    def on_range_select(self, event):
        event.Skip()
    def get_selected_columns(self):
        if self.df is None:
            return []
        selected = []
        for cb, col in zip(self.col_checkboxes, self.df.columns):
            if cb.GetValue():
                selected.append(col)
        return selected
    def on_export_csv(self, event):
        if self.df is None or self.df.empty:
            wx.MessageBox("No hay datos para exportar", "Error", wx.OK | wx.ICON_WARNING)
            return
        selected_cols = self.get_selected_columns()
        start_row = self.row_start.GetValue()
        end_row = self.row_end.GetValue() + 1
        if not selected_cols:
            wx.MessageBox("Seleccione al menos una columna", "Error", wx.OK | wx.ICON_WARNING)
            return
        filtered_df = self.df[selected_cols].iloc[start_row:end_row]
        default_name = f"{Path(self.file_path).stem}_filtered.csv"
        with wx.FileDialog(self, "Guardar CSV", defaultFile=default_name,
                          wildcard="CSV files (*.csv)|*.csv",
                          style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                save_path = dlg.GetPath()
                try:
                    filtered_df.to_csv(save_path, index=False)
                    wx.MessageBox(f"Archivo guardado en:\n{save_path}", "Éxito", wx.OK | wx.ICON_INFORMATION)
                except Exception as e:
                    wx.MessageBox(f"Error al guardar:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

# --- Pestaña de visualización de CSV ---
class CSVViewerTab(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        label = wx.StaticText(self, label="Esta pestaña mostrará los archivos CSV generados")
        label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(label, 0, wx.ALL | wx.CENTER, 20)
        self.SetSizer(sizer)

# --- Ventana principal ---
class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Excel to CSV Converter", size=(1000, 700))
        notebook = wx.Notebook(self)
        self.excel_tab = ExcelLoaderTab(notebook)
        self.csv_tab = CSVViewerTab(notebook)
        notebook.AddPage(self.excel_tab, "Cargar Excel")
        notebook.AddPage(self.csv_tab, "Visualizar CSV")
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(notebook, 1, wx.EXPAND)
        self.SetSizer(sizer)
        self.Centre()
        self.Show()

if __name__ == "__main__":
    app = wx.App()
    frame = MainFrame()
    app.MainLoop()