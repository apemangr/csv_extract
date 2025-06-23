import wx
import wx.grid as gridlib
import pandas as pd

# --- Pestaña de visualización de CSV ---

class CSVViewerTab(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        ctrl_panel = wx.Panel(self)
        ctrl_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.load_btn = wx.Button(ctrl_panel, label="Cargar CSV")
        self.load_btn.Bind(wx.EVT_BUTTON, self.on_load_csv)
        ctrl_sizer.Add(self.load_btn, 0, wx.ALL | wx.EXPAND, 5)

        self.filter_txt = wx.TextCtrl(ctrl_panel, value="", size=(120, -1))
        ctrl_sizer.Add(wx.StaticText(ctrl_panel, label="Texto a buscar:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ctrl_sizer.Add(self.filter_txt, 0, wx.ALL, 5)

        self.color_txt = wx.TextCtrl(ctrl_panel, value="#87CEEB", size=(80, -1))
        ctrl_sizer.Add(wx.StaticText(ctrl_panel, label="Color (HEX o nombre):"), 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ctrl_sizer.Add(self.color_txt, 0, wx.ALL, 5)

        self.add_filter_btn = wx.Button(ctrl_panel, label="Agregar filtro")
        self.add_filter_btn.Bind(wx.EVT_BUTTON, self.on_add_filter)
        ctrl_sizer.Add(self.add_filter_btn, 0, wx.ALL, 5)

        self.remove_filter_btn = wx.Button(ctrl_panel, label="Quitar filtro")
        self.remove_filter_btn.Bind(wx.EVT_BUTTON, self.on_remove_filter)
        ctrl_sizer.Add(self.remove_filter_btn, 0, wx.ALL, 5)

        ctrl_panel.SetSizer(ctrl_sizer)
        main_sizer.Add(ctrl_panel, 0, wx.EXPAND | wx.ALL, 5)

        # Lista de filtros activos
        self.filters_listbox = wx.ListBox(self, style=wx.LB_SINGLE)
        main_sizer.Add(wx.StaticText(self, label="Filtros activos:"), 0, wx.LEFT | wx.TOP, 10)
        main_sizer.Add(self.filters_listbox, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        self.apply_btn = wx.Button(self, label="Aplicar filtros")
        self.apply_btn.Bind(wx.EVT_BUTTON, self.on_apply_filters)
        main_sizer.Add(self.apply_btn, 0, wx.ALL | wx.ALIGN_RIGHT, 5)

        self.grid = gridlib.Grid(self)
        self.grid.CreateGrid(0, 0)
        self.grid.EnableEditing(False)
        main_sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(main_sizer)
        self.df = None

        # Lista de filtros: cada uno es un dict {'text':..., 'color':...}
        self.filters = []

    def on_load_csv(self, event):
        with wx.FileDialog(self, "Seleccionar archivo CSV", 
                          wildcard="CSV files (*.csv)|*.csv|Todos los archivos (*.*)|*.*",
                          style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                file_path = dlg.GetPath()
                self.load_csv(file_path)

    def load_csv(self, file_path):
        try:
            self.df = pd.read_csv(file_path)
            self.update_grid()
        except Exception as e:
            wx.MessageBox(f"Error al leer el archivo:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)

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
            self.grid.SetColLabelValue(col_idx, str(col_name))
        max_rows = min(100, len(self.df))
        for row_idx in range(max_rows):
            for col_idx in range(len(self.df.columns)):
                value = str(self.df.iloc[row_idx, col_idx])
                self.grid.SetCellValue(row_idx, col_idx, value)
        self.grid.AutoSizeColumns()
        self.clear_row_colors()

    def clear_row_colors(self):
        # Limpia los colores de todas las filas
        for row in range(self.grid.GetNumberRows()):
            for col in range(self.grid.GetNumberCols()):
                self.grid.SetCellBackgroundColour(row, col, wx.NullColour)
        self.grid.ForceRefresh()

    def on_apply_color(self, event):
        filtro = self.filter_txt.GetValue().strip().lower()
        color_str = self.color_txt.GetValue().strip()
        try:
            color = wx.Colour(color_str)
            if not color.IsOk():
                raise ValueError
        except Exception:
            wx.MessageBox("Color inválido. Usa HEX (#RRGGBB) o nombre en inglés.", "Error", wx.OK | wx.ICON_ERROR)
            return

        self.clear_row_colors()
        if not filtro or self.df is None:
            return

        max_rows = min(100, len(self.df))
        for row_idx in range(max_rows):
            row_text = " ".join([str(self.df.iloc[row_idx, col_idx]).lower() for col_idx in range(len(self.df.columns))])
            if filtro in row_text:
                for col_idx in range(len(self.df.columns)):
                    self.grid.SetCellBackgroundColour(row_idx, col_idx, color)
        self.grid.ForceRefresh()

    def on_add_filter(self, event):
        filtro = self.filter_txt.GetValue().strip()
        color_str = self.color_txt.GetValue().strip()
        if not filtro or not color_str:
            wx.MessageBox("Debes ingresar texto y color.", "Error", wx.OK | wx.ICON_ERROR)
            return
        try:
            color = wx.Colour(color_str)
            if not color.IsOk():
                raise ValueError
        except Exception:
            wx.MessageBox("Color inválido. Usa HEX (#RRGGBB) o nombre en inglés.", "Error", wx.OK | wx.ICON_ERROR)
            return
        # Evitar duplicados exactos
        for f in self.filters:
            if f['text'].lower() == filtro.lower() and f['color'].lower() == color_str.lower():
                wx.MessageBox("Ese filtro ya existe.", "Aviso", wx.OK | wx.ICON_INFORMATION)
                return
        self.filters.append({'text': filtro, 'color': color_str})
        self.update_filters_listbox()

    def on_remove_filter(self, event):
        sel = self.filters_listbox.GetSelection()
        if sel != wx.NOT_FOUND:
            del self.filters[sel]
            self.update_filters_listbox()

    def update_filters_listbox(self):
        self.filters_listbox.Clear()
        for f in self.filters:
            self.filters_listbox.Append(f"{f['text']} → {f['color']}")

    def on_apply_filters(self, event):
        self.clear_row_colors()
        if not self.filters or self.df is None:
            return
        max_rows = min(100, len(self.df))
        for row_idx in range(max_rows):
            row_text = " ".join([str(self.df.iloc[row_idx, col_idx]).lower() for col_idx in range(len(self.df.columns))])
            for f in self.filters:
                if f['text'].lower() in row_text:
                    try:
                        color = wx.Colour(f['color'])
                        if color.IsOk():
                            for col_idx in range(len(self.df.columns)):
                                self.grid.SetCellBackgroundColour(row_idx, col_idx, color)
                    except Exception:
                        pass
                    break  # Solo aplica el primer filtro que coincida
        self.grid.ForceRefresh()

