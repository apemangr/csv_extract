from excel_editor import *
from csv_viewer import *

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