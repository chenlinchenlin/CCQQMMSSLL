using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

/// <summary>
///ExcelHelper 的摘要说明
/// </summary>
public class ExcelHelper
{
    private object miss = Missing.Value;
    private Excel.Application m_objExcel;
    private Excel.Workbooks m_objBooks;
    private Excel.Workbook m_objBook;
    private Excel.Worksheet sheet;

    public ExcelHelper()
    {
        this.m_objExcel = new Excel.Application();
    }

    public ExcelHelper(string filestr)
    {
        this.m_objExcel = new Excel.Application();

    }

    public Excel.Worksheet CurrentSheet
    {
        get
        {
            return sheet;
        }
        set
        {
            this.sheet = value;
        }
    }

    public Excel.Workbooks CurrentWorkBooks
    {
        get
        {
            return this.m_objBooks;
        }
        set
        {
            this.m_objBooks = value;
        }
    }

    public Excel.Workbook CurrentWorkBook
    {
        get
        {
            return this.m_objBook;
        }
        set
        {
            this.m_objBook = value;
        }
    }

    public void setCell(int rowIndex, int colIndex, string tmpvalue)
    {

        Excel.Range oRng = (Excel.Range)sheet.Cells[rowIndex, colIndex];
        oRng.Value2 = tmpvalue;
    }

    public void OpenExcelFile(string filename)
    {
        UserControl(false);

        m_objExcel.Workbooks.Open(filename, miss, miss, miss, miss, miss, miss, miss,
                                miss, miss, miss, miss, miss, miss, miss);

        m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
        m_objBook = m_objExcel.ActiveWorkbook;
        sheet = (Excel.Worksheet)m_objBook.ActiveSheet;
    }

    public void CreateExceFile()
    {
        UserControl(false);
        m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
        m_objBook = (Excel.Workbook)(m_objBooks.Add(miss));
        sheet = (Excel.Worksheet)m_objBook.ActiveSheet;
    }

    public void SaveAs(string FileName)
    {
        m_objBook.SaveAs(FileName, miss, miss, miss, miss,
         miss, Excel.XlSaveAsAccessMode.xlNoChange,
         Excel.XlSaveConflictResolution.xlLocalSessionChanges,
         miss, miss, miss, miss);
        m_objBook.Close(false, miss, miss);
        m_objBooks.Close();
    }

    public void UserControl(bool usercontrol)
    {
        if (m_objExcel == null) { return; }
        m_objExcel.UserControl = usercontrol;
        m_objExcel.DisplayAlerts = usercontrol;
        m_objExcel.Visible = usercontrol;
    }

    public void ReleaseExcel()
    {
        m_objExcel.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject((object)m_objExcel);
        System.Runtime.InteropServices.Marshal.ReleaseComObject((object)m_objBooks);
        System.Runtime.InteropServices.Marshal.ReleaseComObject((object)m_objBook);
        System.Runtime.InteropServices.Marshal.ReleaseComObject((object)sheet);
        m_objExcel = null;
        m_objBooks = null;
        m_objBook = null;
        sheet = null;
        GC.Collect();
    }

    public bool KillAllExcelApp()
    {
        try
        {
            if (m_objExcel != null)
            {
                m_objExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_objExcel);
                foreach (System.Diagnostics.Process theProc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    if (theProc.CloseMainWindow() == false)
                    {
                        theProc.Kill();
                    }
                }
                m_objExcel = null;
                return true;
            }
        }
        catch
        {
            return false;
        }
        return true;
    }

    [DllImport("User32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
    public void KillSpecialExcel()
    {
        try
        {
            IntPtr t = new IntPtr(m_objExcel.Hwnd);
            int ExcelID = 0;
            GetWindowThreadProcessId(t, out ExcelID);
            Process.GetProcessById(ExcelID).Kill();
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public void ExportExcel(string filename, DataTable data, System.Collections.Generic.IList<string> columnNames)
    {
        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        if (xlApp == null) { return; }
        System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
        Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
        Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

        worksheet.Cells.NumberFormatLocal = "@";

        for (int i = 0; i < columnNames.Count; i++)
        {
            worksheet.Cells[1, i + 1] = columnNames[i].ToString();
        }
        for (int j = 1; j <= data.Rows.Count; j++)
        {
            for (int n = 1; n <= data.Columns.Count; n++)
            {
                //worksheet.Cells[j + 1, n] = "'"+data.Rows[j - 1][n - 1].ToString();
                worksheet.Cells[j + 1, n] = data.Rows[j - 1][n - 1].ToString();
            }
        }
        object objOpt = System.Reflection.Missing.Value;
        workbook.SaveAs(filename, objOpt, objOpt, objOpt, objOpt, objOpt, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, objOpt, objOpt, objOpt, objOpt, objOpt);
        xlApp.DisplayAlerts = false;
        xlApp.Visible = true;
        xlApp.Quit();
        xlApp = null;
        workbook = null;
        worksheet = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
