using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ZExcelAddIn
{
    class ZMethodClass
    {
        static String msg;


        static private void logmsg_add(String msg)
        {
            msg = msg + msg + "\n";
        }
        static private String logmsg_get()
        {
            String ret = msg;
            msg = "";
            return ret;
        }
        static public void delete_customviews()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            foreach (Excel.CustomView cv in activeWorkbook.CustomViews)
            {
                cv.Delete();
            }
        }

        static public void delete_autofilter()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {
                // オートフィルターの削除処理
                if (sheet.AutoFilterMode)
                {
                    sheet.AutoFilterMode = false;
                }

                // 非表示行の削除処理
                Boolean repflg = true;
                while (repflg)
                {
                    repflg = false;
                    var z = sheet.Range["A1"].SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 10;    //最終行が非表示の場合にRowの位置が手前にずれるので、繰り返しにより対処する
                    for (int i = z; i >= 1; i--)
                    {
                        if (sheet.Rows[i].Hidden)
                        {
                            sheet.Rows[i].Hidden = false;
                            repflg = true;
                        }
                    }
                }

                // 非表示列の削除処理
                for (int i = sheet.Columns.Count; i >= 1; i--)
                {
                    if (sheet.Columns[i].Hidden)
                    {
                        sheet.Columns[i].Hidden = false;
                    }
                }
            }
        }

        static public void delete_freezepanes()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWindow = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWindow;
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {
                sheet.Activate();
                if (activeWindow.FreezePanes)
                {
                    activeWindow.FreezePanes = false;
                }
                activeWindow.SplitColumn = 0;
                activeWindow.SplitRow = 0;
            }
        }

        static public void delete_group()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWindow = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWindow;
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {

                sheet.Cells.ClearOutline();
            }
        }

        static public void delete_displaygridlines()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWindow = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWindow;
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {
                sheet.Activate();
                if (activeWindow.DisplayGridlines)
                {
                    activeWindow.DisplayGridlines = false;
                }
            }
        }

        static public void reset_zoom(int pzoom, String prange)
        {

            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWindow = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWindow;
            foreach (Excel.Worksheet sheet in activeWorkbook.Worksheets)
            {
                int izoom = pzoom;
                if (sheet.Name == "集計" || sheet.Name == "更新履歴")
                {
                    izoom = 100;
                }
                sheet.Activate();
                sheet.Range[prange].Select();
                activeWindow.Zoom = izoom;
            }
        }

        static public void add_lf()
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWorksheet = activeWorkbook.ActiveSheet;
            var sel = ZExcelAddIn.Globals.ThisAddIn.Application.Selection;
            if (sel != null)
            {
                for (int i = 1; i < sel.Count + 1; i++)
                {
                    activeWorksheet.Cells[sel[i].Row, sel[i].Column].value = activeWorksheet.Cells[sel[i].Row, sel[i].Column].value + "\n";
                }
            }
        }
        static public void renumber(String retext)
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWorksheet = activeWorkbook.ActiveSheet;
            var sel = ZExcelAddIn.Globals.ThisAddIn.Application.Selection;
            int totalcnt = 0;
            String wks = "";
            String strPattern = "^[0-9]*[．]";
            var rx = new Regex(strPattern, RegexOptions.Compiled);

            for (int i = 1; i < sel.Count() + 1; i++)
            {
                if (i == 1)
                {
                    wks = "";
                }
                else
                {
                    //wks = retext + "\n";
                    wks = retext + "";
                }
                if (sel[i].value != null)
                {
                    String[] tmps = sel[i].value.Split('\n');
                    foreach (var tmp in tmps)
                    {
                        String tmp_rg = rx.Replace(tmp, "");
                        if (tmp_rg.Length > 0 && retext != tmp_rg)
                        {
                            if (" " != tmp_rg.Substring(0, 1) && "　" != tmp_rg.Substring(0, 1))
                            {
                                totalcnt++;
                                wks = wks + totalcnt.ToString() + "．" + tmp_rg + "\n";
                                //activeWorksheet.Cells[sel[i].Row, sel[i].Column].value = wks;
                            }
                            else
                            {
                                wks = wks + tmp_rg + "\n";
                            }
                        }
                    }
                    activeWorksheet.Cells[sel[i].Row, sel[i].Column].value = wks.Trim('\n');
                }
            }
        }


        static public void set_column(String cno)
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWorksheet = activeWorkbook.ActiveSheet;
            var sel = ZExcelAddIn.Globals.ThisAddIn.Application.Selection;
            activeWorksheet.Cells[sel[0].Row, cno].Select();
        }

        static public void jump_column(String cno)
        {
            var activeWorkbook = ZExcelAddIn.Globals.ThisAddIn.Application.ActiveWorkbook as Microsoft.Office.Interop.Excel.Workbook;
            var activeWorksheet = activeWorkbook.ActiveSheet;
            var sel = ZExcelAddIn.Globals.ThisAddIn.Application.Selection;
            activeWorksheet.Cells[sel[1].Row, cno].Select();
        }


    }
}
