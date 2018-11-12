using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Data;
using System.Threading;

namespace TKLRE
{
    /// <summary>
    /// Excel操作类，包含常用的一些操作方法
    /// </summary>

    public class ExcelClass
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        /// <summary>
        /// Excel主程序，必须激活
        /// </summary>
        Excel.Application app { get; set; }
        /// <summary>
        /// 文件集合
        /// </summary>
        Excel.Workbooks workbooks { get; set; }
        /// <summary>
        /// 单个Excel文件
        /// </summary>
        Excel.Workbook workbook { get; set; }
        /// <summary>
        /// 工作表集合
        /// </summary>
        Excel.Worksheets worksheets { get; set; }
        /// <summary>
        /// 单个工作簿
        /// </summary>
        Excel.Worksheet worksheet { get; set; }
        /// <summary>
        /// Excel主程序的PID标识
        /// </summary>
        int AppPid { get; set; }
        /// <summary>
        /// 在主程序初始化前已经运行的PID列表
        /// </summary>
        List<int> PidList = new List<int>();
        /// <summary>
        /// 文件载入时的路径
        /// </summary>
        string Filename = "";


        public ExcelClass()
        {
            app = null;
            workbooks = null;
            workbook = null;
            worksheets = null;
        }

        /// <summary>
        /// 负责初始化Excel环境，如果成功初始化返回0，否则返回-1
        /// </summary>
        /// <returns></returns>
        public int InitExcel()//PASS
        {
            if (app == null)
            {
                app = new Excel.Application();
                app.DisplayAlerts = false;//不弹出确认对话框，直接执行删除工作簿和保存操作
                workbooks = app.Workbooks;
                GetExcelPid();
            }
            else
            {
                return -1;
            }

            return 0;
        }

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <returns></returns>
        public int CreateFile(string name)
        {
            try
            {
                workbook = app.Workbooks.Add(true);
                Filename = name;
                worksheet = workbook.Worksheets.get_Item(1);
            }
            catch
            {
                return -1;
            }

            return 1;
        }

        /// <summary>
        /// 读取路径打开具体文件
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public int OpenFiles(string filepath)//PASS
        {
            try
            {
                workbook = workbooks.Add(filepath);
                Filename = filepath;
            }
            catch
            {
                return -1;
            }

            return 0;
        }


        /// <summary>
        /// 删除工作表内所有图片
        /// </summary>
        /// <param name="wsheet"></param>
        public void DeleteAllPic(Excel.Worksheet wsheet)
        {
            foreach (Excel.Shape i in wsheet.Shapes)
            {
                i.Delete();
            }
        }

        /// <summary>
        /// 添加截图工作表
        /// </summary>
        /// <param name="SheetName"></param>
        public Excel.Worksheet AddOneSheet(string SheetName)//PASS
        {
            Excel.Worksheet ws = workbook.Worksheets.Add(Missing.Value, workbook.Worksheets[GetNumberOfSheets()], Missing.Value, Missing.Value);
            ws.Name = SheetName;

            return ws;
        }

        /// <summary>
        /// 将图片插入工作簿，并返回插入结果（type为0为单图片模式，1为多图模式）
        /// </summary>
        /// <param name="PicPath"></param>
        /// <param name="wsheet"></param>
        /// <returns></returns>
        public int AddPicToSheet(string PicPath, Excel.Worksheet wsheet, List<string> PicPathList, int type)//PASS
        {
            int status = 1, count, i;
            float start = 0;
            string path;

            if (type == 0)
            {
                count = 1;
                path = PicPath;
            }
            else
            {
                count = PicPathList.Count;
                path = PicPathList[0];
            }

            try
            {
                for (i = 0; i < count; i++)
                {
                    Image pic = Image.FromFile(path);
                    float ImgHeight = pic.Height;
                    float ImgWidth = pic.Width;

                    if (ImgHeight > 200)
                    {
                        ImgWidth = 200 * ImgWidth / ImgHeight;
                        ImgHeight = 200;
                    }

                    wsheet.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, start, ImgWidth, ImgHeight);

                    start += ImgHeight + 10;

                    if (type == 1 && i != count - 1)
                    {
                        path = PicPathList[i];
                    }
                }

            }
            catch
            {
                return status;
            }

            return 0;
        }

        /// <summary>
        /// 删除所有截图工作表（初始化用）
        /// </summary>
        public void DeleteSheet()//PASS
        {
            for (int i = GetNumberOfSheets(); i >= 3; i--)
            {
                worksheet = workbook.Worksheets[i];
                worksheet.Delete();
                worksheet = null;
            }

        }

        /// <summary>
        /// 修改单元格的值，type表示插入类型，0为普通值，1为链接，当插入普通值时LinkTitle需要传入空值（x为行坐标，y为列坐标）
        /// </summary>
        /// <param name="value"></param>
        /// <param name="CellX"></param>
        /// <param name="CellY"></param>
        /// <param name="wsheet"></param>
        /// <returns></returns>
        public int EditCellValue(string value, string LinkTitile, int CellX, int CellY, Excel.Worksheet wsheet, int type)//PASS
        {
            try
            {
                if (type == 0)
                {
                    wsheet.Cells[CellX, CellY] = value;
                    wsheet.Cells[CellX, CellY].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    wsheet.Cells[CellX, CellY].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
                else
                {

                    Excel.Range range = wsheet.Cells[CellX, CellY];
                    //Excel.Range excelrange=wsheet.get_Range(range, Missing.Value);
                    range.Hyperlinks.Add(range, value, Missing.Value, Missing.Value, LinkTitile);
                }
            }
            catch
            {
                return -1;
            }

            return 0;
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="wsheet"></param>
        /// <param name="startx"></param>
        /// <param name="starty"></param>
        /// <param name="endx"></param>
        /// <param name="endy"></param>
        public void MergeCells(Excel.Worksheet wsheet,int startx,int starty,int endx,int endy)
        {
            wsheet.Range[wsheet.Cells[startx, starty], wsheet.Cells[endx, endy]].Merge(Type.Missing);
            wsheet.Range[wsheet.Cells[startx, starty], wsheet.Cells[endx, endy]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            wsheet.Range[wsheet.Cells[startx, starty], wsheet.Cells[endx, endy]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public Excel.Worksheet GetSheet(string SheetName)
        {
            Excel.Worksheet s = (Excel.Worksheet)workbook.Worksheets[SheetName];
            return s;
        }

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Excel.Worksheet GetSheet(int index)
        {
            Excel.Worksheet s = (Excel.Worksheet)workbook.Worksheets[index];
            return s;
        }

        /// <summary>
        /// 保存文件并清除Excel环境
        /// </summary>
        /// <param name="FileNames"></param>
        public void SaveFile()//PASS
        {
            //            workbook.SaveAs(Filename, Missing.Value, "", "", Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, false, Missing.Value,
            //Missing.Value, Missing.Value);
                        workbook.SaveAs(Filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            workbooks.Close();
            app.Quit();
            KillSpecialExcel();

            app = null;
            AppPid = -1;
            Filename = "";
            //PidList.Clear();
        }

        /// <summary>
        /// 结束excel进程
        /// </summary>
        public void KillSpecialExcel()
        {
            try
            {
                if (app != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(app.Hwnd), out lpdwProcessId);

                    Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }

        /// <summary>
        /// 获取初始化前已运行的Excel列表
        /// </summary>
        public void GetList()//PASS
        {
            foreach (Process pro in Process.GetProcessesByName("Excel"))
            {
                PidList.Add(pro.Id);
            }
        }

        /// <summary>
        /// 获取初始化后新增加的ExcelPid
        /// </summary>
        public void GetExcelPid()//PASS
        {
            foreach (Process pro in Process.GetProcessesByName("Excel"))
            {
                if (PidList.IndexOf(pro.Id) == -1)
                {
                    AppPid = pro.Id;
                    break;
                }
            }
        }

        /// <summary>
        /// 获取工作簿集合
        /// </summary>
        /// <returns></returns>
        public Excel.Sheets GetWorksheets()//PASS
        {
            return workbook.Worksheets;
        }

        public void SetWorksheetName(string name,int index)
        {
            workbook.Worksheets[index].Name = name;
        }

        /// <summary>
        /// 获取当前工作表个数
        /// </summary>
        /// <returns></returns>
        public int GetNumberOfSheets()//PASS
        {
            return workbook.Worksheets.Count; ;
        }

        /// <summary>
        /// 获取单元格的值，x为行坐标，y为列坐标，index为表格页数
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public string GetCellValue(int x, int y, int index)
        {
            string value = "";

            try
            {
                value = (workbook.Worksheets[index].Cells[x, y]).Text.ToString();
            }
            catch
            {
                value = "Error";
            }

            return value;
        }

        public List<DataTable> GetTableContent()
        {
            try
            {
                List<DataTable> List = new List<DataTable>();
                for(int i = 1; i <= GetNumberOfSheets(); i++)
                {
                    DataTable dt = new DataTable();
                    Excel.Worksheet cursheet = GetSheet(i);
                    string cellContent;
                    dt.TableName = cursheet.Name;
                    int ircnt = cursheet.UsedRange.Rows.Count;
                    int iccnt = cursheet.UsedRange.Columns.Count;
                    Excel.Range range;
                    int ColID = 1;
                    range = (Excel.Range)cursheet.Cells[1, 1];
                    while (iccnt >= ColID)
                    {
                        DataColumn dc = new DataColumn();
                        dc.DataType = Type.GetType("System.String");
                        string strNewColumnName = range.Text.ToString().Trim();
                        if (strNewColumnName.Length == 0) strNewColumnName = "_1";
                        for(int j = 1; j < ColID; j++)
                        {
                            if (dt.Columns[j - 1].ColumnName == strNewColumnName)
                            {
                                strNewColumnName = strNewColumnName + "_1";
                            }
                        }
                        dc.ColumnName = strNewColumnName;
                        dt.Columns.Add(dc);
                        range = (Excel.Range)cursheet.Cells[1, ++ColID];
                    }

                    if (ircnt - 1 > 500)
                    {
                        int b2 = (ircnt - 1) / 10;

                        DataTable dt1 = new DataTable("dt1");
                        dt1 = dt.Clone();
                        SheetOptions sheet1th = new SheetOptions(cursheet, iccnt, 1, b2, dt1);
                        Thread oth1 = new Thread(new ThreadStart(sheet1th.SheetToDataTable));
                        oth1.Start();
                        //sheet1th.SheetToDataTable();
                        Thread.Sleep(1);

                        DataTable dt2 = new DataTable("dt2");
                        dt2 = dt.Clone();
                        SheetOptions sheet2th = new SheetOptions(cursheet, iccnt, b2 + 1, b2 *2 , dt2);
                        Thread oth2 = new Thread(new ThreadStart(sheet2th.SheetToDataTable));
                        oth2.Start();
                        Thread.Sleep(1);

                        DataTable dt3 = new DataTable("dt3");
                        dt3 = dt.Clone();
                        SheetOptions sheet3th = new SheetOptions(cursheet, iccnt, b2 * 2 + 1, b2 * 3, dt3);
                        Thread oth3 = new Thread(new ThreadStart(sheet3th.SheetToDataTable));
                        oth3.Start();
                        Thread.Sleep(1);

                        DataTable dt4 = new DataTable("dt4");
                        dt4 = dt.Clone();
                        SheetOptions sheet4th = new SheetOptions(cursheet, iccnt, b2 * 3 + 1, b2 * 4, dt4);
                        Thread oth4 = new Thread(new ThreadStart(sheet4th.SheetToDataTable));
                        oth4.Start();
                        Thread.Sleep(1);

                        for(int jrow = b2 * 4 + 1; jrow <= ircnt; jrow++)
                        {
                            DataRow dr = dt.NewRow();
                            for(int icol = 1; icol <= iccnt; icol++)
                            {
                                range = (Excel.Range)cursheet.Cells[jrow, icol];
                                cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                                dr[icol - 1] = cellContent;
                            }
                            dt.Rows.Add(dr);
                        }
                        oth1.Join();
                        oth2.Join();
                        oth3.Join();
                        oth4.Join();
                        foreach(DataRow dr in dt2.Rows)
                        {
                            dt1.Rows.Add(dr.ItemArray);
                        }
                        dt2.Clear();
                        dt2.Dispose();
                        foreach (DataRow dr in dt3.Rows)
                        {
                            dt1.Rows.Add(dr.ItemArray);
                        }
                        dt3.Clear();
                        dt3.Dispose();
                        foreach (DataRow dr in dt4.Rows)
                        {
                            dt1.Rows.Add(dr.ItemArray);
                        }
                        dt4.Clear();
                        dt4.Dispose();
                        foreach (DataRow dr in dt.Rows)
                        {
                            dt1.Rows.Add(dr.ItemArray);
                        }
                        dt.Clear();
                        dt.Dispose();
                        List.Add(dt1);
                    }
                    else
                    {
                        for(int irow = 1; irow <= ircnt; irow++)
                        {
                            DataRow dr = dt.NewRow();
                            for(int icol = 1; icol <= iccnt; icol++)
                            {
                                range = (Excel.Range)cursheet.Cells[irow, icol];
                                cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                                dr[icol - 1] = cellContent;
                            }
                            dt.Rows.Add(dr);
                        }
                        List.Add(dt);
                    }
                }
                return List;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 隐藏或显示行，x为行数，0为隐藏，1为显示
        /// </summary>
        /// <param name="x"></param>
        /// <param name="sheet"></param>
        public void HiddenOrShowRows(int x, Excel.Worksheet sheet, int type)
        {
            Excel.Range range = sheet.Rows[x, Missing.Value];
            if (type == 0)
            {
                range.Hidden = true;
            }
            else if (type == 1)
            {
                range.Hidden = false;
            }
        }
        /// <summary>
        /// 确定是否为空值
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public int CellIsNull(int x, int y, int index)
        {
            if (workbook.Worksheets[index].Cells[x, y].value == null)
            {
                return -1;
            }

            return 0;
        }
        /// <summary>
        /// 删除一整行，提供行数和工作表坐标即可
        /// </summary>
        /// <param name="x"></param>
        /// <param name="index"></param>
        public void DeleteRows(int x, int index)
        {
            Excel.Range ranges = GetWorksheets()[index].Rows[x, Missing.Value];
            ranges.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        /// <summary>
        /// 使得宽度自适应
        /// </summary>
        /// <param name="EndX"></param>
        /// <param name="EndY"></param>
        /// <param name="index"></param>
        public void AutoFit(Excel.Worksheet wsheet)
        {
            wsheet.Columns.EntireColumn.AutoFit();
        }

        public void WriteContentViaDataTable(List<DataTable> dtlist)
        {
            int sheetNum = GetNumberOfSheets();
            for(int i = 1; i <= dtlist.Count; i++)
            {
                if (i > sheetNum)
                {
                    AddOneSheet("sheet" + i);
                    sheetNum = GetNumberOfSheets();
                }
                Excel.Worksheet cursheet = GetSheet(i);//获取当前工作表
                int rnum = dtlist[i - 1].Rows.Count;
                int cnum = dtlist[i - 1].Columns.Count;
                int cindex = 0;
                Excel.Range range;
                foreach(DataColumn col in dtlist[i - 1].Columns)
                {
                    cindex++;
                    cursheet.Cells[1, cindex] = col.ColumnName;
                    range = (Excel.Range)cursheet.Cells[1, cindex];
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                object[,] objData = new object[rnum, cnum];

                for(int r = 0; r < rnum; r++)
                {
                    for(int c = 0; c < cnum; c++)
                    {
                        objData[r, c] = dtlist[i - 1].Rows[r][c];
                    }
                }

                range = cursheet.Range[cursheet.Cells[1, 1], cursheet.Cells[rnum == 0 ? 1 : rnum, cnum == 0 ? 1 : cnum]];
                range.NumberFormatLocal = "@";
                range.Value2 = objData;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                cursheet.Columns.EntireColumn.AutoFit();
                workbook.Saved = true;
            }
        }
    }

    class SheetOptions
    {
        Excel.Worksheet worksheet;
        int iColCount;
        int star;
        int end;
        System.Data.DataTable dt;
        public SheetOptions(Excel.Worksheet worksheet, int iColCount, int star, int end, System.Data.DataTable dt)
        {
            this.worksheet = worksheet;
            this.iColCount = iColCount;
            this.star = star;
            this.end = end;
            this.dt = dt;
        }

        public void SheetToDataTable()
        {
            string cellContent;
            Excel.Range range;
            for (int iRow = star; iRow <= end; iRow++)
            {
                DataRow dr = dt.NewRow();
                for (int iCol = 1; iCol <= iColCount; iCol++)
                {
                    range = (Excel.Range)worksheet.Cells[iRow, iCol];
                    cellContent = (range.Value2 == null) ? "" : range.Text.ToString();
                    dr[iCol - 1] = cellContent;
                }
                dt.Rows.Add(dr);
            }
        }
    }

}
