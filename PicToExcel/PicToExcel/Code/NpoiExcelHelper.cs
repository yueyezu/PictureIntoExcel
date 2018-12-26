using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PicToExcel.Code
{
    /// <summary>
    /// Excel写入，或者读取时的对象实体
    /// </summary>
    public class ExcelCellValue
    {
        public ExcelCellValue() { }
        public ExcelCellValue(int row, int col, string value)
        {
            this.row = row;
            this.col = col;
            this.value = value;
        }
        public int row;
        public int col;
        public string value;
    }

    /// <summary>
    /// Excel操作的帮助类
    /// 目前帮助类只支持 03版本的Excel，07的NPOI貌似操作不同，这里自行扩展了。
    /// </summary>
    public class NpoiExcelHelper : IDisposable
    {
        LogHelper logger = new LogHelper("excelHelp");

        private string filePath = "";
        private FileStream fs = null;
        private HSSFWorkbook workBook = null;
        private ISheet sheet = null;

        /// <summary>
        /// 创建Excel操作对象xls
        /// </summary>
        public NpoiExcelHelper()
        {
            workBook = new HSSFWorkbook();
        }

        /// <summary>
        /// Execle文件操作初始化xls
        /// </summary>
        /// <param name="path">Excel文件的路径，绝对路径</param>
        public NpoiExcelHelper(string path)
        {
            filePath = path;
            //文件不存在，则创建一个新的文件
            if (File.Exists(path))
            {
                //获取Excel对象
                fs = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook = new HSSFWorkbook(fs);
            }
            else
            {
                workBook = new HSSFWorkbook();
            }
        }

        /// <summary>
        /// 设置当前操作的sheet页
        /// </summary>
        /// <param name="sheetName"></param>
        public void SetWorkSheet(string sheetName)
        {
            ISheet tempSheet = workBook.GetSheet(sheetName);
            if (tempSheet == null)
            {
                sheet = workBook.CreateSheet(sheetName);
            }
            else
            {
                sheet = tempSheet;
            }
        }

        /// <summary>
        /// 设置列宽度
        /// </summary>
        public void SetColumeWidth(int col)
        {
            sheet.SetColumnWidth(col, 3 * 256);
        }


        #region 数据读取的操作

        /// <summary>
        /// 获取制定单元格的数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public string GetCellValue(string sheetName, int row, int col)
        {
            ISheet sheet = workBook.GetSheet(sheetName);
            return GetCellValue(sheet, row, col);
        }

        public string GetCellValue(int sheetIndex, int row, int col)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);
            return GetCellValue(sheet, row, col);
        }

        private string GetCellValue(ISheet sheet, int row, int col)
        {
            IRow rowData = sheet.GetRow(row); //读取当前行数据
            if (rowData != null)
            {
                ICell cell = rowData.GetCell(col);
                if (cell != null)
                {
                    return cell.ToString();
                }
            }
            return null;
        }

        /// <summary>
        /// 获取一个sheet页中的所有数据信息
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public List<ExcelCellValue> GetSheetValues(string sheetName)
        {
            ISheet sheet = workBook.GetSheet(sheetName);
            return GetSheetValues(sheet);
        }
        public List<ExcelCellValue> GetSheetValues(int sheetIndex)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);
            return GetSheetValues(sheet);
        }

        public List<Dictionary<int, string>> GetSheetValueDictionary(string sheetName)
        {
            ISheet sheet = workBook.GetSheet(sheetName);
            return GetSheetValueDictionary(sheet);
        }


        private List<ExcelCellValue> GetSheetValues(ISheet sheet)
        {
            List<ExcelCellValue> excelCellValues = new List<ExcelCellValue>();
            for (int j = 0; j <= sheet.LastRowNum; j++) //LastRowNum 是当前表的总行数
            {
                IRow row = sheet.GetRow(j); //读取当前行数据
                if (row == null) continue;

                for (int k = 0; k <= row.LastCellNum; k++) //LastCellNum 是当前行的总列数
                {
                    ICell cell = row.GetCell(k); //当前表格
                    if (cell == null) continue;

                    ExcelCellValue cellValue = new ExcelCellValue();
                    cellValue.row = j;
                    cellValue.col = k;
                    cellValue.value = cell.ToString(); //获取表格中的数据并转换为字符串类型

                    excelCellValues.Add(cellValue);
                }
            }
            return excelCellValues;
        }

        private List<Dictionary<int, string>> GetSheetValueDictionary(ISheet sheet)
        {
            var rowValues = new List<Dictionary<int, string>>();
            for (int j = 0; j <= sheet.LastRowNum; j++)
            {
                IRow row = sheet.GetRow(j);
                if (row == null) continue;

                var cellValues = new Dictionary<int, string>();
                for (int k = 0; k <= row.LastCellNum; k++)
                {
                    ICell cell = row.GetCell(k);
                    if (cell == null) continue;
                    cellValues.Add(k, cell.ToString());
                }
                rowValues.Add(cellValues);
            }
            return rowValues;
        }
        #endregion

        #region 向单元格插入数据

        /// <summary>
        /// 向单元格插入数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="cellValue"></param>
        public void SetCellValue(string sheetName, ExcelCellValue cellValue)
        {
            ISheet sheet = workBook.GetSheet(sheetName);

            SetCellValue(sheet, cellValue);
        }
        public void SetCellValue(int sheetIndex, ExcelCellValue cellValue)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);

            SetCellValue(sheet, cellValue);
        }
        private void SetCellValue(ISheet sheet, ExcelCellValue cellValue)
        {
            IRow rowData = sheet.GetRow(cellValue.row); //读取当前行数据
            if (rowData == null)
                rowData = sheet.CreateRow(cellValue.row);

            ICell cell = rowData.GetCell(cellValue.col);
            if (cell == null)
                cell = rowData.CreateCell(cellValue.col);

            cell.SetCellValue(cellValue.value);
        }

        /// <summary>
        /// 批量插入数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="cellValues"></param>
        public void SetCellValue(string sheetName, List<ExcelCellValue> cellValues)
        {
            ISheet sheet = workBook.GetSheet(sheetName);

            SetCellValue(sheet, cellValues);
        }
        public void SetCellValue(int sheetIndex, List<ExcelCellValue> cellValues)
        {
            ISheet sheet = workBook.GetSheetAt(sheetIndex);

            SetCellValue(sheet, cellValues);
        }
        private void SetCellValue(ISheet sheet, List<ExcelCellValue> cellValues)
        {
            foreach (ExcelCellValue cellValue in cellValues)
            {
                SetCellValue(sheet, cellValue);
            }
        }

        #endregion

        #region 设置单元格背景色



        private Dictionary<int, ICellStyle> cellStyles = new Dictionary<int, ICellStyle>();
        private HSSFPalette hssfPalette = null;
        private int colorSpool = 10; //颜色值的精度
        /// <summary>
        /// 设置单元格样式，测试。
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        public bool SetCellColor(int row, int col, Color color)
        {
            IRow oneRow = sheet.GetRow(row);
            if (oneRow == null)
                oneRow = sheet.CreateRow(row);
            ICell cell = oneRow.GetCell(col);
            if (cell == null)
                cell = oneRow.CreateCell(col);

            #region 设置颜色

            ICellStyle cellStyle = ThisCheckNeighbor(color);
            if (cellStyle == null)
            {
                int nowColorrgb = color.ToArgb();

                if (hssfPalette == null)
                {
                    hssfPalette = workBook.GetCustomPalette();
                }

                //NPOI的调色板只允许 0x08-0x40共56种颜色的设置，设置多了就无法展示了。
                int index = cellStyles.Count + 8;
                if (index > 64)
                {
                    return false;
                    //throw new Exception("调色板超过了56种颜色！降低颜色精度试试！");
                }
                hssfPalette.SetColorAtIndex((short)index, color.R, color.G, color.B);
                HSSFColor hssFColor = hssfPalette.FindColor(color.R, color.G, color.B);
                if (hssFColor == null)
                {
                    return false;
                }
                cellStyle = workBook.CreateCellStyle();
                cellStyles.Add(nowColorrgb, cellStyle);
                cellStyle.FillForegroundColor = hssFColor.Indexed;
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }

            #endregion

            cell.CellStyle = cellStyle;
            return true;
        }

        public void SetColotSpool(int spool)
        {
            this.colorSpool = spool;
        }

        /// <summary>
        /// 减少颜色的使用，对相似颜色进行合并处理
        /// </summary>
        /// <param name="colorB"></param>
        /// <returns></returns>
        private ICellStyle ThisCheckNeighbor(Color colorB)
        {
            ICellStyle res = null;
            foreach (int index in cellStyles.Keys)
            {
                Color colorA = Color.FromArgb(index);
                if (Math.Abs(colorA.R - colorB.R) <= colorSpool &&
                Math.Abs(colorA.G - colorB.G) <= colorSpool &&
                Math.Abs(colorA.B - colorB.B) <= colorSpool)
                {
                    res = cellStyles[index];
                    break;
                }
            }
            return res;
        }

        #endregion

        #region 保存部分

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <returns></returns>
        public string SaveToFile()
        {
            using (FileStream filess = File.OpenWrite(filePath))
            {
                workBook.Write(filess);
            }
            return filePath;
        }

        public string SaveToFile(string path)
        {
            using (FileStream filess = File.Open(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
                workBook.Write(filess);
            }
            return path;
        }

        /// <summary>
        /// 将生成的文件保存到内存中
        /// </summary>
        /// <returns></returns>
        public MemoryStream SaveToMemery()
        {
            MemoryStream ms = new MemoryStream();
            workBook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            return ms;
        }

        #endregion

        /// <summary>
        /// 关闭Excel时，执行的方法
        /// </summary>
        public void Dispose()
        {
            if (fs != null)
            {
                fs.Close();
            }
            if (workBook != null)
            {
                workBook = null;
            }
        }
    }
}
