using System;
using System.Drawing;
using PicToExcel.Code;

namespace PicToExcel.Convert
{
    public class PicIntoExcelByNpoi
    {
        LogHelper logger = new LogHelper("PicConvert");

        private string picPath = "";
        private string excelPath = "";
        private int colorSpool;

        /// <summary>
        /// 初始化转化对象
        /// </summary>
        /// <param name="picPath">图片文件路径</param>
        /// <param name="excelPath">Excel文件存放的路径</param>
        /// <param name="colorSpool">转化时颜色的精度，为了控制转化成果</param>
        public PicIntoExcelByNpoi(string picPath, string excelPath, int colorSpool = 10)
        {
            this.picPath = picPath;
            this.colorSpool = colorSpool;
            if (string.IsNullOrEmpty(excelPath))
            {
                this.excelPath = AppDomain.CurrentDomain.BaseDirectory + "excel\\";
            }
            else
            {
                this.excelPath = excelPath;
            }
        }

        /// <summary>
        /// 开始进行转化
        /// </summary>
        /// <param name="excelName">生成Excel的名称</param>
        /// <param name="sheetName">生成ExcelSheet的名称</param>
        /// <returns></returns>
        public string StartConvert(string excelName = "temp.xls", string sheetName = "pic")
        {
            logger.Log("图片转化开始：" + picPath);
            // 从图片文件创建image对象
            Image img = Image.FromFile(picPath);
            Bitmap bmp = new Bitmap(img);
            img.Dispose();
            string path = excelPath + excelName;

            //创建Excel对象，进行转化
            using (ExcelHelper excel = new ExcelHelper())
            {
                excel.SetWorkSheet(sheetName);
                excel.SetColotSpool(colorSpool);

                int width = bmp.Width;
                int height = bmp.Height;

                if (width >= 250)
                {
                    throw new Exception("图片宽度像素须控制在250以内.");
                }

                for (int h = 0; h < height; h++)
                {
                    for (int w = 0; w < width; w++)
                    {
                        Color color = bmp.GetPixel(w, h);
                        excel.SetCellColor(h, w, color);
                    }
                }

                excel.SaveToFile(path);
            }
            logger.Log("图片转化完成：" + picPath);

            return path;
        }
    }
}
