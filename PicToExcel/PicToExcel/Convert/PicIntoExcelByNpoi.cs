using System;
using System.Collections.Generic;
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
                this.excelPath = AppDomain.CurrentDomain.BaseDirectory + "excel";
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
            int width = bmp.Width;
            int height = bmp.Height;
            if (width >= 250)
            {
                throw new Exception("图片宽度像素须控制在250以内.");
            }

            string path = excelPath + "\\" + excelName;

            List<int> hasWriteColor = new List<int>();
            List<int> hasSetWidthCol = new List<int>();
            bool hasNoWriteColor = false;

            while (true)
            {
                hasNoWriteColor = false;
                List<int> CurrentColor = new List<int>();

                #region 一次写入颜色

                //创建Excel对象，进行转化
                using (NpoiExcelHelper excel = new NpoiExcelHelper(path))
                {
                    excel.SetWorkSheet(sheetName);
                    excel.SetColotSpool(colorSpool);

                    for (int h = 0; h < height; h++)
                    {
                        for (int w = 0; w < width; w++)
                        {
                            Color color = bmp.GetPixel(w, h);
                            int tempRgb = color.ToArgb();
                            //判断是否是之前循环已经写入的颜色
                            if (hasWriteColor.Contains(tempRgb)) continue;

                            if (!hasSetWidthCol.Contains(w))
                            {
                                excel.SetColumeWidth(w);
                                hasSetWidthCol.Add(w);
                            }
                            //写入Excel颜色
                            if (excel.SetCellColor(h, w, color))
                            {
                                //记录本次循环颜色中
                                if (!CurrentColor.Contains(tempRgb))
                                {
                                    CurrentColor.Add(tempRgb);
                                }
                            }
                            else
                            {
                                //如果还是有无法写入的颜色，则标记还存在未写入颜色
                                if (hasNoWriteColor) continue;
                                hasNoWriteColor = true;
                            }
                        }
                    }
                    excel.SaveToFile(path);
                }

                #endregion

                logger.Log("本次写入颜色：" + CurrentColor.Count + "种");
                //将本次写入的颜色记录到全部写入颜色中
                hasWriteColor.AddRange(CurrentColor);

                //如果不存在未写入未，则本次结束
                if (!hasNoWriteColor)
                {
                    break;
                }
                logger.Log("存在未写入颜色,继续进行写入");
            }

            logger.Log("图片转化完成：" + picPath);

            return path;
        }
    }
}
