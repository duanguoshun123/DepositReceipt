using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Xml;
using CommonHelper;

namespace ExcelHelperCommon
{
    public class ExportHelper<T> where T : class, new()
    {

        /// <summary>
        /// 将数据转换成excel文件流输出  ->单表单导出接口
        /// </summary>
        /// <returns></returns>
        public static MemoryStream CreateExcelStreamByDatas(List<T> objectDatas, KeyValuePair<string, string> excelHeader, string xmlPath, ref string fileName)
        {
            // excel工作簿
            IWorkbook workbook = new XSSFWorkbook();
            //导入数据到sheet表单
            CreateExcelSheetByDatas(objectDatas, excelHeader.Key, excelHeader.Value, ref workbook, xmlPath);
            ////文件夹不存在则创建
            DirectoryInfo TheFolder = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + @"/Template/");
            if (!TheFolder.Exists)
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "/Template/");
            }
            fileName = fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            var fileNameNew = AppDomain.CurrentDomain.BaseDirectory + "/Template/" + fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            //这里生成ISheet并且填充数据
            FileStream fileStream = new FileStream(fileNameNew, FileMode.CreateNew);
            workbook.Write(fileStream);
            //重新读取文件
            fileStream = File.Open(fileNameNew, FileMode.Open);
            //return fileStream;
            byte[] buffer = new byte[fileStream.Length];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(buffer, 0, (int)fileStream.Length);
            fileStream.Dispose();
            fileStream.Close();
            //删除临时文件
            //System.IO.File.Delete(fileName);
            MemoryStream memoryStream = new MemoryStream(buffer);
            return memoryStream;
            ////保存为Excel文件  
            //using (FileStream fs = new FileStream(savaPath, FileMode.Create, FileAccess.Write))
            //{
            //    workbook.Write(ms);
            //    //ms.Flush();
            //    var buf = ms.ToArray();
            //    KillSpecialExcel();
            //    fs.Write(buf, 0, buf.Length);
            //    fs.Flush();
            //    fs.Position = 0;
            //    fs.Close();
            //    fs.Dispose();
            //    // 导出成功后打开  
            //    System.Diagnostics.Process.Start(savaPath);
            //}
        }


        /// <summary>
        /// 根据传入数据新建sheet表单到指定workbook
        /// </summary>
        /// <param name="objectDatas"></param>
        /// <param name="excelHeader"></param>
        /// <param name="sheetName"></param>
        /// <param name="regulars"></param>
        /// <param name="workbook"></param>
        private static void CreateExcelSheetByDatas(List<T> objectDatas, string excelHeader, string sheetName, ref IWorkbook workbook, string xmlPath)
        {
            var regulars = ModelConvertHelper<T>.GetExportRegulars(xmlPath);

            // excel sheet表单
            ISheet sheet = workbook.CreateSheet(sheetName);
            // excel行数
            int rows = 0;

            #region 单元格 -表头格式

            #region 表头字体

            IFont fontTitle = workbook.CreateFont();
            fontTitle.FontHeightInPoints = 12;
            fontTitle.Boldweight = (short)FontBoldWeight.Bold;

            #endregion

            ICellStyle styleTitle = workbook.CreateCellStyle();
            styleTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleTitle.SetFont(fontTitle);
            styleTitle.VerticalAlignment = VerticalAlignment.Center;
            //styleTitle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index; //XSSFColor
            //styleTitle.FillPattern = FillPattern.SolidForeground;
            #endregion

            #region 单元格 -表体格式

            #region 表体字体

            IFont fontMessage = workbook.CreateFont();
            fontMessage.FontHeightInPoints = 10;

            #endregion

            ICellStyle styleMessage = workbook.CreateCellStyle();
            styleMessage.Alignment = HorizontalAlignment.Center;
            styleMessage.SetFont(fontMessage);
            styleMessage.VerticalAlignment = VerticalAlignment.Center;

            ICellStyle styleMessageSpecialG = workbook.CreateCellStyle();//特殊单元格（填充颜色绿色）
            styleMessageSpecialG.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialG.SetFont(fontMessage);
            styleMessageSpecialG.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialG.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialG.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;

            ICellStyle styleMessageSpecialY = workbook.CreateCellStyle();//特殊单元格(填充颜色黄色)
            styleMessageSpecialY.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialY.SetFont(fontMessage);
            styleMessageSpecialY.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialY.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialY.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;

            ICellStyle styleMessageSpecialR = workbook.CreateCellStyle();//特殊单元格（填充颜色红色）
            styleMessageSpecialR.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialR.SetFont(fontMessage);
            styleMessageSpecialR.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialR.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialR.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            #endregion
            if (!string.IsNullOrEmpty(excelHeader))//表头存在
            {
                // 创建表头并赋值  导出的对象列数为表头的合并列数
                int firstRowCellCount = regulars.Count;//GetAttributeCount(objectDatas.First());
                IRow headerRow = sheet.CreateRow(rows);
                headerRow.HeightInPoints = 40;
                var headerCell = headerRow.CreateCell(0);
                headerCell.SetCellValue(excelHeader);

                // 合并表头
                var cellRangeAddress = new CellRangeAddress(rows, rows, 0, firstRowCellCount - 1);
                sheet.AddMergedRegion(cellRangeAddress);
                // 设置表头格式
                headerCell.CellStyle = styleTitle;
                rows++;
            }
            //生成表头(属性表头)
            if (objectDatas.Any())
            {
                // excel列数
                int cells = -1;
                // 创建数据行
                var firstRow = sheet.CreateRow(rows);
                firstRow.HeightInPoints = 25;
                var objectData = objectDatas.FirstOrDefault();
                foreach (System.Reflection.PropertyInfo p in objectData.GetType().GetProperties())
                {
                    var regular = regulars.Find(t => t.PropertyName == p.Name);
                    if (regular != null)
                    {
                        cells++;
                        //throw new Exception("导出excel时，出现未配置字段。表：" + objectData.GetType().Name + ",字段：" + p.Name);
                        var firstRowCell = firstRow.CreateCell(cells);
                        firstRowCell.SetCellValue(regular.ExportFieldName);
                        sheet.SetColumnWidth(cells, regular.ExportFieldName.Length * 256 * 4);
                        firstRowCell.CellStyle = styleMessage;
                    }
                }
            }
            // 反射object对象，遍历字段
            foreach (var objectData in objectDatas)
            {
                rows++;
                // excel列数
                int cells = -1;
                // 创建数据行
                var messageRow = sheet.CreateRow(rows);
                messageRow.HeightInPoints = 16;
                foreach (PropertyInfo p in objectData.GetType().GetProperties())
                {
                    var regular = regulars.Find(t => t.PropertyName == p.Name);
                    if (regular != null)
                    {
                        cells++;
                        var messageCell = messageRow.CreateCell(cells);
                        var value = p.GetValue(objectData);
                        if (value == null)
                        {
                            messageCell.SetCellValue("");
                        }
                        else
                        {
                            switch (regular.DataType)
                            {
                                case "DateTime":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("yyyy-MM-dd HH:mm:ss"));
                                    }
                                    break;
                                case "Date":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("yyyy-MM-dd"));
                                    }
                                    break;
                                case "Time":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("HH:mm:ss"));
                                    }
                                    break;
                                case "Int":
                                    messageCell.SetCellValue(Convert.ToInt32(value));
                                    break;
                                case "Double":
                                    messageCell.SetCellValue(Convert.ToDouble(value));
                                    break;
                                case "Decimal2":
                                    var valueC2 = Convert.ToDouble(value).ToString("f2");
                                    messageCell.SetCellValue(valueC2);
                                    break;
                                case "Decimal4":
                                    var valueC4 = Convert.ToDouble(value).ToString("f4");
                                    messageCell.SetCellValue(valueC4);
                                    break;
                                case "Decimal5":
                                    var valueC5 = Convert.ToDouble(value).ToString("f5");
                                    messageCell.SetCellValue(valueC5);
                                    break;
                                case "Bool":
                                    var setValue = "是";
                                    if (!(bool)value)
                                    {
                                        setValue = "否";
                                    }
                                    messageCell.SetCellValue(setValue);
                                    break;
                                default:
                                    messageCell.SetCellValue(value.ToString());
                                    break;
                            }
                            if (regular.PropertyName == "IndicatorLight")//指示灯
                            {
                                if (value.ToString() == "@08@")
                                {
                                    messageCell.CellStyle = styleMessageSpecialG;
                                }
                                if (value.ToString() == "@09@")
                                {
                                    //styleMessageSpecial.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                    messageCell.CellStyle = styleMessageSpecialY;
                                }
                                if (value.ToString() == "@0A@")
                                {
                                    messageCell.CellStyle = styleMessageSpecialR;
                                }
                                continue;
                            }
                        }
                        messageCell.CellStyle = styleMessage;
                    }

                }
            }
        }


        /// <summary>
        /// 将数据转换成excel文件流输出  ->单表单导出接口(需要颜色标记)
        /// </summary>
        /// <returns></returns>
        public static MemoryStream CreateExcelStreamByDatasByColor(List<T> objectDatas, KeyValuePair<string, string> excelHeader, string xmlPath, ref string fileName, string needColorProperty)
        {
            // excel工作簿
            IWorkbook workbook = new XSSFWorkbook();
            //导入数据到sheet表单
            CreateExcelSheetByDatas(objectDatas, excelHeader.Key, excelHeader.Value, ref workbook, xmlPath);
            ////文件夹不存在则创建
            DirectoryInfo TheFolder = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + @"/Template/");
            if (!TheFolder.Exists)
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "/Template/");
            }
            fileName = fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            var fileNameNew = AppDomain.CurrentDomain.BaseDirectory + "/Template/" + fileName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            //这里生成ISheet并且填充数据
            FileStream fileStream = new FileStream(fileNameNew, FileMode.CreateNew);
            workbook.Write(fileStream);
            //重新读取文件
            fileStream = File.Open(fileNameNew, FileMode.Open);
            //return fileStream;
            byte[] buffer = new byte[fileStream.Length];
            fileStream.Seek(0, SeekOrigin.Begin);
            fileStream.Read(buffer, 0, (int)fileStream.Length);
            fileStream.Dispose();
            fileStream.Close();
            //删除临时文件
            //System.IO.File.Delete(fileName);
            MemoryStream memoryStream = new MemoryStream(buffer);
            return memoryStream;
            ////保存为Excel文件  
            //using (FileStream fs = new FileStream(savaPath, FileMode.Create, FileAccess.Write))
            //{
            //    workbook.Write(ms);
            //    //ms.Flush();
            //    var buf = ms.ToArray();
            //    KillSpecialExcel();
            //    fs.Write(buf, 0, buf.Length);
            //    fs.Flush();
            //    fs.Position = 0;
            //    fs.Close();
            //    fs.Dispose();
            //    // 导出成功后打开  
            //    System.Diagnostics.Process.Start(savaPath);
            //}
        }
        /// <summary>
        /// 根据传入数据新建sheet表单到指定workbook
        /// </summary>
        /// <param name="objectDatas"></param>
        /// <param name="excelHeader"></param>
        /// <param name="sheetName"></param>
        /// <param name="regulars"></param>
        /// <param name="workbook"></param>
        private static void CreateExcelSheetByDatasByColor(List<T> objectDatas, string excelHeader, string sheetName, ref IWorkbook workbook, string xmlPath, string needColorProperty)
        {
            var regulars = ModelConvertHelper<T>.GetExportRegulars(xmlPath);

            // excel sheet表单
            ISheet sheet = workbook.CreateSheet(sheetName);
            // excel行数
            int rows = 0;

            #region 单元格 -表头格式

            #region 表头字体

            IFont fontTitle = workbook.CreateFont();
            fontTitle.FontHeightInPoints = 12;
            fontTitle.Boldweight = (short)FontBoldWeight.Bold;

            #endregion

            ICellStyle styleTitle = workbook.CreateCellStyle();
            styleTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleTitle.SetFont(fontTitle);
            styleTitle.VerticalAlignment = VerticalAlignment.Center;
            //styleTitle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index; //XSSFColor
            //styleTitle.FillPattern = FillPattern.SolidForeground;
            #endregion

            #region 单元格 -表体格式

            #region 表体字体

            IFont fontMessage = workbook.CreateFont();
            fontMessage.FontHeightInPoints = 10;

            #endregion

            #region 特殊单元格字体集合处理
            ICellStyle styleMessage = workbook.CreateCellStyle();
            styleMessage.Alignment = HorizontalAlignment.Center;
            styleMessage.SetFont(fontMessage);
            styleMessage.VerticalAlignment = VerticalAlignment.Center;

            ICellStyle styleMessageSpecialG = workbook.CreateCellStyle();//特殊单元格（填充颜色绿色）
            styleMessageSpecialG.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialG.SetFont(fontMessage);
            styleMessageSpecialG.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialG.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialG.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;

            ICellStyle styleMessageSpecialY = workbook.CreateCellStyle();//特殊单元格(填充颜色黄色)
            styleMessageSpecialY.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialY.SetFont(fontMessage);
            styleMessageSpecialY.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialY.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialY.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;

            ICellStyle styleMessageSpecialR = workbook.CreateCellStyle();//特殊单元格（填充颜色红色）
            styleMessageSpecialR.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialR.SetFont(fontMessage);
            styleMessageSpecialR.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialR.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialR.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;

            ICellStyle styleMessageSpecialBro = workbook.CreateCellStyle();//特殊单元格（填充颜色棕色）
            styleMessageSpecialBro.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialBro.SetFont(fontMessage);
            styleMessageSpecialBro.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialBro.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialBro.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Brown.Index;

            ICellStyle styleMessageSpecialBlue = workbook.CreateCellStyle();//特殊单元格（填充颜色蓝色）
            styleMessageSpecialBlue.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialBlue.SetFont(fontMessage);
            styleMessageSpecialBlue.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialBlue.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;

            ICellStyle styleMessageSpecialBlue02 = workbook.CreateCellStyle();//特殊单元格（填充颜色蓝色）
            styleMessageSpecialBlue.Alignment = HorizontalAlignment.Center;
            styleMessageSpecialBlue.SetFont(fontMessage);
            styleMessageSpecialBlue.VerticalAlignment = VerticalAlignment.Center;
            styleMessageSpecialBlue.FillPattern = FillPattern.SolidForeground;
            styleMessageSpecialBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            List<ICellStyle> styleList = new List<ICellStyle>()
            {
                styleMessageSpecialG,
                styleMessageSpecialY,
                styleMessageSpecialR,
                styleMessageSpecialBro,
                styleMessageSpecialBlue
            };
            #endregion
            #endregion
            if (!string.IsNullOrEmpty(excelHeader))//表头存在
            {
                // 创建表头并赋值  导出的对象列数为表头的合并列数
                int firstRowCellCount = regulars.Count;//GetAttributeCount(objectDatas.First());
                IRow headerRow = sheet.CreateRow(rows);
                headerRow.HeightInPoints = 40;
                var headerCell = headerRow.CreateCell(0);
                headerCell.SetCellValue(excelHeader);

                // 合并表头
                var cellRangeAddress = new CellRangeAddress(rows, rows, 0, firstRowCellCount - 1);
                sheet.AddMergedRegion(cellRangeAddress);
                // 设置表头格式
                headerCell.CellStyle = styleTitle;
                rows++;
            }
            //生成表头(属性表头)
            if (objectDatas.Any())
            {
                // excel列数
                int cells = -1;
                // 创建数据行
                var firstRow = sheet.CreateRow(rows);
                firstRow.HeightInPoints = 25;
                var objectData = objectDatas.FirstOrDefault();
                foreach (System.Reflection.PropertyInfo p in objectData.GetType().GetProperties())
                {
                    var regular = regulars.Find(t => t.PropertyName == p.Name);
                    if (regular != null)
                    {
                        cells++;
                        //throw new Exception("导出excel时，出现未配置字段。表：" + objectData.GetType().Name + ",字段：" + p.Name);
                        var firstRowCell = firstRow.CreateCell(cells);
                        firstRowCell.SetCellValue(regular.ExportFieldName);
                        sheet.SetColumnWidth(cells, regular.ExportFieldName.Length * 256 * 4);
                        firstRowCell.CellStyle = styleMessage;
                    }
                }
            }
            // 反射object对象，遍历字段
            foreach (var objectData in objectDatas)
            {
                rows++;
                // excel列数
                int cells = -1;
                // 创建数据行
                var messageRow = sheet.CreateRow(rows);
                messageRow.HeightInPoints = 16;
                foreach (PropertyInfo p in objectData.GetType().GetProperties())
                {
                    var regular = regulars.Find(t => t.PropertyName == p.Name);
                    if (regular != null)
                    {
                        cells++;
                        var messageCell = messageRow.CreateCell(cells);
                        var value = p.GetValue(objectData);
                        if (value == null)
                        {
                            messageCell.SetCellValue("");
                        }
                        else
                        {
                            switch (regular.DataType)
                            {
                                case "DateTime":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("yyyy-MM-dd HH:mm:ss"));
                                    }
                                    break;
                                case "Date":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("yyyy-MM-dd"));
                                    }
                                    break;
                                case "Time":
                                    if (Convert.ToDateTime(value) == DateTime.MinValue)
                                    {
                                        messageCell.SetCellValue("");
                                    }
                                    else
                                    {
                                        messageCell.SetCellValue(
                                            Convert.ToDateTime(value).ToString("HH:mm:ss"));
                                    }
                                    break;
                                case "Int":
                                    messageCell.SetCellValue(Convert.ToInt32(value));
                                    break;
                                case "Double":
                                    messageCell.SetCellValue(Convert.ToDouble(value));
                                    break;
                                case "Decimal2":
                                    var valueC2 = Convert.ToDouble(value).ToString("f2");
                                    messageCell.SetCellValue(valueC2);
                                    break;
                                case "Decimal4":
                                    var valueC4 = Convert.ToDouble(value).ToString("f4");
                                    messageCell.SetCellValue(valueC4);
                                    break;
                                case "Decimal5":
                                    var valueC5 = Convert.ToDouble(value).ToString("f5");
                                    messageCell.SetCellValue(valueC5);
                                    break;
                                case "Bool":
                                    var setValue = "是";
                                    if (!(bool)value)
                                    {
                                        setValue = "否";
                                    }
                                    messageCell.SetCellValue(setValue);
                                    break;
                                default:
                                    messageCell.SetCellValue(value.ToString());
                                    break;
                            }
                            if (regular.PropertyName == "IndicatorLight")//指示灯
                            {
                                if (value.ToString() == "@08@")
                                {
                                    messageCell.CellStyle = styleMessageSpecialG;
                                }
                                if (value.ToString() == "@09@")
                                {
                                    //styleMessageSpecial.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                    messageCell.CellStyle = styleMessageSpecialY;
                                }
                                if (value.ToString() == "@0A@")
                                {
                                    messageCell.CellStyle = styleMessageSpecialR;
                                }
                                continue;
                            }
                        }
                        messageCell.CellStyle = styleMessage;
                    }

                }
            }
        }
    }
}
