/********************************************************************************
** auth： 吴毅
** date： 2017/11/4 17:32:12
** desc： Excel帮助类
*********************************************************************************/
using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using NPOI;
using NPOI.XSSF.UserModel;
using WY.Framework.Extensions;

namespace WY.Framework.File
{
    public class ExcelNpoi : IExcelNpoi
    {
        private XSSFWorkbook _wb;

        /// <summary>
        /// 导出Excel到本地
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="path">文件保存路径</param>
        public void FileExport<T>(IEnumerable<T> source, string path)
        {
            CreateExcel(source);
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                _wb.Write(fs);
            }
        }

        /// <summary>
        /// 网络中导出 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="fileName">文件名称（不需要后缀）</param>
        public void HttpExport<T>(IEnumerable<T> source, string fileName = "")
        {
            CreateExcel(source);
            if (string.IsNullOrEmpty(fileName))
                fileName = DateTime.Now.ToString("yyyyMMddHHmmss");
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xlsx", fileName));
            _wb.Write(HttpContext.Current.Response.OutputStream);
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// 本地路径导入文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public IEnumerable<T> FileImport<T>(string fileName) where T : new()
        {
            var fileStream = new FileStream(fileName, FileMode.Open);
            return GetDataFromExcel<T>(fileStream);
        }

        /// <summary>
        /// 网路导入文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="postedFile"></param>
        /// <returns></returns>
        public IEnumerable<T> HttpImport<T>(HttpPostedFileBase postedFile) where T : new()
        {
            return GetDataFromExcel<T>(postedFile.InputStream);
        }

        #region 私有函数
        private void CreateExcel<T>(IEnumerable<T> source)
        {
            _wb = new XSSFWorkbook();
            var sh = (XSSFSheet)_wb.CreateSheet("Sheet1");

            POIXMLProperties props = _wb.GetProperties();
            props.CoreProperties.Creator = "WYRMS";
            props.CoreProperties.Created = DateTime.Now;

            var properties = typeof(T).GetProperties().ToList();

            //构建表头
            var header = sh.CreateRow(0);
            for (var j = 0; j < properties.Count; j++)
            {
                var display = properties[j].GetCustomAttributes(typeof(DisplayAttribute), true)
                    .Cast<DisplayAttribute>()
                    .FirstOrDefault();
                header.CreateCell(j).SetCellValue(display != null ? display.Name : properties[j].Name);
            }
            var list = source.ToList();

            //填充数据
            for (var i = 0; i < list.Count; i++)
            {
                var r = sh.CreateRow(i + 1);
                for (var j = 0; j < properties.Count; j++)
                {
                    var value = properties[j].GetValue(list[i], null);
                    r.CreateCell(j).SetCellValue(value == null ? "" : value.ToString());
                }
            }
        }

        private IEnumerable<T> GetDataFromExcel<T>(Stream excelStrem) where T : new()
        {
            _wb = new XSSFWorkbook(excelStrem);
            var list = new List<T>();
            var sh = _wb.GetSheetAt(0) as XSSFSheet;
            var header = sh.GetRow(0);
            var dicColumns = new Dictionary<int, string>();
            var columns = header.Cells.Select(cell => cell.StringCellValue).ToList();
            var properties = typeof(T).GetProperties().ToList();

            for (var i = 0; i < columns.Count; i++)
            {
                foreach (PropertyInfo t in properties)
                {
                    var display = t.GetCustomAttributes(typeof(DisplayAttribute), true)
                        .Cast<DisplayAttribute>()
                        .FirstOrDefault();
                    if (display != null && display.Name == columns[i])
                        dicColumns.Add(i, t.Name);
                    else if (t.Name == columns[i])
                        dicColumns.Add(i, t.Name);
                }
            }

            for (var i = 1; i <= sh.LastRowNum; i++)
            {
                var obj = new T();
                var row = sh.GetRow(i);

                // 过滤空行
                if (row.IsNull())
                    continue;

                foreach (var key in dicColumns.Keys)
                {
                    var property = properties.Find(p => p.Name == dicColumns[key]);
                    var pType = property.PropertyType.FullName;
                    var cellValue = row.GetCell(key);

                    try
                    {
                        switch (pType)
                        {
                            case "System.Int32":
                                property.SetValue(obj, int.Parse(cellValue.ToString()));
                                break;
                            case "System.Int64":
                                property.SetValue(obj, long.Parse(cellValue.ToString()));
                                break;
                            case "System.Double":
                                property.SetValue(obj, double.Parse(cellValue.ToString()));
                                break;
                            case "System.Decimal":
                                property.SetValue(obj, decimal.Parse(cellValue.ToString()));
                                break;
                            case "System.Boolean":
                                property.SetValue(obj, bool.Parse(cellValue.ToString()));
                                break;
                            case "System.Single":
                                property.SetValue(obj, float.Parse(cellValue.ToString()));
                                break;
                            case "System.DateTime":
                                property.SetValue(obj, cellValue.DateCellValue);
                                break;
                            case "System.String":
                                property.SetValue(obj, cellValue.IsNull() ? "" : cellValue.ToString());
                                break;
                            default:
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        var displayName = property.GetCustomAttributes(typeof(DisplayAttribute), true)
                        .Cast<DisplayAttribute>()
                        .FirstOrDefault();

                        throw new Exception("第{0}行，[{1}]字段输入内容格式错误".Fmt(i + 1, displayName.Name));
                    }
                }
                list.Add(obj);
            }
            return list;
        } 
        #endregion
    }
}
