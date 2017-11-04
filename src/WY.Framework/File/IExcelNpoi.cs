using System.Collections.Generic;

namespace WY.Framework.File
{
    /// <summary>
    /// 使用Npoi操作Excel
    /// </summary>
    public interface IExcelNpoi
    {
        /// <summary>
        /// 导出Excel到本地
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="path">文件保存路径</param>
        void FileExport<T>(IEnumerable<T> source, string path);

        /// <summary>
        /// 网络中导出 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <param name="fileName">文件名称（不需要后缀）</param>
        void HttpExport<T>(IEnumerable<T> source, string fileName = "");

        /// <summary>
        /// 本地路径导入文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        IEnumerable<T> FileImport<T>(string fileName) where T : new();

        /// <summary>
        /// 网路导入文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="postedFile"></param>
        /// <returns></returns>
        IEnumerable<T> HttpImport<T>(System.Web.HttpPostedFileBase postedFile) where T : new();
    }
}
