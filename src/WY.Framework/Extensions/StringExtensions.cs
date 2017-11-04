/********************************************************************************
** auth： 吴毅
** date： 2017/11/4 17:52:45
** desc： 尚未编写描述
*********************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WY.Framework.Extensions
{
    public static class StringExtensions
    {
        #region 空值判断

        public static bool IsNull<T>(this T obj) where T : class
        {
            return obj == null;
        }

        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }

        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }

        #endregion

        #region 字符串处理

        public static string Fmt(this string format, params object[] args)
        {
            return string.Format(format, args);
        }

        /// <summary>
        /// 截断字符串
        /// </summary>
        /// <param name="str"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string Sub(this string str, int length)
        {
            if (str.Length <= length)
                return str;
            return str.Substring(0, length);
        }

        public static string ToStr(this object input)
        {
            return input.IsNull() ? null : input.ToString();
        }

        #endregion
    }
}
