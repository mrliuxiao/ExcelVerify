using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace ExcelVerify
{
    /// <summary>
    /// excel验证特性
    /// </summary>
    public class ExcelAttribute : ValidationAttribute
    {
        /// <summary>
        /// 效验方法 重写此方法，效验方可获得错误信息
        /// </summary>
        /// <param name="value">数据</param>
        /// <param name="errorInfo">错误信息</param>
        /// <returns></returns>
        public virtual bool IsValid(object value, out ErrorInfo errorInfo)
        {
            errorInfo = null;
            return errorInfo == null;
        }
    }
}
