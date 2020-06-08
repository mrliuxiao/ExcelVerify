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
        //返回true表示此字段通过验证
        public virtual bool IsValid(object value, out ErrorInfo errorInfo)
        {
            errorInfo = null;
            return errorInfo == null;
        }
    }
}
