using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Org.BouncyCastle.Asn1.Cmp;

namespace ExcelVerify
{
    ///// <summary>
    ///// 不能为null、""
    ///// </summary>
    //public class NotNullAttribute : ExcelAttribute
    //{
    //    //返回true表示此字段通过验证
    //    public override bool IsValid(object value, out ErrorInfo errorInfo)
    //    {
    //        string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
    //        errorInfo = null;
    //        if (!string.IsNullOrEmpty(stringValue))
    //            errorInfo = new ErrorInfo { ErrorMsg = "不能为空" };
    //        return errorInfo == null;
    //    }
    //}

    ///// <summary>
    ///// 字符串中间不能包含空格
    ///// </summary>
    //public class NoSpaceAttribute : ExcelAttribute
    //{
    //    private static readonly Regex noSpaceRegex = new Regex(@"^[^\s]+$", RegexOptions.Compiled);

    //    //返回false验证失败
    //    public override bool IsValid(object value, out ErrorInfo errorInfo)
    //    {
    //        string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
    //        errorInfo = null;
    //        if (!noSpaceRegex.IsMatch(stringValue))
    //            errorInfo = new ErrorInfo { ErrorMsg = "不能包含空格" };
    //        return errorInfo == null;
    //    }
    //}

    ///// <summary>
    ///// 最大长度验证
    ///// </summary>
    //public class StrMaxLenAttribute : ExcelAttribute
    //{
    //    public StrMaxLenAttribute(int minLength, int maxLength)
    //    {
    //        maximumLength = maxLength;
    //        minimumLength = minLength;
    //    }
    //    /// <summary>
    //    /// 最大长度
    //    /// </summary>
    //    public int maximumLength { get; set; }
    //    /// <summary>
    //    /// 最小长度
    //    /// </summary>
    //    public int minimumLength { get; set; }

    //    public override bool IsValid(object value, out ErrorInfo errorInfo)
    //    {
    //        errorInfo = null;
    //        string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
    //        if (!(stringValue.Length <= maximumLength && stringValue.Length >= minimumLength))
    //            errorInfo = new ErrorInfo { ErrorMsg = $"限制最小长度{minimumLength},最大长度{maximumLength}" };
    //        return errorInfo == null;
    //    }
    //}

    ///// <summary>
    ///// 数据库验证是否存在
    ///// </summary>
    //public class DatabaseHaveAttribute : DatabaseAttribute
    //{
    //    public DatabaseHaveAttribute(string databaseData, Type type) : base(databaseData, type) { }

    //    public override bool IsValid(List<object> values, out List<ErrorInfo> errorInfos)
    //    {
    //        //加载数据库数据
    //        List<DbResult> dbResults = LoadData(values);
    //        errorInfos = new List<ErrorInfo>();
    //        //检查数据加载是否正常
    //        if (base.errorInfo != null)
    //        {
    //            errorInfos.Add(base.errorInfo);
    //            return false;
    //        }
    //        int row = 0;
    //        //数据验证
    //        foreach (string value in values)
    //        {
    //            row++;
    //            if (!dbResults.Any(res => res.Value == value))
    //            {
    //                errorInfos.Add(new ErrorInfo
    //                {
    //                    Row = row,
    //                    ColumnName = databaseData.Column,
    //                    TableName = databaseData.Table,
    //                    ErrorMsg = "数据库不存在此数据"
    //                });
    //            }
    //        }
    //        return errorInfos.Count == 0;
    //    }
    //}

}
