using System.Collections.Generic;

namespace ExcelVerify
{
    #region 异常类
    /// <summary>
    /// 错误对象
    /// </summary>
    public class ErrorRowInfo
    {
        /// <summary>
        /// 行
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// 错误信息
        /// </summary>
        public List<ErrorInfo> ErrorInfos { get; set; }
        /// <summary>
        /// 错误描述
        /// </summary>
        public string Massage { get; set; }
        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }
    }
    /// <summary>
    /// 错误对象
    /// </summary>
    public class ErrorInfo
    {
        /// <summary>
        /// 行
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// 列
        /// </summary>
        public int Column { get; set; }
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// 表名
        /// </summary>
        public string TableName { get; set; }
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMsg { get; set; }
    }

    #endregion
}
