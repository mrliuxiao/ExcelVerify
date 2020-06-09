using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Org.BouncyCastle.Asn1.Cmp;

namespace ExcelVerify
{

    #region 空验证 

    /// <summary>
    /// 不能为null、""
    /// </summary>
    public class ExcelNoNullAttribute : ExcelAttribute
    {
        /// <summary>
        /// 效验效验数据 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public override bool IsValid(object value)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            return !string.IsNullOrEmpty(stringValue);
        }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="value">效验数据</param>
        /// <param name="errorInfo">错误信息</param>
        /// <returns></returns>
        public override bool IsValid(object value, out ErrorInfo errorInfo)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            errorInfo = null;
            if (string.IsNullOrEmpty(stringValue))
                errorInfo = new ErrorInfo { ErrorData = value.ToString(), ErrorMsg = "不能为空" };
            return errorInfo == null;
        }
    }

    /// <summary>
    /// 字符串中间不能包含空格
    /// </summary>
    public class ExcelNoSpaceAttribute : ExcelAttribute
    {
        private static readonly Regex noSpaceRegex = new Regex(@"^[^\s]+$", RegexOptions.Compiled);

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="value">效验数据</param>
        /// <returns></returns>
        public override bool IsValid(object value)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            if (string.IsNullOrEmpty(stringValue))
            {
                return true;
            }
            return noSpaceRegex.IsMatch(stringValue);
        }
        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="value">效验数据</param>
        /// <param name="errorInfo">错误信息</param>
        /// <returns></returns>
        public override bool IsValid(object value, out ErrorInfo errorInfo)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            errorInfo = null;
            if (!noSpaceRegex.IsMatch(stringValue))
                errorInfo = new ErrorInfo { ErrorData = value.ToString(), ErrorMsg = $"'{value}'不能包含空格" };
            return errorInfo == null;
        }
    }

    #endregion

    #region 字符串长度验证

    /// <summary>
    /// 长度验证
    /// </summary>
    public class ExcelStrLengthAttribute : ExcelAttribute
    {
        /// <summary>
        /// 长度验证
        /// </summary>
        /// <param name="minLength">最小长度</param>
        /// <param name="maxLength">最大长度</param>
        public ExcelStrLengthAttribute(int minLength, int maxLength)
        {
            maximumLength = maxLength;
            minimumLength = minLength;
        }
        /// <summary>
        /// 最大长度
        /// </summary>
        public int maximumLength { get; set; }
        /// <summary>
        /// 最小长度
        /// </summary>
        public int minimumLength { get; set; }
        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="value">效验数据</param>
        /// <returns></returns>
        public override bool IsValid(object value)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            return stringValue.Length <= maximumLength && stringValue.Length >= minimumLength;
        }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="value">效验数据</param>
        /// <param name="errorInfo">错误信息</param>
        /// <returns></returns>
        public override bool IsValid(object value, out ErrorInfo errorInfo)
        {
            errorInfo = null;
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            if (!(stringValue.Length <= maximumLength && stringValue.Length >= minimumLength))
                errorInfo = new ErrorInfo { ErrorData = value.ToString(), ErrorMsg = $"'{value}'限制最小长度{minimumLength},最大长度{maximumLength}" };
            return errorInfo == null;
        }
    }
    #endregion

    /// <summary>
    /// 不能与数据库重复
    /// </summary>
    public class ExcelNoRepeatToDatabaseAttribute : DatabaseAttribute
    {
        /// <summary>
        /// 传参
        /// </summary>
        /// <param name="entityProperty">实体属性名</param>
        /// <param name="databaseConfig">数据库连接字符串等信息</param>
        public ExcelNoRepeatToDatabaseAttribute(string entityProperty, Type databaseConfig) : base(entityProperty, databaseConfig) { }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="values">效验数据集</param>
        /// <param name="errorInfos">错误信息</param>
        /// <param name="dbResults">数据库查询结果</param>
        /// <returns></returns>
        public override bool IsValid(List<object> values, out List<ErrorInfo> errorInfos, out List<DbResult> dbResults)
        {
            //加载数据库数据 
            dbResults = LoadData(values);
            errorInfos = new List<ErrorInfo>();
            //List<string> totalData = values.Select(v => v.ToString()).ToList();
            //totalData.InsertRange(0, dbResults.Select(res => res.Value).ToList());
            int row = 0;
            //验证是否重复
            foreach (string value in values)
            {
                var dbRes = dbResults.Find(db => db.Value == value);
                //检查数据加载是否正常
                if (base.errorInfo != null)
                {
                    errorInfos.Add(base.errorInfo);
                    return false;
                }
                //验证
                if (dbRes != null)
                {
                    errorInfos.Add(new ErrorInfo
                    {
                        Row = row,
                        ColumnName = databaseData.Column,
                        TableName = databaseData.Table,
                        ErrorData = value,
                        ErrorMsg = $"'{value}'数据重复"
                    });
                }
                int conut = values.Count(v => v.ToString() == value);
                if (conut > 1)
                {
                    errorInfos.Add(new ErrorInfo
                    {
                        Row = row,
                        ColumnName = databaseData.Column,
                        TableName = databaseData.Table,
                        ErrorData = value,
                        ErrorMsg = $"'{value}'数据重复"
                    });
                }
            }

            //不要映射主键ID
            dbResults = null;
            return errorInfos.Count == 0;
        }
    }


    /// <summary>
    /// 数据库验证是否存在 不获取主键Id
    /// </summary>
    public class ExcelDatabaseHaveAttribute : DatabaseAttribute
    {
        /// <summary>
        /// 数据库是否存在
        /// </summary>
        /// <param name="entityProperty">实体属性</param>
        /// <param name="databaseConfig">数据库配置</param>
        public ExcelDatabaseHaveAttribute(string entityProperty, Type databaseConfig) : base(entityProperty, databaseConfig) { }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="values">数据集</param>
        /// <param name="errorInfos">错误信息</param>
        /// <param name="dbResults">查询数据结果</param>
        /// <returns></returns>
        public override bool IsValid(List<object> values, out List<ErrorInfo> errorInfos, out List<DbResult> dbResults)
        {
            //加载数据库数据 
            dbResults = LoadData(values);
            errorInfos = new List<ErrorInfo>();
            //检查数据加载是否正常
            if (base.errorInfo != null)
            {
                errorInfos.Add(base.errorInfo);
                return false;
            }
            int row = 0;
            //数据验证
            foreach (string value in values)
            {
                row++;
                if (!dbResults.Any(res => res.Value == value))
                {
                    errorInfos.Add(new ErrorInfo
                    {
                        Row = row,
                        ColumnName = databaseData.Column,
                        TableName = databaseData.Table,
                        ErrorData = value,
                        ErrorMsg = $"'{value}'数据库不存在此数据"
                    });
                }
            }
            //dbResults.InsertRange(0, results);
            //如果不想映射主键数据 清空即可
            dbResults = null;
            return errorInfos.Count == 0;
        }
    }

    /// <summary>
    /// 数据库验证是否存在 并获取主键
    /// </summary>
    public class ExcelDatabaseHaveGetKeyAttribute : DatabaseAttribute
    {
        /// <summary>
        /// 数据库验证是否存在 并获取主键
        /// </summary>
        /// <param name="entityProperty">实体属性</param>
        /// <param name="databaseConfig">数据库配置</param>
        public ExcelDatabaseHaveGetKeyAttribute(string entityProperty, Type databaseConfig) : base(entityProperty, databaseConfig) { }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="values">数据集</param>
        /// <param name="errorInfos">错误信息</param>
        /// <param name="dbResults">查询数据结果</param>
        /// <returns></returns>
        public override bool IsValid(List<object> values, out List<ErrorInfo> errorInfos, out List<DbResult> dbResults)
        {
            //加载数据库数据 
            dbResults = LoadData(values);
            errorInfos = new List<ErrorInfo>();
            //检查数据加载是否正常
            if (base.errorInfo != null)
            {
                errorInfos.Add(base.errorInfo);
                return false;
            }
            int row = 0;
            //数据验证
            foreach (string value in values)
            {
                row++;
                if (!dbResults.Any(res => res.Value == value))
                {
                    errorInfos.Add(new ErrorInfo
                    {
                        Row = row,
                        ColumnName = databaseData.Column,
                        TableName = databaseData.Table,
                        ErrorData = value,
                        ErrorMsg = $"'{value}'数据库不存在此数据"
                    });
                }
            }
            //如果不想映射主键数据 清空即可
            //dbResults = null;
            return errorInfos.Count == 0;
        }
    }

    /// <summary>
    /// 数据库验证是否存在 并获取主键 逗号 分隔转 Ids
    /// </summary>
    public class ExcelDatabaseHaveGetIdsAttribute : DatabaseAttribute
    {
        /// <summary>
        /// 数据库验证是否存在 并获取主键
        /// </summary>
        /// <param name="entityProperty">实体属性</param>
        /// <param name="databaseConfig">数据库配置</param>
        public ExcelDatabaseHaveGetIdsAttribute(string entityProperty, Type databaseConfig) : base(entityProperty, databaseConfig) { }

        /// <summary>
        /// 效验
        /// </summary>
        /// <param name="values">数据集</param>
        /// <param name="errorInfos">错误信息</param>
        /// <param name="dbResults">查询数据结果</param>
        /// <returns></returns>
        public override bool IsValid(List<object> values, out List<ErrorInfo> errorInfos, out List<DbResult> dbResults)
        {
            List<string> datas = new List<string>();
            foreach (string value in values)
            {
                //Value格式 "1,2,3"
                datas.InsertRange(0, value.Split(',').ToList());
            }

            //加载数据库数据 
            dbResults = LoadData(datas);
            errorInfos = new List<ErrorInfo>();
            //检查数据加载是否正常
            if (base.errorInfo != null)
            {
                errorInfos.Add(base.errorInfo);
                return false;
            }
            List<DbResult> newDBResults = new List<DbResult>();

            int row = 0;
            //数据验证
            foreach (string value in values)
            {
                DbResult dbResult = new DbResult();
                dbResult.Value = value;
                //Value格式 "1,2,3"
                List<string> valueList = value.Split(',').ToList();
                row++;
                int i= 0;
                foreach (string data in valueList)
                {
                    i++;
                    if (!dbResults.Any(res => res.Value == data))
                    {
                        errorInfos.Add(new ErrorInfo
                        {
                            Row = row,
                            ColumnName = databaseData.Column,
                            TableName = databaseData.Table,
                            ErrorData = value,
                            ErrorMsg = $"'{value}'数据库不存在此数据"
                        });
                    }
                    else
                    {
                        dbResult.PrimaryKey += $"{dbResults.Find(res => res.Value == data).PrimaryKey}{(valueList.Count() == i ? "" : ",")}";
                    }
                }
                newDBResults.Add(dbResult);
            }
            //如果不想映射主键数据 清空即可
            dbResults = newDBResults;
            return errorInfos.Count == 0;
        }
    }
}
