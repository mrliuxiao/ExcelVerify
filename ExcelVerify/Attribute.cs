using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Dapper;
using Org.BouncyCastle.Asn1.Cmp;

namespace ExcelVerify
{
    /// <summary>
    /// 不能为null、""
    /// </summary>
    public class NotNullAttribute : Attribute
    {
        //返回true表示此字段通过验证
        public bool IsValid(object value, out ErrorInfo errorInfo)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            errorInfo = null;
            if (!string.IsNullOrEmpty(stringValue))
                errorInfo = new ErrorInfo { ErrorMsg = "数据库不存在此数据" };
            return errorInfo == null;
        }
    }

    /// <summary>
    /// 字符串中间不能包含空格
    /// </summary>
    public class NoSpaceAttribute : Attribute
    {
        private static readonly Regex noSpaceRegex = new Regex(@"^[^\s]+$", RegexOptions.Compiled);

        //返回false验证失败
        public bool IsValid(object value, out ErrorInfo errorInfo)
        {
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            errorInfo = null;
            if (!noSpaceRegex.IsMatch(stringValue))
                errorInfo = new ErrorInfo { ErrorMsg = "不能包含空格" };
            return errorInfo == null;
        }
    }

    /// <summary>
    /// 最大长度验证
    /// </summary>
    public class StrMaxLenAttribute : Attribute
    {
        public StrMaxLenAttribute(int minLength, int maxLength)
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

        public bool IsValid(object value, out ErrorInfo errorInfo)
        {
            errorInfo = null;
            string stringValue = Convert.ToString(value, CultureInfo.CurrentCulture);
            if (!(stringValue.Length <= maximumLength && stringValue.Length >= minimumLength))
                errorInfo = new ErrorInfo { ErrorMsg = $"限制最小长度{minimumLength},最大长度{maximumLength}" };
            return errorInfo == null;
        }
    }

    /// <summary>
    /// 数据库验证
    /// </summary>
    public class DatabaseAttribute : Attribute
    {
        public DatabaseData databaseData = null;
        private readonly string keyName = "";
        private readonly Type verifyConfig = null;
        public ErrorInfo errorInfo = null;

        public DatabaseAttribute(string key, Type config)
        {
            keyName = key;
            verifyConfig = config;
        }

        public List<DbResult> LoadData(List<object> values)
        {
            object Instance = Activator.CreateInstance(verifyConfig);
            string connect = verifyConfig.GetField(keyName).GetValue(Instance).ToString();
            //解析
            databaseData = DatabaseData.ConverToObject(connect);

            if (databaseData == null || string.IsNullOrEmpty(databaseData.ConnectionString))
            {
                ErrorInfo errorInfo = new ErrorInfo
                {
                    ColumnName = databaseData.Column,
                    TableName = databaseData.Table
                };
                errorInfo.ErrorMsg = "数据库连接配置错误";
            }
            List<DbResult> dbResult = null;
            //获取表名
            using (SqlConnection connection = new SqlConnection(databaseData.ConnectionString))
            {
                string where = "";
                values.ForEach(value => where += $" {databaseData.Column} = '{value}' {(values.Last() == value ? "" : "OR")} ");

                string Sql = string.Format($@"SELECT {databaseData.PrimaryKey} AS PrimaryKey, {databaseData.Column} AS Value FROM {databaseData.Table} WHERE {where};");
                dbResult = connection.Query<DbResult>(Sql).ToList();
            }
            return dbResult;
        }


    }

    /// <summary>
    /// 数据库验证是否存在
    /// </summary>
    public class DatabaseHaveAttribute : DatabaseAttribute
    {
        public DatabaseHaveAttribute(string databaseData, Type type) : base(databaseData, type) { }

        public bool IsValid(List<object> values, out List<ErrorInfo> errorInfos)
        {
            //加载数据库数据
            List<DbResult> dbResults = LoadData(values);
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
                        ErrorMsg = "数据库不存在此数据"
                    });
                }
            }
            return errorInfos.Count == 0;
        }
    }

}
