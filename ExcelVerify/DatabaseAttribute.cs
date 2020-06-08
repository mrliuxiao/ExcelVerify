using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Dapper;

namespace ExcelVerify
{
    /// <summary>
    /// 数据库验证
    /// </summary>
    public class DatabaseAttribute : ExcelAttribute
    {
        public DatabaseData databaseData = null;
        private readonly string keyName = "";
        private readonly Type verifyConfig = null;
        public ErrorInfo errorInfo = null;

        public DatabaseAttribute(string entityProperty, Type databaseConfig)
        {
            keyName = entityProperty;
            verifyConfig = databaseConfig;
        }

        public List<DbResult> LoadData(List<object> values)
        {
            object Instance = Activator.CreateInstance(verifyConfig);
            var field = verifyConfig.GetField(keyName);
            if (field == null)
                throw new Exception($"{keyName},没有配置数据库信息");
            string connect = field.GetValue(Instance).ToString();
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


        //返回true表示此字段通过验证
        public virtual bool IsValid(List<object> values, out List<ErrorInfo> errorInfos, out List<DbResult> dbResults)
        {
            errorInfos = null;
            dbResults = null;
            return errorInfos == null;
        }
    }

}
