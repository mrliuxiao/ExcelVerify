using System.Collections.Generic;
using System.Linq;

namespace ExcelVerify
{
    /// <summary>
    /// 数据库信息
    /// </summary>
    public class DatabaseData
    {
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public string ConnectionString { get; set; }
        /// <summary>
        /// 主键
        /// </summary>
        public string PrimaryKey { get; set; }
        /// <summary>
        /// 列名
        /// </summary>
        public string Column { get; set; }
        /// <summary>
        /// 列名
        /// </summary>
        public string Table { get; set; }
        /// <summary>
        /// 地址
        /// </summary>
        public string Server { get; set; }
        /// <summary>
        /// 数据库
        /// </summary>
        public string Database { get; set; }
        /// <summary>
        /// 用户名
        /// </summary>
        public string UID { get; set; }
        /// <summary>
        /// 密码
        /// </summary>
        public string PassWord { get; set; }

        public static DatabaseData ConverToObject(string str)
        {
            List<string> dicList = str.Split(';').ToList();
            DatabaseData databaseData = new DatabaseData();

            foreach (string dicStr in dicList)
            {
                string[] keyValue = dicStr.Split('=');
                string key = keyValue[0].ToLower();
                string value = keyValue[1];
                switch (key)
                {
                    case "server":
                        databaseData.Server = value;
                        databaseData.ConnectionString += $"{key}={value};";
                        break;
                    case "database":
                        databaseData.Database = value;
                        databaseData.ConnectionString += $"{key}={value};";
                        break;
                    case "uid":
                        databaseData.UID = value;
                        databaseData.ConnectionString += $"{key}={value};";
                        break;
                    case "pwd":
                        databaseData.PassWord = value;
                        databaseData.ConnectionString += $"{key}={value};";
                        break;
                    case "table":
                        databaseData.Table = value;
                        break;
                    case "primarykey":
                        databaseData.PrimaryKey = value;
                        break;
                    case "column":
                        databaseData.Column = value;
                        break;
                    default:
                        break;
                }
            }

            return databaseData;

        }
    }

}
