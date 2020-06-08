using NPOI.POIFS.Storage;
using NPOI.SS.Formula.Functions;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelVerify
{
    /// <summary>
    /// Excel数据效验帮助类
    /// </summary>
    public class ExcelVerifiyHelper
    {

        #region 效验Excel
        /// <summary>
        /// 效验Excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        private DataTable VerifyToDataTable(string path, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows)
        {


            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验Excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        private DataTable VerifyToDataTable(string path, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验Excel
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="path"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public List<TEntity> VerifyToEntitys<TEntity>(string path, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            Stream stream = File.OpenRead(path);
            DataTable dataTable = NPOIHelper.RenderDataTableFromExcel(stream, 0, 0);
            return VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
        }

        /// <summary>
        /// 效验Excel
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="path"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public List<TEntity> VerifyToEntitys<TEntity>(string path, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            Stream stream = File.OpenRead(path);
            DataTable dataTable = NPOIHelper.RenderDataTableFromExcel(stream, 0, 0);
            return VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
        }
        #endregion

        #region 效验DataTable
        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="dataTable">效验数据表</param>
        /// <param name="mapConfigs">列映射配置</param>
        /// <param name="errorRows">错误信息</param>
        /// <returns>效验通过的实体集合</returns>
        public static List<TEntity> VerifyToEntitys<TEntity>(DataTable dataTable, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<ErrorRowInfo> errorRowInfos = new List<ErrorRowInfo>();
            List<ErrorInfo> errorInfos = new List<ErrorInfo>();
            int columnNum = 0;
            List<DataRow> errorDtRow = new List<DataRow>();
            //映射列名
            DataTable mapDataTable = PropertyMapToDataTable(dataTable, mapConfigs);
            //按照列验证 为了性能
            foreach (DataColumn dataColumn in mapDataTable.Columns)
            {
                List<DbResult> dbResults = null;
                columnNum++;
                List<object> values = new List<object>();

                //验证所有列的数据是否有效
                foreach (DataRow dataRow in mapDataTable.Rows)
                {
                    values.Add(dataRow[dataColumn]);
                }
                errorInfos.InsertRange(0, EntityAttributeValueValid<TEntity>(values, dataColumn.ColumnName, columnNum, out dbResults));
                //替换外键数据为Id

                if (dbResults != null)
                {
                    foreach (DataRow row in mapDataTable.Rows)
                    {
                        DbResult dbResult = dbResults.Find(res => res.Value == row[dataColumn].ToString());
                        if (dbResult != null)
                            row[dataColumn] = dbResult.PrimaryKey;
                    }
                }
            }
            //映射DataTable的列明到错误信息
            errorInfos.ForEach(s =>
            {
                MapConfig mapConfig = mapConfigs.Find(cfg => cfg.EntityColumnName == s.ColumnName);
                if (mapConfig != null)
                    s.ColumnName = mapConfig.DataTableColumnName;
            });
            //重新排列数据
            errorRowInfos = errorInfos.GroupBy(info => info.Row).OrderBy(o => o.Key).Select(s =>
            {
                errorDtRow.Add(mapDataTable.Rows[s.Key - 1]);
                return new ErrorRowInfo
                {
                    Row = s.Key,
                    ErrorInfos = s.OrderBy(order => order.Column).ToList(),
                    IsValid = s.Count() == 0,
                    Massage = "未通过效验"
                };
            }).ToList();
            //筛选出验证通过的数据
            errorDtRow.ForEach(row => { mapDataTable.Rows.Remove(row); });
            //转换验证通过的数据为实体
            var VerifiedEntities = ConvertToEntity<TEntity>(mapDataTable);
            errorRows = errorRowInfos;
            return VerifiedEntities;
        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="dataTable">效验数据表</param>
        /// <param name="mapConfigs">列映射配置</param>
        /// <param name="errorRows">错误信息</param>
        /// <returns>效验通过的DataTable</returns>
        private static DataTable VerifyToDataTable<TEntity>(DataTable dataTable, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<TEntity> entities = VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
            return ConvertMapToDataTable(entities, mapConfigs);

        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="dataTable">效验数据表</param>
        /// <param name="mapConfig">列映射配置</param>
        /// <param name="errorRows">错误信息</param>
        /// <returns>效验通过的实体集合</returns>
        public static List<TEntity> VerifyToEntitys<TEntity>(DataTable dataTable, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {

            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            return VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);

        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="dataTable">效验数据表</param>
        /// <param name="mapConfig">列映射配置</param>
        /// <param name="errorRows">错误信息</param>
        /// <returns>效验通过的DataTable</returns>
        private static DataTable VerifyToDataTable<TEntity>(DataTable dataTable, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            List<TEntity> entities = VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
            return ConvertMapToDataTable(entities, mapConfigs);

        }

        #endregion

        #region 效验实体集合
        /// <summary>
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="mapConfigs"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public static List<TEntity> VerifyToEntitys<TEntity>(List<TEntity> entities, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            DataTable dataTable = ConvertMapToDataTable(entities, mapConfigs);
            return VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
        }

        /// <summary>
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="mapConfigs"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        private DataTable VerifyToDataTable<TEntity>(List<TEntity> entities, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            DataTable dataTable = ConvertMapToDataTable(entities, mapConfigs);
            return VerifyToDataTable<TEntity>(dataTable, mapConfigs, out errorRows);
        }

        /// <summary>
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="mapConfig"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public static List<TEntity> VerifyToEntitys<TEntity>(List<TEntity> entities, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            DataTable dataTable = ConvertMapToDataTable(entities, mapConfigs);
            return VerifyToEntitys<TEntity>(dataTable, mapConfigs, out errorRows);
        }


        /// <summary>
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="mapConfig"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        private DataTable VerifyToDataTable<TEntity>(List<TEntity> entities, DesignDelegate mapConfig, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            DataTable dataTable = ConvertMapToDataTable(entities, mapConfigs);
            return VerifyToDataTable<TEntity>(dataTable, mapConfigs, out errorRows);
        }
        #endregion

        #region 核心效验方法

        /// <summary>
        /// 根据实体标记的特性进行数据效验
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="values"></param>
        /// <param name="columnName"></param>
        /// <param name="columnNum"></param>
        /// <param name="dbResults"></param>
        /// <returns></returns>
        private static List<ErrorInfo> EntityAttributeValueValid<TEntity>(List<object> values, string columnName, int columnNum, out List<DbResult> dbResults)
        {
            //获取实体类型
            Type entityType = typeof(TEntity);

            List<ErrorInfo> errorInfos = new List<ErrorInfo>();
            PropertyInfo property = entityType.GetProperty(columnName);

            if (property == null)
                throw new Exception($"'{columnName}'列名与实体不匹配，请检查是否需要映射");
            List<Attribute> attributes = property.GetCustomAttributes().ToList();
            attributes = attributes.FindAll(attr =>
                 attr.GetType().GetMethod("IsValid", new Type[] { typeof(object), typeof(ErrorInfo).MakeByRefType(), typeof(List<DbResult>).MakeByRefType() }) != null ||
                 attr.GetType().GetMethod("IsValid", new Type[] { typeof(object), typeof(ErrorInfo).MakeByRefType() }) != null);
            List<Attribute> commonAttributes = new List<Attribute>();
            if (attributes != null)
                commonAttributes = attributes.FindAll(attr => !attr.GetType().Name.Contains("Database"));
            List<Attribute> databaseAttributes = new List<Attribute>();
            if (attributes != null)
                databaseAttributes = attributes.FindAll(attr => attr.GetType().Name.Contains("Database"));

            int rowNum = 0;
            //每行验证
            foreach (object value in values)
            {
                rowNum++;
                var propertyValue = value == System.DBNull.Value ? "" : value;

                //遍历属性标记
                foreach (Attribute attribute in commonAttributes)
                {
                    //验证所有属性
                    ErrorInfo errorInfo = new ErrorInfo();
                    object[] invokeArgs = new object[] { propertyValue, errorInfo };
                    bool isValid = (bool)attribute.GetType().GetMethod("IsValid", new Type[] { typeof(object), typeof(ErrorInfo).MakeByRefType() }).Invoke(attribute, invokeArgs);
                    if (!isValid)
                    {
                        errorInfo = (ErrorInfo)invokeArgs[1];
                        errorInfo.ColumnName = columnName;
                        errorInfo.Column = columnNum;
                        errorInfo.Row = rowNum;
                        errorInfos.Add(errorInfo);
                    }
                }
            }
            dbResults = null;
            //去数据库验证 一次性验证所有列
            foreach (Attribute attribute in databaseAttributes)
            {
                //验证所有属性
                List<ErrorInfo> errors = new List<ErrorInfo>();
                object[] invokeArgs = new object[] { values, errors, dbResults };
                bool isValid = (bool)attribute.GetType().GetMethod("IsValid", new Type[] { typeof(List<object>), typeof(List<ErrorInfo>).MakeByRefType(), typeof(List<DbResult>).MakeByRefType() }).Invoke(attribute, invokeArgs);
                dbResults = (List<DbResult>)invokeArgs[2];
                if (!isValid)
                {
                    errors = (List<ErrorInfo>)invokeArgs[1];
                    errors.ForEach(info => info.Column = columnNum);
                    errorInfos.InsertRange(0, errors);
                }
            }
            return errorInfos;
        }

        #endregion

        #region Convert DataTable与Entity转换类
        /// <summary>
        /// 映射实体属性名到DataTable
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="mapConfigs">映射</param>
        /// <returns></returns>
        public static DataTable PropertyMapToDataTable(DataTable dataTable, List<MapConfig> mapConfigs)
        {

            DataTable newDataTable = new DataTable();
            //映射列
            foreach (DataColumn dataColumn in dataTable.Columns)
            {
                //映射实体字段到Table
                MapConfig mapConfig = mapConfigs.Find(cfg => cfg.DataTableColumnName == dataColumn.ColumnName);
                string dtColumnName = mapConfig == null ? dataColumn.ColumnName : mapConfig.EntityColumnName;
                newDataTable.Columns.Add(dtColumnName, dataColumn.DataType);
            }

            //插入数据
            foreach (DataRow dataRow in dataTable.Rows)
            {
                newDataTable.Rows.Add(dataRow.ItemArray);
            }
            return newDataTable;
        }

        /// <summary>
        /// 通过映射配置转换DataTable为实体
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable">数据表</param>
        /// <param name="mapConfigs">映射</param>
        /// <returns>映射配置转换后的实体</returns>
        public static List<TEntity> ConvertMapToEntity<TEntity>(DataTable dataTable, List<MapConfig> mapConfigs) where TEntity : new()
        {

            var properts = typeof(TEntity).GetProperties(); ;
            List<TEntity> entities = new List<TEntity>();

            foreach (DataRow dataRow in dataTable.Rows)
            {

                TEntity entity = new TEntity();

                foreach (PropertyInfo property in properts)
                {
                    //映射实体字段到Table
                    MapConfig mapConfig = mapConfigs.Find(cfg => cfg.EntityColumnName == property.Name);
                    string dtColumnName = mapConfig == null ? property.Name : mapConfig.EntityColumnName;

                    if (dataRow.Table.Columns.Contains(dtColumnName))
                    {
                        try
                        {
                            if (!Convert.IsDBNull(dataRow[dtColumnName]))
                            {
                                object v = null;
                                if (property.PropertyType.ToString().Contains("System.Nullable"))
                                {
                                    v = Convert.ChangeType(dataRow[dtColumnName], Nullable.GetUnderlyingType(property.PropertyType));
                                }
                                else
                                {
                                    v = Convert.ChangeType(dataRow[dtColumnName], property.PropertyType);
                                }
                                property.SetValue(entity, v, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"字段[{ dtColumnName }]数据类型出错,{ex.Message}");

                        }
                    }
                }
                entities.Add(entity);
            }
            return entities;
        }

        /// <summary>
        /// 通过映射配置转换DataTable为实体
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable">数据表</param>
        /// <param name="mapConfig">映射配置</param>
        /// <returns>映射配置转换后的实体</returns>
        public static List<TEntity> ConvertMapToEntity<TEntity>(DataTable dataTable, DesignDelegate mapConfig) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            return ConvertMapToEntity<TEntity>(dataTable, mapConfigs);
        }

        /// <summary>
        /// 转换DataTable为实体
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public static List<TEntity> ConvertToEntity<TEntity>(DataTable dataTable) where TEntity : new()
        {

            var properts = typeof(TEntity).GetProperties(); ;
            List<TEntity> entities = new List<TEntity>();

            foreach (DataRow dataRow in dataTable.Rows)
            {

                TEntity entity = new TEntity();

                foreach (PropertyInfo property in properts)
                {
                    string dtColumnName = property.Name;
                    if (dataRow.Table.Columns.Contains(dtColumnName))
                    {
                        try
                        {
                            if (!Convert.IsDBNull(dataRow[dtColumnName]))
                            {
                                object v = null;
                                if (property.PropertyType.ToString().Contains("System.Nullable"))
                                {
                                    v = Convert.ChangeType(dataRow[dtColumnName], Nullable.GetUnderlyingType(property.PropertyType));
                                }
                                else
                                {
                                    v = Convert.ChangeType(dataRow[dtColumnName], property.PropertyType);
                                }
                                property.SetValue(entity, v, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"字段[{ dtColumnName }]数据类型出错,请检查是否没有映射主外键数据！{ex.Message}");
                        }
                    }
                }
                entities.Add(entity);
            }
            return entities;
        }

        /// <summary>
        /// 转换实体为DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <returns></returns>
        public static DataTable ConvertToDataTable<TEntity>(List<TEntity> entities)
        {
            var result = DataTableHelper.CreateTable<TEntity>();
            DataTableHelper.FillData(result, entities);
            return result;
        }

        /// <summary>
        /// 通过映射配置转换实体为DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="mapConfigs"></param>
        /// <returns>映射列名后的DataTable</returns>
        public static DataTable ConvertMapToDataTable<TEntity>(List<TEntity> entities, List<MapConfig> mapConfigs)
        {
            var dataTable = DataTableHelper.CreateTable<TEntity>();

            foreach (var entity in entities)
            {
                DataRow row = dataTable.NewRow();
                var type = typeof(T);
                foreach (var property in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
                {
                    MapConfig config = mapConfigs.Find(cfg => cfg.EntityColumnName == property.Name);
                    string rowName = config == null ? property.Name : config.DataTableColumnName;
                    row[rowName] = property.GetValue(entity) ?? DBNull.Value;
                }
            }
            return dataTable;
        }

        /// <summary>
        /// 通过映射配置转换实体为DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities">实体集合</param>
        /// <param name="mapConfig">映射配置</param>
        /// <returns>映射列名后的DataTable</returns>
        public static DataTable ConvertMapToDataTable<TEntity>(List<TEntity> entities, DesignDelegate mapConfig) where TEntity : new()
        {
            List<MapConfig> mapConfigs = new List<MapConfig>();
            AddMapConfig AddMapConfig = new AddMapConfig();
            mapConfig(AddMapConfig);
            mapConfigs = AddMapConfig.mapConfigs;
            return ConvertMapToDataTable(entities, mapConfigs);
        }
        #endregion
    }
}
