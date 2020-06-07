using NPOI.POIFS.Storage;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelVerify
{
    public class ExcelVerifiyHelper
    {
        ///// <summary>
        ///// 配置委托
        ///// </summary>
        ///// <typeparam name="T"></typeparam>
        ///// <param name="arg"></param>
        //public delegate void DesignDelegate<in T>(T arg);

        /// <summary>
        /// 效验Excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        public DataTable VerifyToDataTable(string path, out List<ErrorRowInfo> errorRows)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        public DataTable VerifyToDataTable(DataTable dataTable, out List<ErrorRowInfo> errorRows)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的DataTable</returns>
        public DataTable VerifyToDataTable<TEntity>(List<TEntity> entities, out List<ErrorRowInfo> errorRows)
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
        public List<TEntity> VerifyToEntitys<TEntity>(string path, out List<ErrorRowInfo> errorRows)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public static List<TEntity> VerifyToEntitys<TEntity>(DataTable dataTable, List<MapConfig> mapConfigs, out List<ErrorRowInfo> errorRows) where TEntity : new()
        {
            List<ErrorRowInfo> errorRowInfos = new List<ErrorRowInfo>();
            List<ErrorInfo> errorInfos = new List<ErrorInfo>();
            int columnNum = 0;
            List<DataRow> errorDtRow = new List<DataRow>();
            //映射列名
            DataTable mapDataTable = MapToDataTable(dataTable, mapConfigs);

            //按照列验证 为了性能
            foreach (DataColumn dataColumn in mapDataTable.Columns)
            {
                columnNum++;
                List<object> values = new List<object>();

                //验证所有列的数据是否有效
                foreach (DataRow dataRow in mapDataTable.Rows)
                {
                    values.Add(dataRow[dataColumn]);
                }
                errorInfos.InsertRange(0, EntityAttributeValueValid<TEntity>(values, dataColumn.ColumnName, columnNum));
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
                    ErrorInfos = s.ToList(),
                    IsValid = s.Count() == 0,
                    Massage = "未通过效验"
                };
            }).ToList();

            //筛选出验证通过的数据
            errorDtRow.ForEach(row =>
            {
                mapDataTable.Rows.Remove(row);
            });
            //转换验证通过的数据为实体
            var VerifiedEntities = ConvertToEntity<TEntity>(mapDataTable);
            errorRows = errorRowInfos;
            return VerifiedEntities;
        }

        /// <summary>
        /// 效验DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="errorRows"></param>
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
        /// 效验实体集合
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="errorRows"></param>
        /// <returns>效验通过的实体集合</returns>
        public List<TEntity> VerifyToEntitys<TEntity>(List<TEntity> entities, out List<ErrorRowInfo> errorRows)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 效验属性值有效性
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entity"></param>
        /// <param name="func">回调效验方法 返回NULL 或 "" 则效验成功</param>
        /// <returns></returns>
        public static List<ErrorInfo> EntityAttributeValueValid<TEntity>(List<object> values, string columnName, int columnNum)
        {

            //获取实体类型
            Type entityType = typeof(TEntity);

            List<ErrorInfo> errorInfos = new List<ErrorInfo>();
            PropertyInfo property = entityType.GetProperty(columnName);

            if (property == null)
                throw new Exception("列名与实体不匹配，请检查是否需要映射");
            List<Attribute> attributes = property.GetCustomAttributes().ToList();
            int rowNum = 0;
            //每行验证
            foreach (object value in values)
            {
                rowNum++;
                var propertyValue = value == System.DBNull.Value ? "" : value;
                //遍历属性标记
                foreach (Attribute attribute in attributes)
                {
                    //判断是否有非空验证属性
                    if (attribute.GetType().IsAssignableFrom(typeof(NotNullAttribute)))
                    {
                        NotNullAttribute attr = ((NotNullAttribute)attribute);
                        ErrorInfo errorInfo = null;
                        if (!attr.IsValid(propertyValue, out errorInfo))
                        {
                            errorInfo.ColumnName = columnName;
                            errorInfo.Column = columnNum;
                            errorInfo.Row = rowNum;
                            errorInfos.Add(errorInfo);
                        }
                    }
                    //判断是否有长度验证属性
                    if (attribute.GetType().IsAssignableFrom(typeof(StrMaxLenAttribute)))
                    {
                        StrMaxLenAttribute attr = ((StrMaxLenAttribute)attribute);
                        ErrorInfo errorInfo = null;
                        if (!attr.IsValid(propertyValue, out errorInfo))
                        {
                            errorInfo.ColumnName = columnName;
                            errorInfo.Column = columnNum;
                            errorInfo.Row = rowNum;
                            errorInfos.Add(errorInfo);
                        }
                    }
                    //判断是否有长度验证属性
                    if (attribute.GetType().IsAssignableFrom(typeof(NoSpaceAttribute)))
                    {
                        NoSpaceAttribute attr = ((NoSpaceAttribute)attribute);
                        ErrorInfo errorInfo = null;
                        if (!attr.IsValid(propertyValue, out errorInfo))
                        {
                            errorInfo.ColumnName = columnName;
                            errorInfo.Column = columnNum;
                            errorInfo.Row = rowNum;
                            errorInfos.Add(errorInfo);
                        }
                    }
                }
            }

            //去数据库验证 一次性验证所有列
            foreach (Attribute attribute in attributes)
            {
                //判断是否有长度验证属性
                if (attribute.GetType().IsAssignableFrom(typeof(DatabaseHaveAttribute)))
                {
                    DatabaseHaveAttribute attr = ((DatabaseHaveAttribute)attribute);
                    List<ErrorInfo> errors = new List<ErrorInfo>();
                    if (!attr.IsValid(values, out errors))
                    {
                        errors.ForEach(info => info.Column = columnNum);
                        errorInfos.InsertRange(0, errors);
                    }
                }
            }


            return errorInfos;
        }

        /// <summary>
        /// 转换DataTable为实体
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="mapConfigs">映射</param>
        /// <returns></returns>
        private static List<TEntity> ConvertMapToEntity<TEntity>(DataTable dataTable, List<MapConfig> mapConfigs) where TEntity : new()
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
        /// 转换DataTable为实体
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="mapConfigs">映射</param>
        /// <returns></returns>
        private static List<TEntity> ConvertToEntity<TEntity>(DataTable dataTable) where TEntity : new()
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
                            throw new Exception($"字段[{ dtColumnName }]数据类型出错,{ex.Message}");

                        }
                    }
                }
                entities.Add(entity);
            }
            return entities;
        }

        /// <summary>
        /// 映射实体字段到DataTable
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="dataTable"></param>
        /// <param name="mapConfigs">映射</param>
        /// <returns></returns>
        private static DataTable MapToDataTable(DataTable dataTable, List<MapConfig> mapConfigs)
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


    }

    /// <summary>
    /// 配置委托
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="arg"></param>
    public delegate void DesignDelegate(AddMapConfig addMapConfig);

    public class AddMapConfig
    {
        public List<MapConfig> mapConfigs = null;
        /// <summary>
        /// 配置委托
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arg"></param>
        public void Add<TConfig>() where TConfig : new()
        {
            Type config = typeof(TConfig);
            object Instance = Activator.CreateInstance(config);
            mapConfigs = new List<MapConfig>();
            foreach (FieldInfo field in config.GetFields())
            {
                mapConfigs.Add(new MapConfig
                {
                    EntityColumnName = field.Name,
                    DataTableColumnName = field.GetValue(Instance).ToString()
                });
            }
        }
    }


}
