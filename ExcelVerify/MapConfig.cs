using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelVerify
{
    /// <summary>
    /// 映射配置
    /// </summary>
    public class MapConfig
    {
        /// <summary>
        /// 映射datatable列
        /// </summary>
        public string DataTableColumnName { get; set; }
        /// <summary>
        /// 映射实体属性
        /// </summary>
        public string EntityColumnName { get; set; }
    }

    /// <summary>
    /// 添加注册映射配置
    /// </summary>
    /// <param name="addMapConfig">添加配置类</param>
    public delegate void DesignDelegate(AddMapConfig addMapConfig);

    /// <summary>
    /// 添加注册映射配置
    /// </summary>
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