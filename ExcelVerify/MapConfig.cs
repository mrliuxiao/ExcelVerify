using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelVerify
{
    public class MapConfig
    {
        public string DataTableColumnName { get; set; }
        public string EntityColumnName { get; set; }
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