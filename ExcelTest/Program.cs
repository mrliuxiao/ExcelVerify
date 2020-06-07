using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelVerify;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Newtonsoft.Json;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            List<TestEntity> testEntity = new List<TestEntity>();
            for (int _i = 0; _i <= 10; _i++)
            {
                testEntity.Add(new TestEntity
                {
                    Id = _i,
                    Name = $"Lucio{_i}",
                    Sex = true
                });
            }
            testEntity.Add(new TestEntity
            {
                Id = 7,
                Name = $"Lucio ",
                Sex = true
            });
            testEntity.Add(new TestEntity
            {
                Id = 7,
                Name = $"Lucio asdasdasidhaiushdiaushdiuashdiuahsdiuahsdiuhasiudhaiusdhiaushdiuashdiuahsfoiusdahfgoiusdhfgoiufsdhgoiuhfdsoiuhgfidsuhfogiuhi",
                Sex = true
            });
            List<MapConfig> mapConfigs = new List<MapConfig>();
            mapConfigs.Add(new MapConfig { DataTableColumnName = "姓名", EntityColumnName = "Name" });
            mapConfigs.Add(new MapConfig { DataTableColumnName = "性别", EntityColumnName = "Sex" });

            DataTable testTable = DataTableHelper.ToDataTable(testEntity);
            testTable.Columns["Name"].ColumnName = "姓名";
            testTable.Columns["Sex"].ColumnName = "性别";
            List<ErrorRowInfo> errorRowInfos = new List<ErrorRowInfo>();
            //第一种
            //List<TestEntity> entities = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(testTable, mapConfigs, out errorRowInfos);
            //第二种
            List<TestEntity> entities = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(testTable, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);

            Console.WriteLine("------------错误信息---------------");
            Console.WriteLine(JsonConvert.SerializeObject(errorRowInfos));
            Console.WriteLine("------------验证成功实体---------------");
            Console.WriteLine(JsonConvert.SerializeObject(entities));

            Console.ReadLine();
        }
    }


    public class TestEntity
    {
        [NotNull]
        public int Id { get; set; }

        [StrMaxLen(0, 50)]
        [NoSpace]
        [DatabaseHave("Name", typeof(VerifyConfig))]
        public string Name { get; set; }
        public bool Sex { get; set; }


    }

    public class VerifyConfig
    {
        public string Name = "server=.;database=Test;uid=Lucio;pwd=Lucio;Table=TestEntity;PrimaryKey=Id;Column=Name";
        public string Sex = "server=.;database=Test;uid=Lucio;pwd=Lucio;Table=TestEntity;PrimaryKey=Id;Column=Sex";
    }

    public class TestEntityMapConfig
    {
        //实体属性 = 表格属性
        public string Name = "姓名";
        public string Sex = "性别";
    }
}
