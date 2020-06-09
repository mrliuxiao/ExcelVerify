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
                    Name2 = $"Lucio{_i}",
                    Sex = true
                });
            }
            testEntity.Add(new TestEntity
            {
                Id = 7,
                Name = $"Lucio ",
                Name2 = $"Lucio ",
                Sex = true
            });
            testEntity.Add(new TestEntity
            {
                Id = 7,
                Name = $"Lucio asdasdasidhaiushdiaushdiuashdiuahsdiuahsdiuhasiudhaiusdhiaushdiuashdiuahsfoiusdahfgoiusdhfgoiufsdhgoiuhfdsoiuhgfidsuhfogiuhi",
                Name2 = $"Lucio asdasdasidhaiushdiaushdiuashdiuahsdiuahsdiuhasiudhaiusdhiaushdiuashdiuahsfoiusdahfgoiusdhfgoiufsdhgoiuhfdsoiuhgfidsuhfogiuhi",
                Sex = true
            });
            List<MapConfig> mapConfigs = new List<MapConfig>();
            mapConfigs.Add(new MapConfig { DataTableColumnName = "姓名2", EntityColumnName = "Name2" });
            mapConfigs.Add(new MapConfig { DataTableColumnName = "姓名", EntityColumnName = "Name" });
            mapConfigs.Add(new MapConfig { DataTableColumnName = "性别", EntityColumnName = "Sex" });

            //MaoConvert
            DataTable testTable1 = ExcelVerifiyHelper.ConvertToDataTable(testEntity);
            DataTable testTable2 = ExcelVerifiyHelper.ConvertMapToDataTable(testEntity, cfg => { cfg.Add<TestEntityMapConfig>(); });
            List<TestEntity> testEntities1 = ExcelVerifiyHelper.ConvertToEntity<TestEntity>(testTable1);
            List<TestEntity> testEntities2 = ExcelVerifiyHelper.ConvertToEntity<TestEntity>(testTable2);
            List<TestEntity> testEntities3 = ExcelVerifiyHelper.ConvertMapToEntity<TestEntity>(testTable1, cfg => { cfg.Add<TestEntityMapConfig>(); });
            List<TestEntity> testEntities4 = ExcelVerifiyHelper.ConvertMapToEntity<TestEntity>(testTable2, cfg => { cfg.Add<TestEntityMapConfig>(); });

            List<ErrorRowInfo> errorRowInfos = new List<ErrorRowInfo>();
            //第一种 效验Datatable
            List<TestEntity> entities1 = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(testTable1, mapConfigs, out errorRowInfos);
            List<TestEntity> entities2 = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(testTable1, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);
            DataTable dataTable3 = ExcelVerifiyHelper.VerifyToDataTable<TestEntity>(testTable1, mapConfigs, out errorRowInfos);
            DataTable dataTable4 = ExcelVerifiyHelper.VerifyToDataTable<TestEntity>(testTable1, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);

            //第三种 效验实体
            List<TestEntity> entities3 = ExcelVerifiyHelper.VerifyToEntitys(testEntities1, mapConfigs, out errorRowInfos);
            List<TestEntity> entities4 = ExcelVerifiyHelper.VerifyToEntitys(testEntities1, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);
            DataTable dataTable5 = ExcelVerifiyHelper.VerifyToDataTable(testEntities1, mapConfigs, out errorRowInfos);
            DataTable dataTable6 = ExcelVerifiyHelper.VerifyToDataTable(testEntities3, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);

            //第四种 效验Excel
            string path = "D:\\Excel测试";
            string fileName = "2020年6月9日.xlsx";

            string excelFileName = $"{path}\\{DateTime.Now.ToString("D")}.xlsx";
            string excelFileName2 = $"{path}\\{DateTime.Now.ToString("D")}22.xlsx";
            ////导出Excel
            //ExcelVerifiyHelper.DataTableToExcel(dataTable5, excelFileName);
            //ExcelVerifiyHelper.EntityToExcel(entities3, config => { config.Add<TestEntityMapConfig>(); }, excelFileName2);

            List<TestEntity> entities5 = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(excelFileName, mapConfigs, out errorRowInfos);
            List<TestEntity> entities6 = ExcelVerifiyHelper.VerifyToEntitys<TestEntity>(excelFileName, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);
            DataTable dataTable7 = ExcelVerifiyHelper.VerifyToDataTable<TestEntity>(excelFileName2, mapConfigs, out errorRowInfos);
            DataTable dataTable8 = ExcelVerifiyHelper.VerifyToDataTable<TestEntity>(excelFileName2, config => { config.Add<TestEntityMapConfig>(); }, out errorRowInfos);

          

            Console.WriteLine("------------错误信息---------------");
            Console.WriteLine(JsonConvert.SerializeObject(errorRowInfos));
            Console.WriteLine("------------验证成功实体---------------");
            Console.WriteLine(JsonConvert.SerializeObject(entities1));

            Console.ReadLine();
        }
    }


    public class TestEntity
    {
        //[NotNull]
        public int Id { get; set; }

        [ExcelStrLength(0, 50)]
        [ExcelNoSpace]
        [ExcelDatabaseHaveGetIds("Name2", typeof(VerifyConfig))]
        public string Name2 { get; set; }
        [ExcelDatabaseHave("Name", typeof(VerifyConfig))]
        public string Name { get; set; }
        public bool Sex { get; set; }


    }

    public class VerifyConfig
    {
        public string Name2 = "server=.;database=Test;uid=Lucio;pwd=Lucio;Table=TestEntity;PrimaryKey=Id;Column=Name";
        public string Name = "server=.;database=Test;uid=Lucio;pwd=Lucio;Table=TestEntity;PrimaryKey=Id;Column=Name";
        public string Sex = "server=.;database=Test;uid=Lucio;pwd=Lucio;Table=TestEntity;PrimaryKey=Id;Column=Sex";
    }

    public class TestEntityMapConfig
    {
        //实体属性 = 表格属性
        public string Name2 = "姓名2";
        public string Name = "姓名";
        public string Sex = "性别";
    }
}
