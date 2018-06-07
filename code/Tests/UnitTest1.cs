using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web;
using System.Web.SessionState;
using CotideFW.Framework.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using CotideFW.Framework.Extensions;

namespace Tests
{
    [TestClass]
    public class Helper
    { 
        [TestInitialize]
        public void TestSetup()
        {  
        }


        [TestMethod]
        public void NPOIHelperTest()
        { 
            var list = new List<TestData>()
            {

                new TestData("用户1",25,DateTime.Now),
                new TestData("用户2",26,DateTime.Now),
                new TestData("用户3",27,DateTime.Now),
            }; 

            var outPut = list.Select(x => new
            {
                名称 = x.Name,
                年龄 = x.Age,
                创建时间 = x.CreateTime

            }).ToList().ToDataTable(); 
            NPOIHelper.ExportByFile(outPut, "用户数据", "C:\\TT\\Excel\\1.xls");
        }


        public class TestData
        {
            public TestData()
            {

            }

            public TestData(string name,int age,DateTime createTime)
            {
                Name = name;
                Age = age;
                createTime = createTime;
            }

            public string Name { get; set; }

            public int Age { get; set; }

            public DateTime CreateTime { get; set; }
        }
    }
}
