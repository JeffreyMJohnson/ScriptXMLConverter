using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScriptXMLConvert;
using Excel;


namespace ScriptXMLConverter_unit_test
{
    [TestClass]
    public class XLSXDataSource_Test
    {

        XLSXDataSource _dataSource;
        PrivateObject _privateDataSource;

        const string TEST_DATA_PATH_1 = @"..\..\Data\TestData_1.xlsx";


        [TestInitialize]
        public void TestInit()
        {
            _dataSource = new XLSXDataSource(TEST_DATA_PATH_1);
            _privateDataSource = new PrivateObject(_dataSource);
        }

        
        [TestMethod]
        public void GetRows()
        {
            Assert.AreEqual(61, _dataSource.GetRows().Length);
            SheetRow row = _dataSource.GetRows()[8];
            Assert.AreEqual("1", row.Scene);
            Assert.AreEqual("Witches gain power", row.Moment);
            Assert.AreEqual("These technomancers used their control of system technologies", row.Line);
            Assert.AreEqual("1.00", row.Duration);
            Assert.AreEqual("", row.Location);
            Assert.AreEqual("Witches gain power", row.Moment);
            Assert.AreEqual("GREEN_LM", row.SFX);
        }
    }
}
