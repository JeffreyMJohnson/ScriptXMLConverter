using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScriptXMLConvert;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ScriptXMLConverter_unit_test
{
    [TestClass()]
    public class ScriptTests
    {
        const string TEST_DATA_PATH_1 = @"..\..\Data\TestData_1.xlsx";

        [TestInitialize]
        public void TestInit()
        {

        }

        [TestMethod()]
        public void ScriptTest()
        {
            Script script = new Script();

            Assert.AreEqual(0, script.Acts.Length);
            Assert.IsNull(script.TotalTime);
        }

        [TestMethod()]
        public void ScriptTest2()
        {
            XLSXDataSource data = new XLSXDataSource(TEST_DATA_PATH_1);
            SheetRow[] rows = data.GetRows();
            Script script = new Script(rows);

            Assert.AreEqual("01:30:00", script.TotalTime);
            Assert.AreEqual(2, script.Acts.Length);
            Assert.AreEqual(3, script.Acts[0].Scenes.Length);
            Assert.AreEqual(3, script.Acts[1].Scenes.Length);

            Assert.AreEqual(12, script.Acts[0].Scenes[0].Moments.Length);
            Assert.AreEqual("FIRST SET : INTRO", script.Acts[0].Scenes[0].Moments[0].Title);
            Assert.AreEqual("", script.Acts[0].Scenes[0].Moments[0].Duration);

            Moment momentUnderTest = script.Acts[1].Scenes[1].Moments[6];
            Assert.AreEqual("Out of sex scene", momentUnderTest.Title);
            Assert.AreEqual("Guess you're not in Kansas any more", momentUnderTest.Line);
            Assert.AreEqual("", momentUnderTest.Duration);
            Assert.AreEqual("", momentUnderTest.Location);
            Assert.AreEqual("", momentUnderTest.SFX);

        }

        [TestMethod()]
        public void AddAttribute()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("root");
            doc.AppendChild(root);
            Script script = new Script();
            PrivateObject priv_script = new PrivateObject(script);

            priv_script.Invoke("AddAttribute", doc, root, "rootAttribute", "attribute value");

            Assert.AreEqual(1, doc.ChildNodes.Count);
            Assert.AreEqual(1, root.Attributes.Count);
            Assert.AreEqual("rootAttribute", root.Attributes.GetNamedItem("rootAttribute").Name);
            Assert.AreEqual("attribute value", doc.FirstChild.Attributes.GetNamedItem("rootAttribute").Value);
        }

        [TestMethod()]
        public void GetXML()
        {
            XLSXDataSource data = new XLSXDataSource(TEST_DATA_PATH_1);
            SheetRow[] rows = data.GetRows();
            Script script = new Script(rows);
            XmlDocument doc = script.GetXML();

            Assert.AreEqual(2, doc.SelectNodes("script/act").Count);
            Assert.AreEqual(3, doc.SelectNodes("script/act[@number='1']/scene").Count);
            Assert.AreEqual(4, doc.SelectNodes("script/act[@number='2']/scene[@number='3']/moment").Count);
            Assert.AreEqual("We will need to shut down her ", doc.SelectSingleNode("script /act[@number='2']/scene[@number='3']/moment[@title='Planning Dialogue, Leo joins']/@line").Value);

        }
    }
}