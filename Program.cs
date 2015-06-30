using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace ScriptXMLConvert
{
    public enum Node
    {
        Scene = 0,
        Act,
        Moment,
        Line
    }

    class Program
    {
        private static XmlDocument sceneBreakdown = new XmlDocument();
        private static XmlDocument originalFile = new XmlDocument();
        private static string path = @"c:/users/jeffrey.johnson/Documents/Github/Project_Kansas/Assets/";
        private static string defaultNameSpace = "urn:schemas-microsoft-com:office:spreadsheet";

        static void Main(string[] args)
        {
            LoadOriginalFile();
            BuildSceneBreakdown();

        }

        static void BuildSceneBreakdown()
        {
            //create root
            sceneBreakdown.AppendChild(sceneBreakdown.CreateElement("script"));
            //get rows
            XmlNodeList rowsList = GetRowsFromOriginal();
            foreach (XmlNode row in rowsList)
            {
                //skip first row
                if (row.FirstChild.FirstChild.Attributes.GetNamedItem("Type", "urn:schemas-microsoft-com:office:spreadsheet").Value == "String")
                {
                    //Console.WriteLine("here");
                    continue;
                }


                //get all the cells in this row
                XmlNodeList cellsList = GetCellList(row);

                XmlElement actNode = sceneBreakdown.CreateElement("act");
                XmlElement sceneNode = sceneBreakdown.CreateElement("scene");
                string currentScene = "0";
                

                //check if this is a row declaring the Act
                if (IsActRow(cellsList))
                {
                    string innerText = (cellsList.Item(0).FirstChild.InnerText);
                    string actNum = innerText[innerText.Length - 1].ToString();
                    Console.WriteLine("Act: " + actNum);
                    //write the xml
                    
                    XmlAttribute data = sceneBreakdown.CreateAttribute("number");
                    data.Value = actNum;
                    actNode.Attributes.Append(data);
                    //don't need to go further
                    continue;
                }

                //unless blank have data row
                if ( !IsBlankRow(cellsList))
                {
                    //check if need new scene node
                    string scene = cellsList.Item(0).FirstChild.InnerText;
                    if (scene == currentScene)
                    {
                        //add another moment child

                    }

                }


            }
        }

        private static bool IsBlankRow(XmlNodeList cellsList)
        {
            return cellsList.Item(0).ChildNodes.Count == 0 && cellsList.Item(1).ChildNodes.Count == 0;
        }

        private static bool IsActRow(XmlNodeList cellsList)
        {
            //check if cell 1 has data and cell 2 has none
            if (cellsList.Item(0).ChildNodes.Count == 1 && cellsList.Item(1).ChildNodes.Count == 0)
            {
                return true;
            }
            return false;
        }

        private static XmlNodeList GetCellList(XmlNode rowNode)
        {
            XmlNamespaceManager nsMan = new XmlNamespaceManager(originalFile.NameTable);
            nsMan.AddNamespace("def", defaultNameSpace);

            return rowNode.SelectNodes("./def:Cell", nsMan);
        }

        static void LoadOriginalFile()
        {
            originalFile.Load(path + "SCENE BREAKDOWN - KANSAS.xml");
        }

        static XmlNodeList GetRowsFromOriginal()
        {
            XmlNamespaceManager nsMan = new XmlNamespaceManager(originalFile.NameTable);
            nsMan.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
            nsMan.AddNamespace("def", defaultNameSpace);


            return originalFile.SelectNodes("/def:Workbook/ss:Worksheet/def:Table/def:Row", nsMan);
        }


    }
}
