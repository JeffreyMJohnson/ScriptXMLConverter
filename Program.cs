using System;
using System.Xml;

namespace ScriptXMLConvert
{

    /// <summary>
    /// This tool converts the xml data from Excel sheet SCENE BREAKDOWN - KANSAS.xml to an xml file with the following schema:
    ///     <script>
    ///        <act number="1">
    ///           <scene number="1">
    ///               <moment title="foo" line="bar"/>
    ///            </scene>
    ///       </act>
    ///     </script>
    ///     Single script root node with one or more act nodes children with one or more scene nodes children with one or more moment children.
    ///     
    /// The excel xm file must be in the same directory as this exe is running and will output the data file SceneBreakdown.xml in the same directory.
    /// 
    /// Written by Jeffrey M. Johnson
    /// jefjohms@gmail.com
    /// 
    /// </summary>
    class Program
    {
        public enum ColumnHeader
        {
            Scene = 0,
            Moment,
            Line
        }

        private static XmlDocument sceneBreakdown = new XmlDocument();
        private static XmlDocument originalFile = new XmlDocument();
        private static string defaultNameSpace = "urn:schemas-microsoft-com:office:spreadsheet";

        static void Main(string[] args)
        {
            originalFile.Load("SCENE BREAKDOWN - KANSAS.xml");
            BuildSceneBreakdown();
            sceneBreakdown.Save("SceneBreakdown.xml");
        }

        static void BuildSceneBreakdown()
        {
            //create root
            XmlElement scriptNode = sceneBreakdown.CreateElement("script");
            sceneBreakdown.AppendChild(scriptNode);


            //get rows
            XmlNodeList rowsList = GetRowsFromOriginal();

            XmlElement actNode = null;
            XmlElement sceneNode = null;
            string currentScene = "0";

            foreach (XmlNode row in rowsList)
            {
                //get all the cells in this row
                XmlNodeList cellsList = GetCellList(row);

                //note:the end of the file has a bunch of rows for excel formatting, want to skip them
                //HACK:using this to signify the end of the data in the file so appending last actNode to it's parent
                if (cellsList.Count < 2)
                {
                    actNode.AppendChild(sceneNode);
                    scriptNode.AppendChild(actNode);
                    break;
                }

                //skip blank row
                if (IsBlankRow(cellsList))
                {
                    //move to next row
                    continue;
                }

                //skip first row
                if (row.FirstChild.FirstChild.Attributes.GetNamedItem("Type", "urn:schemas-microsoft-com:office:spreadsheet").Value == "String" && row.ChildNodes.Item(1).ChildNodes.Count > 0)
                {
                    continue;
                }


                //check if this is a row declaring the Act
                if (IsActRow(cellsList))
                {

                    //if not the first act child (actNode == null) then add to script parent
                    if (null != actNode)
                    {
                        actNode.AppendChild(sceneNode);
                        scriptNode.AppendChild(actNode);
                    }

                    //create new act node
                    actNode = sceneBreakdown.CreateElement("act");

                    //reset scene and currentScene
                    sceneNode = null;
                    currentScene = "0";

                    string innerText = (cellsList.Item(0).FirstChild.InnerText);
                    string actNum = innerText[innerText.Length - 1].ToString();


                    //write the xml
                    XmlAttribute data = sceneBreakdown.CreateAttribute("number");
                    data.Value = actNum;
                    actNode.Attributes.Append(data);
                    //move to next row
                    continue;
                }



                //get scene number
                string scene = GetCellData(ColumnHeader.Scene, cellsList);

                //create new scene node and set the data if new
                if (scene != currentScene)
                {
                    //if not the first scene child (sceneNode == null) then write scene element to parent
                    if (null != sceneNode)
                    {
                        actNode.AppendChild(sceneNode);
                    }

                    currentScene = scene;
                    sceneNode = sceneBreakdown.CreateElement("scene");

                    XmlAttribute sceneNum = sceneBreakdown.CreateAttribute("number");
                    sceneNum.Value = scene;
                    sceneNode.Attributes.Append(sceneNum);
                }

                //create another moment node
                XmlElement momentNode = sceneBreakdown.CreateElement("moment");

                //add title attribute
                XmlAttribute titleNode = sceneBreakdown.CreateAttribute("title");
                titleNode.Value = GetCellData(ColumnHeader.Moment, cellsList);
                momentNode.Attributes.Append(titleNode);

                //add line attribute
                XmlAttribute lineNode = sceneBreakdown.CreateAttribute("line");
                lineNode.Value = GetCellData(ColumnHeader.Line, cellsList);
                momentNode.Attributes.Append(lineNode);

                //write to scene parent
                sceneNode.AppendChild(momentNode);


            }
            //append scene to act


        }


        private static string GetCellData(ColumnHeader column, XmlNodeList cellsList)
        {
            try
            {
                return cellsList.Item((int)column).FirstChild.InnerText;
            }
            catch (NullReferenceException e)
            {
                //make sure this is only for the line column
                if (column == ColumnHeader.Line)
                {
                    return "";
                }
                else
                {
                    throw e;
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

        static XmlNodeList GetRowsFromOriginal()
        {
            XmlNamespaceManager nsMan = new XmlNamespaceManager(originalFile.NameTable);
            nsMan.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
            nsMan.AddNamespace("def", defaultNameSpace);


            return originalFile.SelectNodes("/def:Workbook/ss:Worksheet/def:Table/def:Row", nsMan);
        }


    }
}
