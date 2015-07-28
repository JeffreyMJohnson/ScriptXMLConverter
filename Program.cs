using System;
using System.IO;
using System.Xml;
using Excel;
using System.Collections.Generic;


namespace ScriptXMLConvert
{

    /// <summary>
    /// This tool converts the xml data from Excel sheet SCENE BREAKDOWN - KANSAS.xml to an xml file with the following schema:
    ///     <script>
    ///        <act number="1">
    ///           <scene number="1">
    ///               <text>[TEXT OF SCRIPT FOR THIS SCENE]</text>
    ///               <moments>
    ///                  <moment title="foo" line="bar"/>
    ///               </moments>
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
            Line,
            Duration
        }

        private static XmlDocument sceneBreakdown = new XmlDocument();
        private static XmlDocument originalFile = new XmlDocument();
        private static string defaultNameSpace = "urn:schemas-microsoft-com:office:spreadsheet";
        private static StreamReader scriptFileStream;




        static void Main(string[] args)
        {
            originalFile.Load("SCENE BREAKDOWN - KANSAS.xml");
            //BuildSceneBreakdown();
            BuildSceneBreakdown();
            //scriptFileStream = new StreamReader("SCRIPT-CODENAMEKANSAS.txt");
            //AddScriptText();


            sceneBreakdown.Save("SceneBreakdown.xml");

        }

        /// <summary>
        /// Break up the script text into scenes and add it to xmlfile. Note there is a CHARACTERS section in beginning that is not being 
        /// included. This could be put into seperate section if customer wants.
        /// </summary>
        //private static void AddScriptText()
        //{
        //    //load the file into a string
        //    //string scriptText = scriptFileStream.ReadToEnd();
        //    string scriptText = "";

        //    while (scriptFileStream.Peek() >= 0)
        //    {
        //        string line = scriptFileStream.ReadLine();
        //        string cleanedLine = "";

        //        for (int i = 0; i < line.Length; i++)
        //        {
        //            int ascii = (int)line[i];
        //            char c = line[i];
        //            if (ascii > 31 && ascii < 127)
        //            {
        //                cleanedLine += c;
        //            }
        //        }
        //        scriptText += cleanedLine + "\n";

        //    }


        //    //loop through the scenes in the xml file and add text script using "SCENE " as delimiter.
        //    //scenes are in order so can assume act
        //    XmlNodeList scenes = sceneBreakdown.SelectNodes("script/act/scene");

        //    int cursorIndex = 0;
        //    for (int i = 0; i < scenes.Count; i++)
        //    {
        //        //set cursor to beginning of the "Act ..." for this scene
        //        int sceneNum = int.Parse(scenes[i].Attributes.GetNamedItem("number").Value);
        //        int sceneIndex = scriptText.IndexOf("SCENE " + sceneNum, cursorIndex);
        //        string sceneText;
        //        int nextSceneIndex;
        //        //need next scene location (if one) for end delim
        //        if (i + 1 < scenes.Count)
        //        {
        //            int nextSceneNum = int.Parse(scenes[i + 1].Attributes.GetNamedItem("number").Value);
        //            nextSceneIndex = scriptText.IndexOf("SCENE " + nextSceneNum, sceneIndex);
        //            sceneText = scriptText.Substring(sceneIndex, nextSceneIndex - sceneIndex);
        //            cursorIndex = nextSceneIndex;
        //        }
        //        else
        //        {
        //            sceneText = scriptText.Substring(sceneIndex);
        //            cursorIndex = scriptText.Length;
        //        }
        //        //create the node
        //        XmlElement text = sceneBreakdown.CreateElement("text");
        //        text.InnerText = sceneText;
        //        //add data as cdata because text has illegal xml chars
        //        //XmlCDataSection data = sceneBreakdown.CreateCDataSection(sceneText);
        //        //text.AppendChild(data);

        //        scenes[i].AppendChild(text);



        //    }

        //}

        static int NumberOfFilledCells(Cell[] cellsList)
        {
            int result = 0;
            foreach (Cell cell in cellsList)
            {
                if (!IsCellEmpty(cell))
                {
                    result++;
                }
            }
            return result;
        }

        static bool IsCellEmpty(Cell cell)
        {
            return null == cell || cell.Text == "";
        }

        static void BuildSceneBreakdown()
        {
            //create root
            XmlElement scriptNode = sceneBreakdown.CreateElement("script");
            sceneBreakdown.AppendChild(scriptNode);

            //get rows from excel
            IEnumerator<worksheet> sheets = Workbook.Worksheets("SCENE BREAKDOWN - KANSAS.xlsx").GetEnumerator();
            sheets.MoveNext();
            Row[] rowsList = sheets.Current.Rows;
            //get rows
            //XmlNodeList rowsList = GetRowsFromOriginal();

            XmlElement actNode = null;
            XmlElement sceneNode = null;
            string currentScene = "0";

            foreach (Row row in rowsList)
            {
                //get all the cells in this row
                //XmlNodeList cellsList = GetCellList(row);
                Cell[] cellsList = row.Cells;

                //note:the end of the file has a bunch of rows for excel formatting, want to skip them
                //HACK:using this to signify the end of the data in the file so appending last actNode to it's parent
                if (cellsList.Length < 2)
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
                //if (row.FirstChild.FirstChild.Attributes.GetNamedItem("Type", "urn:schemas-microsoft-com:office:spreadsheet").Value == "String" && row.ChildNodes.Item(1).ChildNodes.Count > 0)
                if (cellsList[(int)ColumnHeader.Scene].Text == "SCENE" && cellsList[(int)ColumnHeader.Moment].Text == "MOMENT" && cellsList[(int)ColumnHeader.Line].Text == "LINE")
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

                    //string innerText = (cellsList.Item(0).FirstChild.InnerText);
                    //string actNum = innerText[innerText.Length - 1].ToString();
                    string innerText = cellsList[(int)ColumnHeader.Scene].Text;
                    string actNum = innerText.Substring(innerText.LastIndexOf(' ') + 1);

                    //write the xml
                    XmlAttribute data = sceneBreakdown.CreateAttribute("number");
                    data.Value = actNum;
                    actNode.Attributes.Append(data);
                    //move to next row
                    continue;
                }



                //get scene number
                string scene = cellsList[(int)ColumnHeader.Scene].Text;
                

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
                //titleNode.Value = GetCellData(ColumnHeader.Moment, cellsList);
                titleNode.Value = cellsList[(int)ColumnHeader.Moment].Text;
                momentNode.Attributes.Append(titleNode);

                //add line attribute
                XmlAttribute lineNode = sceneBreakdown.CreateAttribute("line");
                //lineNode.Value = GetCellData(ColumnHeader.Line, cellsList);
                if(null != cellsList[(int)ColumnHeader.Line])
                {
lineNode.Value = cellsList[(int)ColumnHeader.Line].Text;
                }
                else
                {
                    lineNode.Value = "";
                }
                
                
                momentNode.Attributes.Append(lineNode);

                //write to scene parent
                sceneNode.AppendChild(momentNode);


            }
            //append scene to act


        }


        //private static string GetCellData(ColumnHeader column, XmlNodeList cellsList)
        //{
        //    try
        //    {
        //        return cellsList.Item((int)column).FirstChild.InnerText;
        //    }
        //    catch (NullReferenceException e)
        //    {
        //        //make sure this is only for the line column
        //        if (column == ColumnHeader.Line)
        //        {
        //            return "";
        //        }
        //        else
        //        {
        //            throw e;
        //        }

        //    }


        //}

        private static bool IsBlankRow(Cell[] cellsList)
        {
            //return cellsList.Item(0).ChildNodes.Count == 0 && cellsList.Item(1).ChildNodes.Count == 0;
            return NumberOfFilledCells(cellsList) == 0;
        }

        private static bool IsActRow(Cell[] cellsList)
        {
            //check if cell 1 has data and cell 2 has none
            //if (cellsList.Item(0).ChildNodes.Count == 1 && cellsList.Item(1).ChildNodes.Count == 0)
            if(NumberOfFilledCells(cellsList) == 1)
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
