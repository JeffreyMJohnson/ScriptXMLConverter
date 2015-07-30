using System;
using System.IO;
using System.Xml;
using Excel;
using System.Collections.Generic;


namespace ScriptXMLConvert
{

    /// <summary>
    /// This tool converts the xml data from Excel sheet SCENE BREAKDOWN - KANSAS.xml to an xml file with the following schema:
    ///     <script filmLength="01:30:00">
    ///        <act number="1">
    ///           <scene number="1" time="00:00:00">
    ///             <!--  <text>[TEXT OF SCRIPT FOR THIS SCENE]</text> THIS WAS REMOVED BUT CAN BE ADDED, SEE COMMENTS -->
    ///               <moment title="foo" line="bar" duration="5.5" location="1.0, 5.0, 0.0" sfx="Fire"/>
    ///            </scene>
    ///       </act>
    ///     </script>
    ///     
    /// The excel .xslx file must be in the same directory as this exe when running and will output the data file SceneBreakdown.xml in the same directory.
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
            Duration,
            Location,
            SFX
        }

        private static XmlDocument sceneBreakdown = new XmlDocument();

        /*
         * un comment next line to add script text to output data file.
         */
        //private static StreamReader scriptFileStream;




        static void Main(string[] args)
        {
            BuildSceneBreakdown();

            /*
             *This is commented to remove the script text from the outputed data file.  If want it in uncomment the next 2 lines.
             *The file with the script text must be in same directory as exe when running.
             */
            //scriptFileStream = new StreamReader("SCRIPT-CODENAMEKANSAS.txt");
            //AddScriptText();

            sceneBreakdown.Save("SceneBreakdown.xml");

        }



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

            XmlElement actNode = null;
            XmlElement sceneNode = null;
            string currentScene = "0";

            foreach (Row row in rowsList)
            {
                //get all the cells in this row
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
                    //move to next row
                    continue;
                }

                //skip time and total duration rows
                if (cellsList[(int)ColumnHeader.Scene].Text == "TIME" || cellsList[(int)ColumnHeader.Scene].Text == "SCRIPT TOTAL DURATION")
                {
                    //move to next row
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

                    string innerText = cellsList[(int)ColumnHeader.Scene].Text;
                    string actNum = innerText.Substring(innerText.LastIndexOf(' ') + 1);

                    //add number attribute
                    XmlAttribute numberAttribute = sceneBreakdown.CreateAttribute("number");
                    numberAttribute.Value = actNum;
                    actNode.Attributes.Append(numberAttribute);

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

                    XmlAttribute numberAttribute = sceneBreakdown.CreateAttribute("number");
                    numberAttribute.Value = scene;
                    sceneNode.Attributes.Append(numberAttribute);
                }

                //create another moment node
                XmlElement momentNode = sceneBreakdown.CreateElement("moment");

                //add title attribute
                SetMomentAttribute(momentNode, ColumnHeader.Moment, cellsList);

                //add line attribute
                SetMomentAttribute(momentNode, ColumnHeader.Line, cellsList);

                //add duration attribute
                SetMomentAttribute(momentNode, ColumnHeader.Duration, cellsList);

                //add location attribute
                SetMomentAttribute(momentNode, ColumnHeader.Location, cellsList);

                //add sfx attribute
                SetMomentAttribute(momentNode, ColumnHeader.SFX, cellsList);


                //write moment to scene parent
                sceneNode.AppendChild(momentNode);


            }
            //looped through all rows, still have to add final scene to act and that to the script.
            //append scene to act
            actNode.AppendChild(sceneNode);
            scriptNode.AppendChild(actNode);

            //added some attributes after the fact, and it was easier to just loop through rows again and set the new attributes.
            SetSceneTimes(rowsList);
        }

        static void SetMomentAttribute(XmlElement momentNode, ColumnHeader header, Cell[] cells)
        {
            //add attribute
            XmlAttribute attributeNode = null;
            switch (header)
            {
                case ColumnHeader.Duration:
                    attributeNode = sceneBreakdown.CreateAttribute("duration");
                    break;
                case ColumnHeader.Line:
                    attributeNode = sceneBreakdown.CreateAttribute("line");
                    break;
                case ColumnHeader.Moment:
                    attributeNode = sceneBreakdown.CreateAttribute("title");
                    break;
                case ColumnHeader.Location:
                    attributeNode = sceneBreakdown.CreateAttribute("location");
                    break;
                case ColumnHeader.SFX:
                    attributeNode = sceneBreakdown.CreateAttribute("sfx");
                    break;
                default:
                    Console.WriteLine("Wrong header type given in SetMomentAttribute.");
                    break;
            }

            if (null != cells[(int)header])
            {
                attributeNode.Value = cells[(int)header].Text;
            }
            else
            {
                attributeNode.Value = "";
            }
            momentNode.Attributes.Append(attributeNode);
        }

        static void SetSceneTimes(Row[] rows)
        {
            //need the current act and scene of each moment for xpath lookup.
            int currentAct = 0;
            for (int i = 0; i < rows.Length; i++)
            {
                Cell[] cells = rows[i].Cells;
                if (IsActRow(cells))
                {
                    string innerText = cells[(int)ColumnHeader.Scene].Text;
                    string actNum = innerText.Substring(innerText.LastIndexOf(' ') + 1);
                    currentAct = int.Parse(actNum);
                }

                if (cells[(int)ColumnHeader.Scene].Text == "TIME")
                {

                    if (i + 1 < rows.Length)
                    {
                        //TIME row precedes scenes it pertains to , so get correct scene by looking a row ahead.
                        string scene = rows[i + 1].Cells[(int)ColumnHeader.Scene].Text;
                        XmlNode sceneNode = sceneBreakdown.SelectSingleNode("script/act[@number='" + currentAct + "']/scene[@number='" + scene + "']");
                        XmlAttribute timeNode = sceneBreakdown.CreateAttribute("time");
                        timeNode.Value = cells[(int)ColumnHeader.Duration].Text;
                        sceneNode.Attributes.Append(timeNode);
                    }
                }

                if (cells[(int)ColumnHeader.Scene].Text == "SCRIPT TOTAL DURATION")
                {
                    XmlAttribute timeNode = sceneBreakdown.CreateAttribute("filmLength");
                    timeNode.Value = cells[(int)ColumnHeader.Duration].Text;
                    sceneBreakdown.FirstChild.Attributes.Append(timeNode);
                }
            }
        }

        private static bool IsBlankRow(Cell[] cellsList)
        {
            //return cellsList.Item(0).ChildNodes.Count == 0 && cellsList.Item(1).ChildNodes.Count == 0;
            return NumberOfFilledCells(cellsList) == 0;
        }

        private static bool IsActRow(Cell[] cellsList)
        {
            //check if cell 1 has data and cell 2 has none
            //if (cellsList.Item(0).ChildNodes.Count == 1 && cellsList.Item(1).ChildNodes.Count == 0)
            if (NumberOfFilledCells(cellsList) == 1)
            {
                return true;
            }
            return false;
        }


        /*
* un comment this method to add script text to output data file.
*/
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

    }

}
