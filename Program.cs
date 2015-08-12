using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
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

        //this switch is for which to use for data, local excel sheet or google sheet in cloud.
        private static bool usingExcel = true;

        private static XmlDocument sceneBreakdown = new XmlDocument();

        /*
         * un comment next line to add script text to output data file.
         */
        //private static StreamReader scriptFileStream;



        [STAThread]
        static void Main(string[] args)
        {
            BuildSceneBreakdown();
            try
            {
                //BuildSceneBreakdown();
            }
            catch (Exception e)
            {
                Console.WriteLine("Catastrophic failure occured:");
                Console.WriteLine("If this continues contact jefjohms@gmail.com");
                Console.WriteLine("press any key to terminate...");
                Console.ReadKey();
                return;
            }


            /*
             *This is commented to remove the script text from the outputed data file.  If want it in uncomment the next 2 lines.
             *The file with the script text must be in same directory as exe when running.
             */
            //scriptFileStream = new StreamReader("SCRIPT-CODENAMEKANSAS.txt");
            //AddScriptText();

            Console.WriteLine("Saving xml data file...");
            sceneBreakdown.Save("SceneBreakdown.xml");

            Console.WriteLine("Application finished.");
            Console.WriteLine("press any key to terminate...");
            Console.ReadKey();

        }




        static ListFeed GetRows()
        {
            //OAuth config
            string CLIENT_ID = "898242977449-gdhq44lj4h22jgv2gougnaktg6i482p9.apps.googleusercontent.com";
            string CLIENT_SECRET = "kiR_ogLwko8r_HviWpSWyj2p";
            string SCOPE = "https://spreadsheets.google.com/feeds";
            string REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";

            //setup OAuth object
            OAuth2Parameters parameters = new OAuth2Parameters();
            parameters.ClientId = CLIENT_ID;
            parameters.ClientSecret = CLIENT_SECRET;
            parameters.RedirectUri = REDIRECT_URI;
            parameters.Scope = SCOPE;

            //get auth URL
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            Console.WriteLine(authorizationUrl);

            //occasionally the clipboard decides to not work throwing this exception.  I copy to the clipboard for convenience, so if
            //doesn't work, user can select and copy the URL manually.
            try
            {
                System.Windows.Forms.Clipboard.Clear();
                System.Windows.Forms.Clipboard.SetText(authorizationUrl);
                Console.WriteLine("URL Copied to clipboard.");
            }
            catch (System.Runtime.InteropServices.ExternalException e)
            {
                Console.WriteLine("There was a problem copying to the clipboard, so copy the URL manually.");
            }

            Console.WriteLine("Please visit the URL above to authorize your OAuth " + "request token.  Once that is complete, type in your access code to "
                + "continue...");

            parameters.AccessCode = Console.ReadLine();


            try
            {
                OAuthUtil.GetAccessToken(parameters);
            }
            catch (System.Net.WebException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Did you copy/paste the given code exactly?");
                throw e;
            }



            string accessToken = parameters.AccessToken;
            Console.WriteLine("OAuth Access Token: " + accessToken);

            GOAuth2RequestFactory requestFactory = new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
            SpreadsheetsService service = new SpreadsheetsService("MySpreadsheetIntegration-v1");
            service.RequestFactory = requestFactory;

            // Instantiate a SpreadsheetQuery object to retrieve spreadsheets.
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = service.Query(query);

            if (feed.Entries.Count == 0)
            {
                Console.WriteLine("No Spreadsheet found");
                throw new Exception();
            }

            SpreadsheetEntry spreadsheet = null;
            foreach (SpreadsheetEntry entry in feed.Entries)
            {
                if (entry.Title.Text == "SCENE BREAKDOWN - KANSAS")
                {
                    spreadsheet = (SpreadsheetEntry)entry;
                }
            }

            if (null == spreadsheet)
            {
                Console.WriteLine("Could not find spreadsheet 'SCENE BREAKDOWN - KANSAS'");
                Console.WriteLine("Did you login with account that has access?");
                throw new Exception();
            }
            else
            {
                Console.WriteLine("Spreadsheet 'SCENE BREAKDOWN - KANSAS' found.");
            }


            WorksheetFeed wsFeed = spreadsheet.Worksheets;
            WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];

            // Define the URL to request the list feed of the worksheet.
            AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // Fetch the list feed of the worksheet.
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            ListFeed listFeed = service.Query(listQuery);
            if (listFeed.Entries.Count < 1)
            {
                Console.WriteLine("No rows returned with the data.\nDid you sign on with account that has access to spreadsheet?");
            }

            return listFeed;
        }


        static Script LoadScript()
        {
            Script script = new Script();
            if (!usingExcel)
            {
                Console.WriteLine("Using Google sheet on cloud for data...");
                ListFeed listFeed = GetRows();
                Act act = null;
                Scene scene = null;

                //loop rows
                foreach (ListEntry row in listFeed.Entries)
                {
                    string sceneValue = row.Elements[(int)ColumnHeader.Scene].Value;

                    //is act label
                    if (sceneValue.Contains("ACT "))
                    {
                        //if not first
                        if (null != act)
                        {
                            act.AddScene(scene);
                            scene = null;
                            script.AddAct(act);
                        }
                        act = new Act();
                        act.Number = sceneValue.Substring(sceneValue.LastIndexOf(' ') + 1);
                        //go to next row
                        continue;
                    }

                    //if new scene
                    if (sceneValue.Contains("TIME"))
                    {
                        //if not first scene 
                        if (null != scene)
                        {
                            act.AddScene(scene);
                        }
                        scene = new Scene();
                        scene.Time = row.Elements[(int)ColumnHeader.Duration].Value;
                        //go to next row
                        continue;
                    }

                    //if last element
                    if (sceneValue.Contains("SCRIPT TOTAL"))
                    {
                        script.TotalTime = row.Elements[(int)ColumnHeader.Duration].Value;
                        //add last scene to last act
                        act.AddScene(scene);
                        //add last act to script
                        script.AddAct(act);
                        //all done no need to continue checking rows
                        break;
                    }

                    //not above so it's a new moment
                    Moment moment = new Moment(row.Elements[(int)ColumnHeader.Moment].Value,
                                               row.Elements[(int)ColumnHeader.Line].Value,
                                               row.Elements[(int)ColumnHeader.Duration].Value,
                                               row.Elements[(int)ColumnHeader.Location].Value,
                                               row.Elements[(int)ColumnHeader.SFX].Value);
                    if (scene.Number != row.Elements[(int)ColumnHeader.Scene].Value)
                    {
                        scene.Number = row.Elements[(int)ColumnHeader.Scene].Value;
                    }
                    scene.AddMoment(moment);
                }
            }
            else
            {
                Console.WriteLine("Using .xlsx Excel sheet for data....");
                //get rows from excel
                IEnumerator<worksheet> sheets = Workbook.Worksheets("SCENE BREAKDOWN - KANSAS.xlsx").GetEnumerator();

                try
                {
                    sheets.MoveNext();
                }
                catch (System.IO.FileNotFoundException e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Is 'SCENE BREAKDOWN - KANSAS.xlsx' located in working directory?");
                    throw e;
                }

                Row[] rows = sheets.Current.Rows;

                Act act = null;
                Scene scene = null;

                //loop rows
                foreach (Row row in rows)
                {
                    //skip the first row;
                    if(GetCellText(row, ColumnHeader.Scene).Contains("SCENE") && GetCellText(row, ColumnHeader.Moment).Contains("MOMENT"))
                    {
                        continue;
                    }
                    //skip blank lines
                    if(GetCellText(row, ColumnHeader.Scene) == "" && GetCellText(row, ColumnHeader.Moment) == "")
                    {
                        continue;
                    }

                    string sceneValue = row.Cells[(int)ColumnHeader.Scene].Text;

                    //is act label
                    if (sceneValue.Contains("ACT "))
                    {
                        //if not first
                        if (null != act)
                        {
                            act.AddScene(scene);
                            scene = null;
                            script.AddAct(act);
                        }
                        act = new Act();
                        act.Number = sceneValue.Substring(sceneValue.LastIndexOf(' ') + 1);
                        //go to next row
                        continue;
                    }

                    //if new scene
                    if (sceneValue.Contains("TIME"))
                    {
                        //if not first scene 
                        if (null != scene)
                        {
                            act.AddScene(scene);
                        }
                        scene = new Scene();
                        scene.Time = row.Cells[(int)ColumnHeader.Duration].Text;
                        //go to next row
                        continue;
                    }

                    //if last element
                    if (sceneValue.Contains("SCRIPT TOTAL"))
                    {
                        script.TotalTime = row.Cells[(int)ColumnHeader.Duration].Text;
                        //add last scene to last act
                        act.AddScene(scene);
                        //add last act to script
                        script.AddAct(act);
                        //all done no need to continue checking rows
                        break;
                    }

                    //not above so it's a new moment
                    Moment moment = new Moment(row.Cells[(int)ColumnHeader.Moment].Text,
                                               row.Cells[(int)ColumnHeader.Line].Text,
                                               row.Cells[(int)ColumnHeader.Duration].Text,
                                               row.Cells[(int)ColumnHeader.Location].Text,
                                               row.Cells[(int)ColumnHeader.SFX].Text);
                    if (scene.Number != row.Cells[(int)ColumnHeader.Scene].Text)
                    {
                        scene.Number = row.Cells[(int)ColumnHeader.Scene].Text;
                    }
                    scene.AddMoment(moment);
                }
            }
            return script;

        }

        static string GetCellText(Row row, ColumnHeader column)
        {
            return row.Cells[(int)column].Text;
        }




        static void BuildSceneBreakdown()
        {
            //Create Script object
            Script script = LoadScript();

            //create root
            XmlElement scriptNode = sceneBreakdown.CreateElement("script");
            AddAttribute(scriptNode, "filmLength", script.TotalTime);
            sceneBreakdown.AppendChild(scriptNode);

            foreach (Act act in script.Acts)
            {
                XmlElement actNode = sceneBreakdown.CreateElement("act");
                AddAttribute(actNode, "number", act.Number);
                scriptNode.AppendChild(actNode);

                foreach (Scene scene in act.Scenes)
                {
                    XmlElement sceneNode = sceneBreakdown.CreateElement("scene");
                    AddAttribute(sceneNode, "number", scene.Number);
                    AddAttribute(sceneNode, "time", scene.Time);
                    actNode.AppendChild(sceneNode);

                    foreach (Moment moment in scene.Moments)
                    {
                        XmlElement momentNode = sceneBreakdown.CreateElement("moment");
                        AddAttribute(momentNode, "title", moment.Title);
                        AddAttribute(momentNode, "line", moment.Line);
                        AddAttribute(momentNode, "duration", moment.Duration);
                        AddAttribute(momentNode, "location", moment.Location);
                        AddAttribute(momentNode, "sfx", moment.SFX);
                        sceneNode.AppendChild(momentNode);
                    }
                }

            }




        }

        static void AddAttribute(XmlElement elementNode, string name, string value)
        {
            XmlAttribute att = sceneBreakdown.CreateAttribute(name);
            att.Value = value;
            elementNode.Attributes.Append(att);
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
