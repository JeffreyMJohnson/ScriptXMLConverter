﻿using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
using System.Xml;


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



        [STAThread]
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
            System.Windows.Forms.Clipboard.Clear();
            Console.WriteLine(authorizationUrl);
            System.Windows.Forms.Clipboard.SetText(authorizationUrl);
            Console.WriteLine("Please visit the URL above to authorize your OAuth " + "request token.  Once that is complete, type in your access code to "
                + "continue...");
            parameters.AccessCode = Console.ReadLine();

            OAuthUtil.GetAccessToken(parameters);
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
                // TODO: There were no spreadsheets, act accordingly.
            }

            // TODO: Choose a spreadsheet more intelligently based on your
            // app's needs.
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry)feed.Entries[0];
            Console.WriteLine(spreadsheet.Title.Text);

            // Get the first worksheet of the first spreadsheet.
            // TODO: Choose a worksheet more intelligently based on your
            // app's needs.
            WorksheetFeed wsFeed = spreadsheet.Worksheets;
            WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];

            // Define the URL to request the list feed of the worksheet.
            AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // Fetch the list feed of the worksheet.
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            ListFeed listFeed = service.Query(listQuery);
            return listFeed;
        }
        

        static Script LoadScript()
        {
            ListFeed listFeed = GetRows();
            Script script = new Script();
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
                    if(null != act)
                    {
                        act.AddScene(scene);
                        scene = null;
                        script.AddAct(act);
                    }
                    act = new Act();
                    act.Number = sceneValue.Substring(sceneValue.LastIndexOf(' '));
                    //go to next row
                    continue;
                }

                //if new scene
                if(sceneValue.Contains("TIME"))
                {
                    //if not first scene 
                    if(null != scene)
                    {
                        act.AddScene(scene);
                    }
                    scene = new Scene();
                    scene.Time = row.Elements[(int)ColumnHeader.Duration].Value;
                    //go to next row
                    continue;
                }

                //if last element
                if(sceneValue.Contains("SCRIPT TOTAL"))
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
                if(scene.Number != row.Elements[(int)ColumnHeader.Scene].Value)
                {
                    scene.Number = row.Elements[(int)ColumnHeader.Scene].Value;
                }
                scene.AddMoment(moment);
            }
            return script;

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

                    foreach(Moment moment in scene.Moments)
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
