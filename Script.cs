using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ScriptXMLConvert
{
    public class Script
    {
        public string TotalTime { get; set; }
        public Act[] Acts { get { return _acts.ToArray(); } }

        private List<Act> _acts = new List<Act>();

        public Script() { }

        public Script(SheetRow[] rows)
        {
            Act act = null;
            Scene scene = null;

            //loop rows
            foreach (SheetRow row in rows)
            {
                //skip first header row
                if (row.Scene.Contains("SCENE") && row.Moment.Contains("MOMENT"))
                {
                    continue;
                }

                if (row.Scene.Contains("ACT "))
                {
                    //if not first
                    if (null != act)
                    {
                        act.AddScene(scene);
                        scene = null;
                        AddAct(act);
                    }
                    act = new Act();
                    act.Number = row.Scene.Substring(row.Scene.LastIndexOf(' ') + 1);
                    //go to next row
                    continue;
                }
                //if new scene
                if (row.Scene.Contains("TIME"))
                {
                    //if not first scene 
                    if (null != scene)
                    {
                        act.AddScene(scene);
                    }
                    scene = new Scene();
                    scene.Time = row.Duration;
                    //go to next row
                    continue;
                }
                //if last element
                if (row.Scene.Contains("SCRIPT TOTAL"))
                {
                    TotalTime = row.Duration;
                    //add last scene to last act
                    act.AddScene(scene);
                    //add last act to script
                    AddAct(act);
                    //all done no need to continue checking rows
                    break;
                }
                //not above so it's a new moment
                Moment moment = new Moment(row.Moment,
                                           row.Line,
                                           row.Duration,
                                           row.Location,
                                           row.SFX);
                if (scene.Number != row.Scene)
                {
                    scene.Number = row.Scene;
                }
                scene.AddMoment(moment);
            }
        }

        public void AddAct(Act act)
        {
            _acts.Add(act);
        }

        public XmlDocument GetXML()
        {
            XmlDocument sceneBreakdown = new XmlDocument();
            //create root
            XmlElement scriptNode = sceneBreakdown.CreateElement("script");
            AddAttribute(sceneBreakdown, scriptNode, "filmLength", TotalTime);
            sceneBreakdown.AppendChild(scriptNode);

            foreach (Act act in Acts)
            {
                XmlElement actNode = sceneBreakdown.CreateElement("act");
                AddAttribute(sceneBreakdown, actNode, "number", act.Number);
                scriptNode.AppendChild(actNode);

                foreach (Scene scene in act.Scenes)
                {
                    XmlElement sceneNode = sceneBreakdown.CreateElement("scene");
                    AddAttribute(sceneBreakdown, sceneNode, "number", scene.Number);
                    AddAttribute(sceneBreakdown, sceneNode, "time", scene.Time);
                    actNode.AppendChild(sceneNode);

                    foreach (Moment moment in scene.Moments)
                    {
                        XmlElement momentNode = sceneBreakdown.CreateElement("moment");
                        AddAttribute(sceneBreakdown, momentNode, "title", moment.Title);
                        AddAttribute(sceneBreakdown, momentNode, "line", moment.Line);
                        AddAttribute(sceneBreakdown, momentNode, "duration", moment.Duration);
                        AddAttribute(sceneBreakdown, momentNode, "location", moment.Location);
                        AddAttribute(sceneBreakdown, momentNode, "sfx", moment.SFX);
                        sceneNode.AppendChild(momentNode);
                    }
                }
            }

            return sceneBreakdown;
        }

        private void AddAttribute(XmlDocument xmlDoc, XmlElement elementNode, string name, string value)
        {
            XmlAttribute att = xmlDoc.CreateAttribute(name);
            att.Value = value;
            elementNode.Attributes.Append(att);
        }

    }

    public class Act
    {
        public string Number { get; set; }
        public Scene[] Scenes { get { return _scenes.ToArray(); } }

        private List<Scene> _scenes = new List<Scene>();

        public void AddScene(Scene scene)
        {
            _scenes.Add(scene);
        }

    }


    public class Scene
    {
        public string Number { get; set; }
        public string Time { get; set; }
        public Moment[] Moments { get { return _moments.ToArray(); } }

        private List<Moment> _moments = new List<Moment>();

        public void AddMoment(Moment moment)
        {
            _moments.Add(moment);
        }


    }


    public class Moment
    {
        public string Title { get; set; }
        public string Line { get; set; }
        public string Duration { get; set; }
        public string Location { get; set; }
        public string SFX { get; set; }

        public Moment(string a_title, string a_line, string a_duration, string a_location, string a_sfx)
        {
            Title = a_title;
            Line = a_line;
            Duration = a_duration;
            Location = a_location;
            SFX = a_sfx;
        }

    }

}
