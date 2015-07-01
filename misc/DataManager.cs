using UnityEngine;
using System.Collections.Generic;
using System.Xml;
using System;

public class DataManager : MonoBehaviour
{
    public class Moment
    {
        public string Title { get; set; }
        public string Line { get; set; }

        public Moment(string a_title, string a_line)
        {
            Title = a_title;
            Line = a_line;
        }
    }

    public class Scene
    {
        public List<Moment> moments;
        public int Number { get; set; }

        public Scene()
        {
            moments = new List<Moment>();
        }

        public Scene(List<Moment> momentList)
        {
            moments = momentList;
        }

    }

    public class Act
    {
        public List<Scene> scenes;
        public int Number { get; set; }

        public Act()
        {
            scenes = new List<Scene>();
        }

        public Act(List<Scene> sceneList)
        {
            scenes = sceneList;
        }

        public Scene GetScene(int sceneNum)
        {
            foreach(Scene scene in scenes)
            {
                if(scene.Number == sceneNum)
                {
                    return scene;
                }
            }
            return null;
        }


    }

    public class Script
    {
        public List<Act> acts;

        public Script()
        {
            acts = new List<Act>();
        }

        public Script(List<Act> actList)
        {
            acts = actList;
        }

        public Act GetAct(int actNumber)
        {
            foreach (Act act in acts)
            {
                if (act.Number == actNumber)
                {
                    return act;
                }
            }
            return null;
        }
    }

    public List<Act> Acts
    {
        get
        {
            return m_Script.acts;
        }
    }

    public Act GetAct(int actNumber)
    {
        return m_Script.GetAct(actNumber);
    }

    private XmlDocument mDataDoc = new XmlDocument();
    private Script m_Script = new Script();



    private void Start()
    {
        mDataDoc.Load("SceneBreakdown.xml");
        if (mDataDoc.ChildNodes.Count != 1)
        {
            Debug.Log("Data file did not load.");
        }

        //loop acts and add to script
        XmlNodeList actsNodeList = mDataDoc.SelectNodes("script/act");
        Debug.Log("loading " + actsNodeList.Count + " acts...");
        foreach (XmlElement act in actsNodeList)
        {
            Act newAct = new Act();
            newAct.Number = Int32.Parse(act.Attributes.GetNamedItem("number").Value);
            //loop scenes and add to act
            XmlNodeList scenesNodeList = act.SelectNodes("scene");
            Debug.Log("loading " + scenesNodeList.Count + " scenes...");
            foreach (XmlElement scene in scenesNodeList)
            {
                Scene newScene = new Scene();
                newScene.Number = Int32.Parse(scene.Attributes.GetNamedItem("number").Value);
                //loop moments and add to scene
                XmlNodeList momentsNodeList = scene.SelectNodes("moment");
                Debug.Log("loading " + momentsNodeList.Count + " moments...");
                foreach (XmlElement moment in momentsNodeList)
                {
                    string title = moment.Attributes.GetNamedItem("title").Value;
                    string line = moment.Attributes.GetNamedItem("line").Value;

                    Moment newMoment = new Moment(title, line);
                    newScene.moments.Add(newMoment);
                }
                newAct.scenes.Add(newScene);
            }
            m_Script.acts.Add(newAct);
        }


    }



}
