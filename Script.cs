using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptXMLConvert
{
    public class Script
    {
        public string TotalTime { get; set; }
        public Act[] Acts { get { return _acts.ToArray(); } }

        private List<Act> _acts = new List<Act>();

        public void AddAct(Act act)
        {
            _acts.Add(act);
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
