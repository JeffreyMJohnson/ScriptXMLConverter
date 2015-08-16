using System;
using System.Collections.Generic;
using Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptXMLConvert
{
    public class XLSXDataSource : IScriptDataProvider
    {
        private List<SheetRow> _rows = new List<SheetRow>();


        public XLSXDataSource(string a_path)
        {
            Path = a_path;
            IEnumerator<worksheet> sheets = LoadWorkbook(a_path);
            LoadRows(sheets);
        }

        public string Path { get; private set; }

        public SheetRow[] GetRows()
        {
            return _rows.ToArray();
        }

        private IEnumerator<worksheet> LoadWorkbook(string path)
        {
            IEnumerator<worksheet> sheets = Workbook.Worksheets(path).GetEnumerator();

            //verify file found
            try
            {
                sheets.MoveNext();
            }
            catch (System.IO.FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Is '" + path + "' located in working directory?");
                throw e;
            }
            return sheets;
        }

        private void LoadRows(IEnumerator<worksheet> sheets)
        {
            Row[] rows = sheets.Current.Rows;
            foreach (Row row in rows)
            {
                //skip blank rows
                if (IsBlankRow(row))
                {
                    continue;
                }
                _rows.Add(new SheetRow
                {
                    Scene = GetCellText(row, ColumnHeader.Scene),
                    Moment = GetCellText(row, ColumnHeader.Moment),
                    Line = GetCellText(row, ColumnHeader.Line),
                    Duration = GetCellText(row, ColumnHeader.Duration),
                    Location = GetCellText(row, ColumnHeader.Location),
                    SFX = GetCellText(row, ColumnHeader.SFX)
                });
                if (GetCellText(row, ColumnHeader.Scene) == "SCRIPT TOTAL DURATION")
                {
                    //all done last line
                    break;
                }
            }

        }

        private bool IsBlankRow(Row row)
        {
            return GetCellText(row, ColumnHeader.Scene) == "";
        }

        private string GetCellText(Row row, ColumnHeader column)
        {
            if (null == row.Cells[(int)column])
            {
                return "";
            }
            return row.Cells[(int)column].Text;
        }



        /*
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
                */


    }
}
