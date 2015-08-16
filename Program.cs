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
    /// 

    public enum ColumnHeader
    {
        Scene = 0,
        Moment,
        Line,
        Duration,
        Location,
        SFX
    }

    class Program
    {


        //this switch is for which to use for data, local excel sheet or google sheet in cloud.
        private static bool usingExcel = true;
        private const string EXCEL_SHEET_PATH = "SCENE BREAKDOWN - KANSAS.xlsx";
        
        [STAThread]
        static void Main(string[] args)
        {


            try
            {
                SheetRow[] rows = null;

                if (usingExcel)
                {
                    Console.WriteLine("Using .xlsx Excel sheet for data....");
                    rows = new XLSXDataSource(EXCEL_SHEET_PATH).GetRows();
                }
                else
                {
                    Console.WriteLine("Using Google sheet (on Google Drive) for data....");
                    rows = new GoogleDataSource().GetRows();
                }

                Script script = new Script(rows);
                Console.WriteLine("Saving xml data file...");
                script.GetXML().Save("SceneBreakdown.xml");
                Console.WriteLine("Application finished.");
                Console.WriteLine("press any key to terminate...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("Catastrophic failure occured:");
                Console.WriteLine("If this continues contact jefjohms@gmail.com");
                Console.WriteLine("press any key to terminate...");
                Console.ReadKey();
                return;
            }
        }
    }

}
