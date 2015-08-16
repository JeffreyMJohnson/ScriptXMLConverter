using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace ScriptXMLConvert
{
    public class GoogleDataSource : IScriptDataProvider
    {

        //OAuth config
        private string CLIENT_ID = "898242977449-gdhq44lj4h22jgv2gougnaktg6i482p9.apps.googleusercontent.com";
        private string CLIENT_SECRET = "kiR_ogLwko8r_HviWpSWyj2p";
        private string SCOPE = "https://spreadsheets.google.com/feeds";
        private string REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";

        private List<SheetRow> _rows = new List<SheetRow>();

        public GoogleDataSource()
        {
            ListFeed feedList = GetListFeed();
            LoadRows(feedList);
        }

        public SheetRow[] GetRows()
        {
            return _rows.ToArray();
        }

        private void LoadRows(ListFeed listFeed)
        {
            foreach (ListEntry row in listFeed.Entries)
            {
                //no blank rows with google data or data will stop
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

        private string GetCellText(ListEntry row, ColumnHeader column)
        {
            if(null == row.Elements[(int)column])
            {
                return "";
            }
            return row.Elements[(int)column].Value;
        }


        private ListFeed GetListFeed()
        {
            //    //setup OAuth object
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
                Console.WriteLine("*****URL Copied to clipboard.*****");
            }
            catch (System.Runtime.InteropServices.ExternalException e)
            {
                Console.WriteLine("*****There was a problem copying to the clipboard, so copy the URL manually.*****");
            }

            Console.WriteLine("Please visit the URL above to authorize your OAuth request token.  Once that is complete, type in your access code to continue...");

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
                Console.WriteLine("No Spreadsheets found on Google Drive.");
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
                Console.WriteLine("Could not find spreadsheet 'SCENE BREAKDOWN - KANSAS' on Google Drive.");
                Console.WriteLine("Did you login with account that has access?");
                throw new Exception();
            }
            else
            {
                Console.WriteLine("Spreadsheet 'SCENE BREAKDOWN - KANSAS' found on Google Drive.");
            }

            WorksheetFeed wsFeed = spreadsheet.Worksheets;
            WorksheetEntry worksheet =  (WorksheetEntry)wsFeed.Entries[0];
            // Define the URL to request the list feed of the worksheet.
            AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // Fetch the list feed of the worksheet.
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            ListFeed listFeed = service.Query(listQuery);
            if (listFeed.Entries.Count < 1)
            {
                Console.WriteLine("No rows returned with the data.\nDid you sign on with account that has access to spreadsheet on Google Drive?");
            }

            return listFeed;
        }

    }
}
