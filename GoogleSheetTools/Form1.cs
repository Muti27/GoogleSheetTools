using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;

namespace GoogleSheetTools
{
    public partial class MainForm : Form
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        private SheetsService Service;

        public MainForm()
        {
            InitializeComponent();
        }

        private void OnClickConvert(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox.Text))
            {
                return;
            }

            ConvertSheet(comboBox.Text);
        }

        private void OnClickSearch(object sender, EventArgs e)
        {
            if (CheckService() == false)
                return;

            if (comboBox.Items.Count > 0)
                comboBox.Items.Clear();

            var sheetId = textBox1.Text;
            var sheetRequest = Service.Spreadsheets.Get(sheetId);
            var sheet = sheetRequest.Execute();

            var sheetsList = sheet.Sheets;
            for (int i = 0; i < sheetsList.Count; i++)
            {
                comboBox.Items.Add(sheetsList[i].Properties.Title);
            }
        }

        private void ConvertSheet(string sheetName)
        {
            if (CheckService() == false)
                return;
                        
            var sheetId = textBox1.Text;           
            var sheetRequest = Service.Spreadsheets.Values.Get(sheetId, sheetName);
            var sheet = sheetRequest.Execute();
            IList<IList<Object>> rows = sheet.Values;

            string path = "Output";
            if (Directory.Exists(path) == false)
            {
                Directory.CreateDirectory(path);
            }

            using (StreamWriter stringWriter = new StreamWriter(Path.Combine(path, sheetName + ".xml")))
            {
                //紀錄資料標頭及開始位置
                List<string> title = new List<string>();
                int startIndex = 0;

                for (int i = 0; i < rows[0].Count; i++)
                {
                    if (title.Count == 0)
                    {
                        //尋找ID作為開頭
                        if (rows[0][i].ToString() != "ID")
                            continue;

                        startIndex = i;
                        Console.WriteLine("Start Index: " + startIndex);
                    }

                    title.Add(rows[0][i].ToString());
                    Console.WriteLine("Title: " + rows[0][i].ToString());
                }

                stringWriter.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                stringWriter.WriteLine("<ns1:CoreObjectRoot xmlns:ns1=\"http://www.moregeek.com/Scarlet/" + sheetName + "/XMLSchema\"");

                for (int i = 1; i < rows.Count; i++)
                {
                    //無資料跳過
                    if (rows[i].Count == 0)
                        continue;

                    //ID空跳過
                    if (string.IsNullOrEmpty(rows[i][startIndex].ToString()))
                        continue;

                    Console.WriteLine(string.Format("   <{0} {1}=\"{2}\">", sheetName, title[0], rows[i][startIndex]));
                    stringWriter.WriteLine(string.Format("  <{0} {1}=\"{2}\">", sheetName, title[0], rows[i][startIndex]));

                    for (int k = 1, j = startIndex + 1; k < title.Count; k++, j++)
                    {
                        if (j >= rows[i].Count || string.IsNullOrEmpty(rows[i][j].ToString()))
                        {
                            Console.WriteLine(string.Format("       <{0}/>", title[k]));
                            stringWriter.WriteLine(string.Format("      <{0}/>", title[k]));
                            continue;
                        }

                        Console.WriteLine(string.Format("       <{0}>{1}</{0}>", title[k], rows[i][j].ToString()));
                        stringWriter.WriteLine(string.Format("      <{0}>{1}</{0}>", title[k], rows[i][j].ToString()));
                    }

                    Console.WriteLine(string.Format("   </{0}>", sheetName));
                    stringWriter.WriteLine(string.Format("  </{0}>", sheetName));
                }

                stringWriter.WriteLine("</ns1:CoreObjectRoot>");
            }

            MessageBox.Show("轉檔成功!");
        }

        private bool CheckService()
        {
            if (Service != null)
                return true;

            try
            {
                UserCredential credential;
                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    /* The file token.json stores the user's access and refresh tokens, and is created
                     automatically when the authorization flow completes for the first time. */
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                Service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "GoogleSheetTools"
                });

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
