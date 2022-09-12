using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using System.Linq;

namespace GoogleSheetTools
{
    public partial class MainForm : Form
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        private SheetsService Service;
        private string Path;
        private Dictionary<string, string> Sheets;

        public MainForm()
        {
            InitializeComponent();

            LoadConfig();
        }

        private void LoadConfig()
        {
            if (Sheets == null)
                Sheets = new Dictionary<string, string>();

            if (Sheets.Count > 0)
                Sheets.Clear();

            string configPath = "config.txt";
            if (File.Exists(configPath))
            {
                using (StreamReader reader = new StreamReader(configPath))
                {
                    while(reader.EndOfStream == false)
                    {
                        string line = reader.ReadLine();
                        
                        if(line.Contains("Path"))                        
                            Path = line.Substring(5);
                        
                        if(line.Contains(","))
                        {
                            var sheet = line.Split(',');

                            Sheets.Add(sheet[0], sheet[1]);
                        }
                    }
                }
            }
            else
            {
                Path = "output";
            }

            if (Sheets.Count > 0)
            {
                foreach(var item in Sheets)
                    urlComboBox.Items.Add(item.Key);
            }
        }

        private void OnClickConvert(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(comboBox.Text))
            {
                return;
            }

            ConvertSheet();
        }

        private void OnClickSearch(object sender, EventArgs e)
        {
            if (CheckService() == false)
                return;

            if (comboBox.Items.Count > 0)
            {
                comboBox.Items.Clear();
                comboBox.Text = string.Empty;
            }

            //先抓urlText欄位
            var sheetId = sheetUrlText.Text;
            if (string.IsNullOrEmpty(sheetId))
            {
                if (Sheets.ContainsKey(urlComboBox.Text))
                    sheetId = Sheets[urlComboBox.Text];                
            }

            //還是空則跳出
            if (string.IsNullOrEmpty(sheetId))
                return;

            try
            {
                var sheetRequest = Service.Spreadsheets.Get(sheetId);
                var sheet = sheetRequest.Execute();

                var sheetsList = sheet.Sheets;
                for (int i = 0; i < sheetsList.Count; i++)
                {
                    comboBox.Items.Add(sheetsList[i].Properties.Title);
                }
            }
            catch(Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private void ConvertSheet()
        {
            if (CheckService() == false)
                return;

            //先抓urlText欄位
            var sheetId = sheetUrlText.Text;
            if (string.IsNullOrEmpty(sheetId))
            {
                if (Sheets.ContainsKey(urlComboBox.Text))
                    sheetId = Sheets[urlComboBox.Text];                
            }

            //還是空則跳出
            if (string.IsNullOrEmpty(sheetId))
                return;

            try
            {
                var sheetName = comboBox.Text;
                var sheetRequest = Service.Spreadsheets.Values.Get(sheetId, sheetName);
                var sheet = sheetRequest.Execute();
                IList<IList<Object>> rows = sheet.Values;

                if (Directory.Exists(Path) == false)
                {
                    Directory.CreateDirectory(Path);
                }

                using (StreamWriter stringWriter = new StreamWriter(System.IO.Path.Combine(Path, sheetName + ".xml")))
                {
                    //紀錄資料標頭及開始位置
                    List<string> title = new List<string>();
                    int startIndex = 0;

                    for (int i = 0; i < rows[0].Count; i++)
                    {
                        if (title.Count == 0)
                        {
                            //尋找ID作為開頭
                            if (rows[0][i].ToString().ToUpper() != "ID")
                                continue;

                            startIndex = i;
                            Console.WriteLine("Start Index: " + startIndex);
                        }

                        title.Add(rows[0][i].ToString());
                        Console.WriteLine("Title: " + rows[0][i].ToString());
                    }

                    stringWriter.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    stringWriter.WriteLine("GameObject filename=\"" + sheetName + "\"");

                    for (int i = 1; i < rows.Count; i++)
                    {
                        //無資料跳過
                        if (rows[i].Count == 0)
                            continue;

                        //ID空跳過
                        if (string.IsNullOrEmpty(rows[i][startIndex].ToString()))
                            continue;

                        //備註欄跳過
                        int id = 0;
                        if (int.TryParse(rows[i][startIndex].ToString(), out id) == false)
                            continue;

                        stringWriter.WriteLine(string.Format("  <{0} {1}=\"{2}\">", sheetName, title[0], rows[i][startIndex]));

                        for (int k = 1, j = startIndex + 1; k < title.Count; k++, j++)
                        {
                            if (j >= rows[i].Count || string.IsNullOrEmpty(rows[i][j].ToString()))
                            {
                                stringWriter.WriteLine(string.Format("      <{0}/>", title[k]));
                                continue;
                            }

                            stringWriter.WriteLine(string.Format("      <{0}>{1}</{0}>", title[k], rows[i][j].ToString()));
                        }

                        stringWriter.WriteLine(string.Format("  </{0}>", sheetName));
                    }

                    stringWriter.WriteLine("</GameObject>");
                }

                MessageBox.Show("轉檔成功!");
            }
            catch
            {

            }
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
