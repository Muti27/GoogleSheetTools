using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace GoogleSheetTools
{
    public partial class MainForm : Form
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        private SheetsService Service;

        private string path = string.Empty;

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

            ConvertSheet();
        }

        private void OnClickSearch(object sender, EventArgs e)
        {
            if (CheckService() == false)
                return;

            if (comboBox.Items.Count > 0)
                comboBox.Items.Clear();

            var sheetId = urlText.Text;
            var sheetRequest = Service.Spreadsheets.Get(sheetId);
            var sheet = sheetRequest.Execute();

            var sheetsList = sheet.Sheets;
            for (int i = 0; i < sheetsList.Count; i++)
            {
                comboBox.Items.Add(sheetsList[i].Properties.Title);
            }
        }

        private void ConvertSheet()
        {
            if (CheckService() == false)
                return;

            path = "output";
            if (File.Exists("config.txt"))
            {
                StreamReader reader = File.OpenText("config.txt");
                path = reader.ReadToEnd();
                Console.WriteLine("path: " + path);

                reader.Close();
            }

            ConvertToJson();
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

        private void ConvertToJson(string sheetName = "")
        {
            if (string.IsNullOrEmpty(sheetName))
                sheetName = comboBox.Text;

            var sheetId = urlText.Text;            
            var sheetRequest = Service.Spreadsheets.Values.Get(sheetId, sheetName);
            var sheet = sheetRequest.Execute();

            if (Directory.Exists(path) == false)
            {
                Directory.CreateDirectory(path);
            }

            //紀錄資料標頭及開始位置
            List<string> title = new List<string>();
            int startIndex = 0;
            IList<IList<Object>> rows = sheet.Values;

            if (rows == null)
                return;

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

            using (StreamWriter stringWriter = new StreamWriter(Path.Combine(path, sheetName + ".json")))
            {
                List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();

                for (int i = 1; i < rows.Count; i++)
                {
                    //無資料跳過
                    if (rows[i].Count == 0)
                        continue;

                    //ID空跳過
                    if (string.IsNullOrEmpty(rows[i][startIndex].ToString()))
                        continue;

                    Dictionary<string, object> data = new Dictionary<string, object>();

                    for (int k = startIndex, j = 0; k < title.Count; k++, j++)
                    {
                        if (j >= rows[i].Count || string.IsNullOrEmpty(rows[i][j].ToString()))
                        {
                            data.Add(title[k], null);
                            continue;
                        }

                        int parse = 0;
                        if (int.TryParse(rows[i][j].ToString(), out parse))                        
                            data.Add(title[k], parse);
                        else
                            data.Add(title[k], rows[i][j]);
                    }

                    dataList.Add(data);
                }

                JsonSerializerSettings setting = new JsonSerializerSettings();
                setting.Formatting = Formatting.Indented;
                setting.NullValueHandling = NullValueHandling.Ignore;

                string json = JsonConvert.SerializeObject(dataList, setting);
                stringWriter.Write(json);
            }

            MessageBox.Show("轉檔成功!");
        }

        private void ConvertAll_Click(object sender, EventArgs e)
        {
            foreach (string sheetName in comboBox.Items)
            {
                if (sheetName == "規劃")
                    continue;

                ConvertToJson(sheetName);
            }

            MessageBox.Show("轉檔成功!");
        }
    }
}
