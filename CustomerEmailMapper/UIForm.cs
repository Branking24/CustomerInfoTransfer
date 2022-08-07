using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CustomerEmailMapper
{
    public partial class UIForm : Form
    {
        public string spreadsheetID;
        public string mailKey;
        public string clientID;
        public string secret;
        public string sheetName;
        public string folderName;
        public UIForm()
        {
            InitializeComponent();
            string filename = "C:\\Program Files (x86)\\CERResources\\Setup.txt";
            spreadsheetID = null;
            mailKey = null;
            clientID = null;
            secret = null;
            sheetName = null;
            folderName = null;
            var lines = File.ReadLines(filename);
            foreach (var line in lines)
            {
                if (line.StartsWith("SpreadsheetID:"))
                {
                    spreadsheetID = line.Remove(0, 15);
                }
                if (line.StartsWith("MailKey:"))
                {
                    mailKey = line.Remove(0, 9);
                }
                if (line.StartsWith("ClientID:"))
                {
                    clientID = line.Remove(0, 10);
                }
                if (line.StartsWith("ClientSecret:"))
                {
                    secret = line.Remove(0, 14);
                }
                if (line.StartsWith("SheetName:"))
                {
                    sheetName = line.Remove(0, 11);
                }
                if (line.StartsWith("GmailFolderName:"))
                {
                    folderName = line.Remove(0, 17);
                }
            }
        }

        private void UIForm_Load(object sender, EventArgs e)
        {

        }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            if (spreadsheetID == null)
            {
                WriteStatus("No Spreadsheet ID Found");
            }
            else if (mailKey == null)
            {
                WriteStatus("No Mail Key Found");
            }
            else if (clientID == null)
            {
                WriteStatus("No ClientID Found");
            }
            else if (secret == null)
            {
                WriteStatus("No Client Secret Found");
            }
            else if (sheetName == null)
            {
                WriteStatus("No Sheet Name Found");
            }
            else if (folderName == null)
            {
                WriteStatus("No Gmail Folder Name Found");
            }
            else
            {
                try
                {
                    string username = this.UsernameBox.Text;
                    EmailFetcher fetcher = new EmailFetcher(username, mailKey, this, folderName);
                    List<CustomerInfoTemplate> info = fetcher.FetchEmailInfo();
                    CsvClient csvClient = new CsvClient();
                    try
                    {
                        csvClient.AddToSheet(spreadsheetID, this, info, clientID, secret, sheetName);
                    }
                    catch (Exception ex)
                    {
                        if (ex is Google.GoogleApiException)
                        {
                            fetcher.MarkAsUnseen();
                        }
                        Console.WriteLine(ex.Message);
                    }
                    fetcher.MoveMessages();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        public void WriteStatus(string message)
        {
            this.StatusLabel.Text = message;
        }
    }
}
