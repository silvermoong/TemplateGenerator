using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Template
{
    public partial class Form1 : Form
    {
        private int index = 1;
        public Form1()
        {
            InitializeComponent();
            LoadTemplates();
            AddOnClick();
        }

        private void LoadTemplates()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string sourceDirectory = path + @"\Templates";
            if (!Directory.Exists(sourceDirectory))
                Directory.CreateDirectory(sourceDirectory);
            string[] fileEntries = Directory.GetFiles(sourceDirectory);

            foreach (string fileName in fileEntries)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = Path.GetFileName(fileName);
                item.Value = fileName;

                cbSelect.Items.Add(item);
            }

            if(cbSelect.Items.Count > 0)
            {
                cbSelect.SelectedIndex = 0;
            }
        }

        private void AddOnClick()
        {
            foreach (QuestionUserControl c in tableLayoutPanelRight.Controls.OfType<QuestionUserControl>())
            {
                c.OnUserControlButtonClicked += (s, e) => AddNewLine(c.Question);
            }
        }

        private void AddNewLine(string s)
        {
            txtTemplate.AppendText(index + ". " + s + Environment.NewLine);
            txtTemplate.AppendText("" + Environment.NewLine);
            txtTemplate.AppendText("" + Environment.NewLine);
            index++;
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtTemplate.Text);
        }

        private void cbSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (QuestionUserControl c in tableLayoutPanelRight.Controls.OfType<QuestionUserControl>())
            {
                c.Question = "";
            }
            string line;
            int counter = 1;
            StreamReader file = new StreamReader(((ComboboxItem)((ComboBox)sender).SelectedItem).Value);
            while ((line = file.ReadLine()) != null)
            {
                QuestionUserControl c = (QuestionUserControl)FindControlRecursive(tableLayoutPanelRight, "questionUserControl" + counter);
                c.Question = line;
                counter++;
            }
        }

        private Control FindControlRecursive(Control rootControl, string controlID)
        {
            if (rootControl.Name == controlID) return rootControl;

            foreach (Control controlToSearch in rootControl.Controls)
            {
                Control controlToReturn = FindControlRecursive(controlToSearch, controlID);
                if (controlToReturn != null) return controlToReturn;
            }
            return null;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            txtTemplate.Text = "";
            index = 1;
        }


        /* Method to create a table format string which can directly be set to RichTextBoxControl.Rows,
        columns and cell width are passed as parameters rather than hard coding as in previous example.*/
        private static String InsertTableInRichTextBox(int rows, int cols, int width, List<string> data)
        {
            //Create StringBuilder Instance
            StringBuilder sringTableRtf = new StringBuilder();

            //beginning of rich text format
            sringTableRtf.Append(@"{\rtf1 ");

            //Variable for cell width
            int cellWidth;

            //Start row
            sringTableRtf.Append(@"\trowd");

            //Loop to create table string
            for (int i = 0; i < rows; i++)
            {
                sringTableRtf.Append(@"\trowd");

                for (int j = 0; j < cols; j++)
                {
                    //Calculate cell end point for each cell
                    cellWidth = (j + 1) * width;

                    //A cell with width 1000 in each iteration.
                    sringTableRtf.Append(@"\cellx" + cellWidth.ToString() + data[1]);
                }

                //Append the row in StringBuilder
                sringTableRtf.Append(@"\intbl \cell \row");
            }
            sringTableRtf.Append(@"\pard");
            sringTableRtf.Append(@"}");

            return sringTableRtf.ToString();
        }

        private void btnCopyTable_Click(object sender, EventArgs e)
        {
            List<string> temp = new List<string>();

            using (StringReader reader = new StringReader(txtTemplate.Text))
            {
                string line = string.Empty;
                do
                {
                    line = reader.ReadLine();
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        temp.Add(line);
                    }

                } while (line != null);
            }
            //txtTemplate.Rtf = InsertTableInRichTextBox(5, 2, 1000, temp);
            CreateEmail();
        }

        private void CreateEmail()
        {
            try
            {
                // Create the Outlook application by using inline initialization.
                Outlook.Application oApp = new Outlook.Application();

                //Create the new message by using the simplest approach.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                //Add a recipient.
                // TODO: Change the following recipient where appropriate.
                //Outlook.Recipient oRecip = (Outlook.Recipient)oMsg.Recipients.Add("e-mail address");
                //oRecip.Resolve();

                //Set the basic properties.
                //oMsg.Subject = "";
                oMsg.HTMLBody = "<html><body><p>Please plan to present your status for the following projects...</p></body></html";
                //oMsg.Body = txtTemplate.Text;

                //Add an attachment.
                // TODO: change file path where appropriate
                //String sSource = "C:\\setupxlg.txt";
                //String sDisplayName = "MyFirstAttachment";
                //int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //Outlook.Attachment oAttach = oMsg.Attachments.Add(sSource, iAttachType, iPosition, sDisplayName);

                // If you want to, display the message.
                // oMsg.Display(true);  //modal

                //Send the message.
                oMsg.Save();
                oMsg.Send();

                //Explicitly release objects.
                //oRecip = null;
                //oAttach = null;
                oMsg = null;
                oApp = null;
            }

            // Simple error handler.
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);
            }
        }
    }

    public class ComboboxItem
    {
        public string Text { get; set; }
        public string Value { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }
}
