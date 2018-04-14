using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordApplicaton = Microsoft.Office.Interop.Word.Application;
using WordDocument = Microsoft.Office.Interop.Word.Document;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;

namespace MatabYar_ToolsAddIn
{
    public partial class MatabYarToolsUserControl : UserControl
    {
        public MatabYarToolsUserControl()
        {
            InitializeComponent();
        }

        private void MatabYarToolsUserControl_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            WordApplicaton template_word = new WordApplicaton();
            WordDocument doc = new WordDocument();

            object fileName = @"C:\templates\Anomaly scan.docx";
            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            doc = template_word.Documents.Open(ref fileName,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            String read = string.Empty;
            List<string> data = new List<string>();
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                string temp = doc.Paragraphs[i + 1].Range.Text.Trim();
                if (temp != string.Empty)
                    data.Add(temp);
            }
            //((_WordDocument)doc).Close();
            //((_WordApplication)word).Quit();
            // insert text at the current cursor location
            Microsoft.Office.Interop.Word.Application
              word = Globals.ThisAddIn.Application;
            Word.Range selection = word.Selection.Range;
            selection.Text = data.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // format text
            StringBuilder buffer = new StringBuilder();
            buffer.AppendLine("teststr");
            // insert text at the current cursor location
            Microsoft.Office.Interop.Word.Application
              word = Globals.ThisAddIn.Application;
            Word.Range selection = word.Selection.Range;
            selection.Text = buffer.ToString();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
