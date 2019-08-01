using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CCUSpielplanAuslesen
{
    public partial class Form1 : Form
    {
        private string pathOfFile;
        private string pathOnly;
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "Auswahl Spielplan in Excel-Format";
            openDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openDialog.FilterIndex = 1;
            var result = openDialog.ShowDialog();
            if (result.ToString() == "OK")
            {
                pathOfFile = openDialog.FileName;
                textBox1.Text = pathOfFile;
                pathOnly = pathOfFile.Substring(0, pathOfFile.LastIndexOf(openDialog.SafeFileName));

            }
        }

        private void btn_StartAuslesen_Click(object sender, EventArgs e)
        {
            StartFormContent startFormContent = new StartFormContent();
            startFormContent.SourceFile = pathOfFile;
            startFormContent.TargetPath = pathOnly;
            int tempInt;
            bool ok = int.TryParse(tB_ErsteSpalte.Text, out tempInt);
            if (ok)
            {
                startFormContent.FirstColumn = tempInt;
            }
            ok = int.TryParse(tB_ErsteZeile.Text, out tempInt);
            if (ok)
            {
                startFormContent.FirstRow = tempInt;
            }
            else
            {
                startFormContent.FirstRow = 1;
            }
            ok = int.TryParse(tB_TeamNummer.Text, out tempInt);
            if (ok)
            {
                startFormContent.TeamNumber = tempInt;
            }
            startFormContent.StartDate = dateTimePicker1.Text;

            Ausfuehren.Run(startFormContent);



            Close();
        }
    }
}
