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
            startFormContent.AddToOriginal = cB_OutToOriginal.Checked;

            Ausfuehren.Run(startFormContent);



            Close();
        }

        private void bt_AlleTeams_Click(object sender, EventArgs e)
        {
            int[] teamListe= { 1,2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,22,23,25,26,27,28,29,30,31,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49};
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
            startFormContent.StartDate = dateTimePicker1.Text;
            startFormContent.AddToOriginal = true;

                foreach (var team in teamListe)
                {

                    startFormContent.TeamNumber = team;
                    Ausfuehren.Run(startFormContent);

                }
                Close();
        }
    }
}
