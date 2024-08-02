namespace CCUSpielplanAuslesen
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tB_TeamNummer = new System.Windows.Forms.TextBox();
            this.lbl_TeamNummer = new System.Windows.Forms.Label();
            this.tB_ErsteSpalte = new System.Windows.Forms.TextBox();
            this.lbl_ErsteSpalte = new System.Windows.Forms.Label();
            this.tB_ErsteZeile = new System.Windows.Forms.TextBox();
            this.lbl_ErsteZeile = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.lbl_StartDatum = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lbl_Spielplan = new System.Windows.Forms.Label();
            this.btn_StartAuslesen = new System.Windows.Forms.Button();
            this.cB_OutToOriginal = new System.Windows.Forms.CheckBox();
            this.lbl_OutToOriginal = new System.Windows.Forms.Label();
            this.bt_AlleTeams = new System.Windows.Forms.Button();
            this.TBteamliste = new System.Windows.Forms.TextBox();
            this.LBteamliste = new System.Windows.Forms.Label();
            this.cbFinalrundeAdd = new System.Windows.Forms.CheckBox();
            this.lbFinalrundeAdd = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.tbFinalFreitagStundenText = new System.Windows.Forms.TextBox();
            this.tbFinalSamstagStundenText = new System.Windows.Forms.TextBox();
            this.dateTimePicker3 = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // tB_TeamNummer
            // 
            this.tB_TeamNummer.Location = new System.Drawing.Point(127, 38);
            this.tB_TeamNummer.Name = "tB_TeamNummer";
            this.tB_TeamNummer.Size = new System.Drawing.Size(100, 20);
            this.tB_TeamNummer.TabIndex = 0;
            // 
            // lbl_TeamNummer
            // 
            this.lbl_TeamNummer.AutoSize = true;
            this.lbl_TeamNummer.Location = new System.Drawing.Point(13, 41);
            this.lbl_TeamNummer.Name = "lbl_TeamNummer";
            this.lbl_TeamNummer.Size = new System.Drawing.Size(79, 13);
            this.lbl_TeamNummer.TabIndex = 1;
            this.lbl_TeamNummer.Text = "Team Nummer:";
            // 
            // tB_ErsteSpalte
            // 
            this.tB_ErsteSpalte.Location = new System.Drawing.Point(127, 65);
            this.tB_ErsteSpalte.Name = "tB_ErsteSpalte";
            this.tB_ErsteSpalte.Size = new System.Drawing.Size(100, 20);
            this.tB_ErsteSpalte.TabIndex = 2;
            this.tB_ErsteSpalte.Text = "5";
            // 
            // lbl_ErsteSpalte
            // 
            this.lbl_ErsteSpalte.AutoSize = true;
            this.lbl_ErsteSpalte.Location = new System.Drawing.Point(13, 68);
            this.lbl_ErsteSpalte.Name = "lbl_ErsteSpalte";
            this.lbl_ErsteSpalte.Size = new System.Drawing.Size(92, 13);
            this.lbl_ErsteSpalte.TabIndex = 3;
            this.lbl_ErsteSpalte.Text = "Erste Spalte (E=5)";
            // 
            // tB_ErsteZeile
            // 
            this.tB_ErsteZeile.Location = new System.Drawing.Point(127, 92);
            this.tB_ErsteZeile.Name = "tB_ErsteZeile";
            this.tB_ErsteZeile.Size = new System.Drawing.Size(100, 20);
            this.tB_ErsteZeile.TabIndex = 4;
            this.tB_ErsteZeile.Text = "4";
            // 
            // lbl_ErsteZeile
            // 
            this.lbl_ErsteZeile.AutoSize = true;
            this.lbl_ErsteZeile.Location = new System.Drawing.Point(13, 95);
            this.lbl_ErsteZeile.Name = "lbl_ErsteZeile";
            this.lbl_ErsteZeile.Size = new System.Drawing.Size(57, 13);
            this.lbl_ErsteZeile.TabIndex = 5;
            this.lbl_ErsteZeile.Text = "Erste Zeile";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(127, 119);
            this.dateTimePicker1.MinDate = new System.DateTime(2015, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 6;
            // 
            // lbl_StartDatum
            // 
            this.lbl_StartDatum.AutoSize = true;
            this.lbl_StartDatum.Location = new System.Drawing.Point(13, 125);
            this.lbl_StartDatum.Name = "lbl_StartDatum";
            this.lbl_StartDatum.Size = new System.Drawing.Size(58, 13);
            this.lbl_StartDatum.TabIndex = 7;
            this.lbl_StartDatum.Text = "Startdatum";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(127, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(302, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.Click += new System.EventHandler(this.textBox1_Click);
            // 
            // lbl_Spielplan
            // 
            this.lbl_Spielplan.AutoSize = true;
            this.lbl_Spielplan.Location = new System.Drawing.Point(13, 15);
            this.lbl_Spielplan.Name = "lbl_Spielplan";
            this.lbl_Spielplan.Size = new System.Drawing.Size(50, 13);
            this.lbl_Spielplan.TabIndex = 9;
            this.lbl_Spielplan.Text = "Spielplan";
            // 
            // btn_StartAuslesen
            // 
            this.btn_StartAuslesen.Location = new System.Drawing.Point(127, 281);
            this.btn_StartAuslesen.Name = "btn_StartAuslesen";
            this.btn_StartAuslesen.Size = new System.Drawing.Size(200, 23);
            this.btn_StartAuslesen.TabIndex = 10;
            this.btn_StartAuslesen.Text = "Auslesen starten";
            this.btn_StartAuslesen.UseVisualStyleBackColor = true;
            this.btn_StartAuslesen.Click += new System.EventHandler(this.btn_StartAuslesen_Click);
            // 
            // cB_OutToOriginal
            // 
            this.cB_OutToOriginal.AutoSize = true;
            this.cB_OutToOriginal.Location = new System.Drawing.Point(127, 146);
            this.cB_OutToOriginal.Name = "cB_OutToOriginal";
            this.cB_OutToOriginal.Size = new System.Drawing.Size(15, 14);
            this.cB_OutToOriginal.TabIndex = 11;
            this.cB_OutToOriginal.UseVisualStyleBackColor = true;
            // 
            // lbl_OutToOriginal
            // 
            this.lbl_OutToOriginal.AutoSize = true;
            this.lbl_OutToOriginal.Location = new System.Drawing.Point(13, 146);
            this.lbl_OutToOriginal.Name = "lbl_OutToOriginal";
            this.lbl_OutToOriginal.Size = new System.Drawing.Size(97, 13);
            this.lbl_OutToOriginal.TabIndex = 12;
            this.lbl_OutToOriginal.Text = "Spielplan ergänzen";
            // 
            // bt_AlleTeams
            // 
            this.bt_AlleTeams.Location = new System.Drawing.Point(127, 311);
            this.bt_AlleTeams.Name = "bt_AlleTeams";
            this.bt_AlleTeams.Size = new System.Drawing.Size(200, 23);
            this.bt_AlleTeams.TabIndex = 13;
            this.bt_AlleTeams.Text = "Alle Teams in Excel";
            this.bt_AlleTeams.UseVisualStyleBackColor = true;
            this.bt_AlleTeams.Click += new System.EventHandler(this.bt_AlleTeams_Click);
            // 
            // TBteamliste
            // 
            this.TBteamliste.Enabled = false;
            this.TBteamliste.Location = new System.Drawing.Point(16, 366);
            this.TBteamliste.Multiline = true;
            this.TBteamliste.Name = "TBteamliste";
            this.TBteamliste.Size = new System.Drawing.Size(490, 48);
            this.TBteamliste.TabIndex = 14;
            // 
            // LBteamliste
            // 
            this.LBteamliste.AutoSize = true;
            this.LBteamliste.Location = new System.Drawing.Point(16, 347);
            this.LBteamliste.Name = "LBteamliste";
            this.LBteamliste.Size = new System.Drawing.Size(42, 13);
            this.LBteamliste.TabIndex = 15;
            this.LBteamliste.Text = "Teams:";
            // 
            // cbFinalrundeAdd
            // 
            this.cbFinalrundeAdd.AutoSize = true;
            this.cbFinalrundeAdd.Location = new System.Drawing.Point(127, 166);
            this.cbFinalrundeAdd.Name = "cbFinalrundeAdd";
            this.cbFinalrundeAdd.Size = new System.Drawing.Size(15, 14);
            this.cbFinalrundeAdd.TabIndex = 16;
            this.cbFinalrundeAdd.UseVisualStyleBackColor = true;
            // 
            // lbFinalrundeAdd
            // 
            this.lbFinalrundeAdd.AutoSize = true;
            this.lbFinalrundeAdd.Location = new System.Drawing.Point(13, 167);
            this.lbFinalrundeAdd.Name = "lbFinalrundeAdd";
            this.lbFinalrundeAdd.Size = new System.Drawing.Size(103, 13);
            this.lbFinalrundeAdd.TabIndex = 17;
            this.lbFinalrundeAdd.Text = "Finalrunde ergänzen";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(149, 166);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker2.TabIndex = 18;
            // 
            // tbFinalFreitagStundenText
            // 
            this.tbFinalFreitagStundenText.Location = new System.Drawing.Point(367, 165);
            this.tbFinalFreitagStundenText.Name = "tbFinalFreitagStundenText";
            this.tbFinalFreitagStundenText.Size = new System.Drawing.Size(100, 20);
            this.tbFinalFreitagStundenText.TabIndex = 19;
            this.tbFinalFreitagStundenText.Text = "18";
            // 
            // tbFinalSamstagStundenText
            // 
            this.tbFinalSamstagStundenText.Location = new System.Drawing.Point(367, 192);
            this.tbFinalSamstagStundenText.Name = "tbFinalSamstagStundenText";
            this.tbFinalSamstagStundenText.Size = new System.Drawing.Size(100, 20);
            this.tbFinalSamstagStundenText.TabIndex = 20;
            this.tbFinalSamstagStundenText.Text = "10";
            // 
            // dateTimePicker3
            // 
            this.dateTimePicker3.Location = new System.Drawing.Point(149, 193);
            this.dateTimePicker3.Name = "dateTimePicker3";
            this.dateTimePicker3.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker3.TabIndex = 21;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(518, 428);
            this.Controls.Add(this.dateTimePicker3);
            this.Controls.Add(this.tbFinalSamstagStundenText);
            this.Controls.Add(this.tbFinalFreitagStundenText);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.lbFinalrundeAdd);
            this.Controls.Add(this.cbFinalrundeAdd);
            this.Controls.Add(this.LBteamliste);
            this.Controls.Add(this.TBteamliste);
            this.Controls.Add(this.bt_AlleTeams);
            this.Controls.Add(this.lbl_OutToOriginal);
            this.Controls.Add(this.cB_OutToOriginal);
            this.Controls.Add(this.btn_StartAuslesen);
            this.Controls.Add(this.lbl_Spielplan);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.lbl_StartDatum);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.lbl_ErsteZeile);
            this.Controls.Add(this.tB_ErsteZeile);
            this.Controls.Add(this.lbl_ErsteSpalte);
            this.Controls.Add(this.tB_ErsteSpalte);
            this.Controls.Add(this.lbl_TeamNummer);
            this.Controls.Add(this.tB_TeamNummer);
            this.Name = "Form1";
            this.Text = "CCU Spielplan Daten auslesen";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tB_TeamNummer;
        private System.Windows.Forms.Label lbl_TeamNummer;
        private System.Windows.Forms.TextBox tB_ErsteSpalte;
        private System.Windows.Forms.Label lbl_ErsteSpalte;
        private System.Windows.Forms.TextBox tB_ErsteZeile;
        private System.Windows.Forms.Label lbl_ErsteZeile;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label lbl_StartDatum;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label lbl_Spielplan;
        private System.Windows.Forms.Button btn_StartAuslesen;
        private System.Windows.Forms.CheckBox cB_OutToOriginal;
        private System.Windows.Forms.Label lbl_OutToOriginal;
        private System.Windows.Forms.Button bt_AlleTeams;
        private System.Windows.Forms.TextBox TBteamliste;
        private System.Windows.Forms.Label LBteamliste;
        private System.Windows.Forms.CheckBox cbFinalrundeAdd;
        private System.Windows.Forms.Label lbFinalrundeAdd;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.TextBox tbFinalFreitagStundenText;
        private System.Windows.Forms.TextBox tbFinalSamstagStundenText;
        private System.Windows.Forms.DateTimePicker dateTimePicker3;
    }
}

