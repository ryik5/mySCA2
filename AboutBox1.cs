using System;
using System.Windows.Forms;
using System.Linq;


namespace ASTA
{
    partial class AboutBox1 : Form
    {
        public bool boolButtonOk = false;

        public AboutBox1()
        {
            InitializeComponent();

            System.Diagnostics.FileVersionInfo fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);

            foreach (Label label in tableLayoutPanel.Controls.OfType<Label>())
            { label.Font = new System.Drawing.Font("Courier New", 8, System.Drawing.FontStyle.Regular); }

            string ver1 = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString();
            string ver2 = fileVersionInfo.FileVersion;
           // string ver3 = Application.ProductVersion;

            this.Text = String.Format("{0,-14}{1}", "О программе:", fileVersionInfo.ProductName);
            this.labelProductName.Text = String.Format("{0,-14}{1}", "Название:", fileVersionInfo.Comments);
            this.labelVersion.Text = String.Format("{0,-14}{1,0} ({2,0})", "Версия:", ver1, ver2);
            this.labelCopyright.Text = String.Format("{0,-14}{1}", "Разработчик:", fileVersionInfo.LegalCopyright);
            this.labelPath.Text = String.Format("{0,-14}{1}", "Полный путь:", Application.ExecutablePath);
            this.labelFile.Text = String.Format("{0,-14}{1}", "Имя файла:", fileVersionInfo.OriginalFilename);
            this.labelPC.Text = String.Format("{0,-14}{1}", "Имя ПК:", Environment.MachineName.ToString());
            this.labelOS.Text = String.Format("{0,-14}{1}", "Верися ОС:", Environment.OSVersion.ToString());
            this.labelOcupiedRAM.Text = String.Format("{0,-14}{1}", "Память, MB:", (Environment.WorkingSet / 1024 / 1024).ToString());
            this.toolTip1.SetToolTip(labelVersion, "Версия сборки и версия файла");
            this.toolTip1.SetToolTip(labelOcupiedRAM, "Занятый приложением объем памяти в RAM");
            this.toolTip1.SetToolTip(labelPath, Application.ExecutablePath);
        }

        public bool OKButtonClicked
        {
            get { return boolButtonOk; }
        }

        private void buttonYes_Click(object sender, EventArgs e)
        {
            boolButtonOk = true;
            this.Close();
        }

        private void buttonNo_Click(object sender, EventArgs e)
        {
            boolButtonOk = false;
            this.Close();
        }

        private void logoPictureBox_Click(object sender, EventArgs e)
        {
            string url = "mailto:ryik.yuri@gmail.com";
            System.Diagnostics.Process.Start(url);
        }
    }
}
