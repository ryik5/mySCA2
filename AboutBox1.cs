using System;
using System.Windows.Forms;
using System.Linq;


namespace mySCA2
{
    partial class AboutBox1 : Form
    {
        private bool boolButtonOk = false;

        public AboutBox1()
        {
            InitializeComponent();

            System.Diagnostics.FileVersionInfo myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);

            foreach (Label label in tableLayoutPanel.Controls.OfType<Label>())
            { label.Font = new System.Drawing.Font("Courier New", 8, System.Drawing.FontStyle.Regular); }

            this.Text = String.Format("{0,-14}{1}", "О программе:", myFileVersionInfo.ProductName);
            this.labelProductName.Text = String.Format("{0,-14}{1}", "Название:", myFileVersionInfo.Comments);
            this.labelVersion.Text = String.Format("{0,-14}{1}", "Версия:", myFileVersionInfo.FileVersion);
            this.labelCopyright.Text = String.Format("{0,-14}{1}", "Разработчик:", myFileVersionInfo.LegalCopyright);
            this.labelPath.Text = String.Format("{0,-14}{1}", "Полный путь:", Application.ExecutablePath);
            this.labelFile.Text = String.Format("{0,-14}{1}", "Имя файла:", myFileVersionInfo.OriginalFilename);
            this.labelPC.Text = String.Format("{0,-14}{1}", "Имя ПК:", Environment.MachineName.ToString());
            this.labelOS.Text = String.Format("{0,-14}{1}", "Верися ОС:", Environment.OSVersion.ToString());
            this.labelOcupiedRAM.Text = String.Format("{0,-14}{1}", "Память, MB:", (Environment.WorkingSet / 1024 / 1024).ToString());
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
    }
}
