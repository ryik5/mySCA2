using System;
using System.Linq;
using System.Windows.Forms;

namespace ASTA
{
    partial class AboutBox1 : Form
    {
        private bool boolButtonOk = false;

        public AboutBox1()
        {
            InitializeComponent();

            System.Diagnostics.FileVersionInfo fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);

            foreach (Label label in tableLayoutPanel.Controls.OfType<Label>())
            { label.Font = new System.Drawing.Font("Courier New", 8, System.Drawing.FontStyle.Regular); }

            string ver1 = System.Reflection.Assembly.GetEntryAssembly().GetName().Version.ToString();
            // string ver3 = Application.ProductVersion;

            Text = $"{"О программе:",-14}{fileVersionInfo.ProductName}";
            labelProductName.Text = $"{"Название:",-14}{fileVersionInfo.Comments}";
            labelVersion.Text = $"{"Версия:",-14}{ver1,0} ({fileVersionInfo.FileVersion,0})";
            labelCopyright.Text = $"{"Разработчик:",-14}{fileVersionInfo.LegalCopyright}";
            labelPath.Text = $"{"Полный путь:",-14}{Application.ExecutablePath}";
            labelFile.Text = $"{"Имя файла:",-14}{fileVersionInfo.OriginalFilename}";
            labelPC.Text = $"{"Имя ПК:",-14}{Environment.MachineName}";
            labelOS.Text = $"{"Версия ОС:",-14}{Environment.OSVersion}";
            labelOcupiedRAM.Text = $"{"Память,MB:",-14}{(Environment.WorkingSet / 1024 / 1024).ToString()}";
            toolTip1.SetToolTip(labelVersion, "Версия сборки и версия файла");
            toolTip1.SetToolTip(labelOcupiedRAM, "Объем памяти в RAM, занятый приложением");
            toolTip1.SetToolTip(labelPath, Application.ExecutablePath);
        }

        public bool OKButtonClicked
        {
            get { return boolButtonOk; }
        }

        private void buttonYes_Click(object sender, EventArgs e)
        {
            boolButtonOk = true;
            Close();
        }

        private void buttonNo_Click(object sender, EventArgs e)
        {
            boolButtonOk = false;
            Close();
        }

        private void logoPictureBox_Click(object sender, EventArgs e)
        {
            const string url = "mailto:ryik.yuri@gmail.com";
            System.Diagnostics.Process.Start(url);
        }
    }
}