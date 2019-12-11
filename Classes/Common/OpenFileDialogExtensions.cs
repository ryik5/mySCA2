using System.Windows.Forms;

namespace ASTA.Classes
{
    public class OpenFileDialogExtentions
    {
        public string FilePath { get; private set; }


        public  OpenFileDialogExtentions(string titleWindowDialog = null, string maskFiles = "Все файлы (*.*)|*.*")
        {
            string filePath = null;
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.FileName = "";

                if (titleWindowDialog != null)
                { openFileDialog1.Title = titleWindowDialog; }

                openFileDialog1.Filter = maskFiles;

                DialogResult res = openFileDialog1.ShowDialog();
                if (res != DialogResult.Cancel)
                { filePath = openFileDialog1.FileName; }
            }
            FilePath = filePath;
        }
    }
}