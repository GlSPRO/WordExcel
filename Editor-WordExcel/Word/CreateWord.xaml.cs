using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace Editor_WordExcel.Word
{
    public partial class CreateWord : Window
    {
        public CreateWord()
        {
            InitializeComponent();
        }

        void SaveRtfFile(string _fileName)
        {
            TextRange range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            FileStream fst = new FileStream(_fileName, FileMode.Create);
            range.Save(fst, DataFormats.Rtf);
            fst.Close();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf";
            saveFileDialog.Title = "Выбери куда сохранить документ Word";

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                SaveRtfFile(filePath);
            }
        }

        private void Send_Click(object sender, RoutedEventArgs e)
        {
            SendOnline send = new SendOnline();
            send.Show();
        }
    }
}
