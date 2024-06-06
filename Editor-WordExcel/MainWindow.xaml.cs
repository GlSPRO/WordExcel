using Editor_WordExcel.Excel;
using Editor_WordExcel.Word;
using Microsoft.Win32;
using Spire.Doc;
using Spire.Xls;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace Editor_WordExcel
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenWord(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents (*.docx)|*.docx";
            openFileDialog.Title = "Выбери документик";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string fileExtension = Path.GetExtension(filePath);

                if (fileExtension.ToLower() == ".docx")
                {
                    Document doc = new Document();
                    doc.LoadFromFile(filePath);
                    doc.SaveToFile("convert.rtf", Spire.Doc.FileFormat.Rtf);

                    var createWordWindow = new CreateWord();
                    var rtb = createWordWindow.FindName("rtb") as RichTextBox;
                    var fs = new FileStream("convert.rtf", FileMode.OpenOrCreate);

                    rtb.Document.Blocks.Clear();
                    rtb.Selection.Load(fs, DataFormats.Rtf);
                    fs.Close();
                    createWordWindow.ShowDialog();                                                                                                                                                                  
                }
                else
                {
                    MessageBox.Show("Выбранный файл не в формате docx", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }      

        private void CreateWord(object sender, RoutedEventArgs e)
        {
            CreateWord window = new CreateWord();
            window.Show();          
        }

        private void OpenExcel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog.Title = "Выберите файл Excel";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string fileExtension = Path.GetExtension(filePath);

                if (fileExtension.ToLower() == ".xls" || fileExtension.ToLower() == ".xlsx")
                {
                    Workbook wb = new Workbook();
                    wb.LoadFromFile(filePath);
                    Worksheet sheet = wb.Worksheets[0];
                    CellRange localRange = sheet.AllocatedRange;

                    var createExcelWindow = new CreateExcel();
                    var dataTable = sheet.ExportDataTable(localRange, true);                
                    createExcelWindow.grid.ItemsSource = dataTable.DefaultView;
                    createExcelWindow.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Выбранный файл не является Excel файлом", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CreateExcel(object sender, RoutedEventArgs e)
        {
            CreateExcel window = new CreateExcel();
            window.Show();
        }
    }
}
