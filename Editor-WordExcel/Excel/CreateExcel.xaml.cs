using Microsoft.Win32;
using Spire.Xls;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
namespace Editor_WordExcel.Excel
{
    public partial class CreateExcel : Window
    {
        public string SelectedFilePath { get; set; }
        public CreateExcel()
        {
            InitializeComponent();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Выберите куда сохранить документ Excel";

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                var dataTable = grid.ItemsSource as DataView;
                Workbook wb = new Workbook();
                wb.Worksheets.Clear();

                Worksheet sheet = wb.Worksheets.Add("Лист 1");
                sheet.InsertDataView(dataTable, true, 1, 1);

                wb.SaveToFile(filePath, FileFormat.Version2016);
            }
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            DataTable dataTable;
            if (grid.ItemsSource != null)
            {
                dataTable = (grid.ItemsSource as DataView).Table;
            }
            else
            {
                dataTable = new DataTable();
                grid.ItemsSource = dataTable.DefaultView;
            }
            if (texts.Text != null && !string.IsNullOrWhiteSpace(texts.Text))
            {
                if (dataTable.Columns[texts.Text] == null)
                {
                    dataTable.Columns.Add(texts.Text);
                    grid.Columns.Add(new DataGridTextColumn { Header = texts.Text, Binding = new Binding(texts.Text) });
                    texts.Text = string.Empty;
                }
                else
                {
                    MessageBox.Show("Колонка с названием '" + texts.Text + "' уже существует", "Тфьу", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Введите название колонки!", "Тьфу", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            texts.Text = string.Empty;
        }
        private void Send_Click(object sender, RoutedEventArgs e)
        {
            SendOnline send = new SendOnline();
            send.Show();
        }
    }
}
