using Editor_WordExcel.Excel;
using Spire.Xls;
using System;
using System.Net.Mail;
using System.Net;
using System.Windows;

namespace Editor_WordExcel
{
    public partial class SendExcel : Window
    {
        public SendExcel()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Login.Text) || string.IsNullOrEmpty(Password.Text) || string.IsNullOrEmpty(ToWhom.Text) || string.IsNullOrEmpty(Subject.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля перед отправкой сообщения.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!IsValidEmail(Login.Text))
            {
                MessageBox.Show("Неверный формат адреса электронной почты для входа.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!IsValidEmail(ToWhom.Text))
            {
                MessageBox.Show("Неверный формат адреса электронной почты для получателя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            CellRange localRange = sheet.AllocatedRange;
            var createExcelWindow = new CreateExcel();
            var dataTable = sheet.ExportDataTable(localRange, true);
            createExcelWindow.grid.ItemsSource = dataTable.DefaultView;
            createExcelWindow.ShowDialog();
            string excelFilePath = createExcelWindow.SelectedFilePath;
            if (!string.IsNullOrEmpty(excelFilePath))
            {
                MailMessage message = new MailMessage();
                message.From = new MailAddress(Login.Text);
                message.To.Add(ToWhom.Text);
                message.Subject = Subject.Text;
                message.Attachments.Add(new Attachment(excelFilePath));
                EmailProvider emailProvider = GetEmailProvider(Login.Text);
                string smtpServer;
                bool enableSsl;
                switch (emailProvider)
                {
                    case EmailProvider.Gmail:
                        smtpServer = "smtp.gmail.com";
                        enableSsl = true;
                        break;
                    case EmailProvider.Mail:
                        smtpServer = "smtp.mail.ru";
                        enableSsl = true;
                        break;
                    case EmailProvider.Rambler:
                        smtpServer = "smtp.rambler.ru";
                        enableSsl = false;
                        break;
                    case EmailProvider.Yandex:
                        smtpServer = "smtp.yandex.ru";
                        enableSsl = true;
                        break;
                    default:
                        MessageBox.Show("Неизвестный провайдер электронной почты. Невозможно отправить письмо.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                }
                SmtpClient smtpClient = new SmtpClient(smtpServer);
                smtpClient.EnableSsl = enableSsl;
                smtpClient.Credentials = new NetworkCredential(Login.Text, Password.Text);
                try
                {
                    smtpClient.Send(message);
                    MessageBox.Show("Письмо отправлено успешно.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (SmtpException ex)
                {
                    if (ex.Message.Contains("аутентификация"))
                    {
                        MessageBox.Show("Неверный логин или пароль.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else if (ex.Message.Contains("получатель"))
                    {
                        MessageBox.Show("Неверный адрес электронной почты получателя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        MessageBox.Show("Ошибка при отправке письма. Пожалуйста, попробуйте еще раз позже.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Произошла ошибка при отправке письма. Пожалуйста, попробуйте еще раз позже.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var mailAddress = new MailAddress(email);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public enum EmailProvider
        {
            Gmail,
            Mail,
            Rambler,
            Yandex
        }

        public EmailProvider GetEmailProvider(string email)
        {
            string domain = email.Substring(email.IndexOf('@') + 1);

            switch (domain)
            {
                case "gmail.com":
                    return EmailProvider.Gmail;
                case "mail.com":
                    return EmailProvider.Mail;
                case "rambler.ru":
                    return EmailProvider.Rambler;
                case "yandex.ru":
                    return EmailProvider.Yandex;
                default:
                    return EmailProvider.Mail;
            }
        }
    }
}
