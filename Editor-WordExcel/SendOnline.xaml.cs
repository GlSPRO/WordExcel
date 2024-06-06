using Editor_WordExcel.Word;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Windows.Documents;

namespace Editor_WordExcel
{
    public partial class SendOnline : Window
    {
        public SendOnline()
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

            var createWordWindow = new CreateWord();
            TextRange range = new TextRange(createWordWindow.rtb.Document.ContentStart, createWordWindow.rtb.Document.ContentEnd);

            MailMessage message = new MailMessage(Login.Text, ToWhom.Text, Subject.Text, range.Text);
            message.IsBodyHtml = true;

            string smtpServer;
            bool enableSsl = true;
            int port = 465;

            switch (GetEmailProvider(Login.Text))
            {
                case EmailProvider.Gmail:
                    smtpServer = "smtp.gmail.com";
                    break;
                case EmailProvider.Mail:
                    smtpServer = "smtp.mail.ru";
                    break;
                case EmailProvider.Rambler:
                    smtpServer = "smtp.rambler.ru";
                    break;
                case EmailProvider.Yandex:
                    smtpServer = "smtp.yandex.ru";
                    break;
                default:
                    MessageBox.Show("Неизвестный провайдер электронной почты. Невозможно отправить электронное письмо.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;            
            }

            SmtpClient smtpClient = new SmtpClient(smtpServer, port);
            smtpClient.EnableSsl = enableSsl;
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential(Login.Text, Password.Text);
            //smtpClient.Timeout = 100;

            try
            {
                smtpClient.Send(message);
                MessageBox.Show("Сообщение успешно отправлено.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (SmtpException ex)
            {
                if (ex.Message.Contains("аутентификация"))
                {
                    MessageBox.Show("Неверный логин или пароль.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (ex.Message.Contains("получатель"))
                {
                    MessageBox.Show("Неверный адрес получателя.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show("При отправке сообщения произошла ошибка. Пожалуйста, попробуйте еще раз позже.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                case "mail.ru":
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
