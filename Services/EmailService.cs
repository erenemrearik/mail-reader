using System;
using System.Collections.Generic;
using System.Threading;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;
using mail_reader.Models;
using mail_reader;

namespace mail_reader.Services
{
    public class EmailService
    {
        private readonly AppConfig config;
        private readonly SpamFilterService spamFilterService;
        private readonly AttachmentService attachmentService;
        private readonly ExcelService excelService;
        private readonly bool generateExcelReport;
        private readonly string language;
        private List<EmailRecord> emailRecords = new List<EmailRecord>();

        public EmailService(AppConfig config, bool generateExcelReport, string language)
        {
            this.config = config;
            this.generateExcelReport = generateExcelReport;
            this.language = language;
            spamFilterService = new SpamFilterService(config.Paths.FiltersFilePath, language);
            attachmentService = new AttachmentService(config.Paths.AttachmentPath, language);
            excelService = new ExcelService(config.Paths.ExcelOutputPath);
        }

        public void ProcessEmails()
        {
            using (var client = new ImapClient())
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine(language == "ENG"
                        ? "Starting IMAP connection..."
                        : "IMAP bağlantısı başlatılıyor...");

                    client.Connect(config.EmailSettings.ImapServer, config.EmailSettings.ImapPort, true);
                    Console.WriteLine(language == "ENG"
                        ? "Connected to IMAP server."
                        : "IMAP sunucusuna bağlandı.");

                    client.Authenticate(config.EmailSettings.UserMail, config.EmailSettings.AppPassword);
                    if (!client.IsAuthenticated)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(language == "ENG"
                            ? "Authentication failed!"
                            : "Kimlik doğrulama başarısız!");
                        return;
                    }
                    Console.ResetColor();

                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);

                    int waitTime = config.TimeSettings.WaitTime;
                    int addMinutes = waitTime / 60000;

                    while (true)
                    {
                        DateTime today = DateTime.Today;
                        DateTime startDate = today;
                        DateTime endDate = today.AddDays(1).AddSeconds(-1);
                        string currentTime = language == "ENG"
                            ? $"({DateTime.Now:HH:mm} - {DateTime.Now.AddMinutes(addMinutes):HH:mm})"
                            : $"({DateTime.Now:HH:mm} - {DateTime.Now.AddMinutes(addMinutes):HH:mm})";

                        inbox.Close();
                        inbox.Open(FolderAccess.ReadWrite);

                        var unreadMessages = inbox.Search(MailKit.Search.SearchQuery.NotSeen);

                        if (unreadMessages.Count > 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine(language == "ENG"
                                ? $"{unreadMessages.Count} new emails found!"
                                : $"{unreadMessages.Count} yeni e-posta bulundu!");
                            Console.ResetColor();

                            foreach (var uid in unreadMessages)
                            {
                                var message = inbox.GetMessage(uid);
                                DateTime emailDate = message.Date.DateTime;
                                if (emailDate >= startDate && emailDate <= endDate)
                                {
                                    if (!spamFilterService.IsWhitelisted(message))
                                    {
                                        Console.WriteLine(language == "ENG"
                                            ? $"Email not in whitelist, skipping: {message.Subject}"
                                            : $"Whitelist dışı e-posta işlenmeyecek: {message.Subject}");

                                        if (spamFilterService.IsSpam(message))
                                        {
                                            Console.WriteLine(language == "ENG"
                                                ? $"Spam/blacklisted detected, deleting email: {message.Subject}"
                                                : $"Blacklist veya spam tespit edildi, e-posta siliniyor: {message.Subject}");
                                            DeleteEmail(inbox, uid);
                                        }
                                        else
                                        {
                                            inbox.AddFlags(uid, MessageFlags.Seen, true);
                                        }
                                        continue;
                                    }

                                    Console.WriteLine("====================================");
                                    Console.WriteLine(language == "ENG"
                                        ? $"Subject: {message.Subject}"
                                        : $"Konu: {message.Subject}");
                                    Console.WriteLine(language == "ENG"
                                        ? $"Sender: {message.From}"
                                        : $"Gönderen: {message.From}");
                                    Console.WriteLine(language == "ENG"
                                        ? $"Date: {message.Date}"
                                        : $"Tarih: {message.Date}");
                                    Console.WriteLine(language == "ENG"
                                        ? "Processing content..."
                                        : "İçerik işleniyor...");

                                    attachmentService.ProcessAttachments(message);

                                    string imageInfo = attachmentService.ProcessedImageInfo.Any()
                                        ? string.Join(" | ", attachmentService.ProcessedImageInfo)
                                        : (language == "TR" ? "Yok" : "None");

                                    emailRecords.Add(new EmailRecord
                                    {
                                        Subject = message.Subject,
                                        Sender = message.From.ToString(),
                                        Date = message.Date.DateTime,
                                        Body = message.TextBody,
                                        ImageInfo = imageInfo
                                    });

                                    inbox.AddFlags(uid, MessageFlags.Seen, true);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine(language == "ENG"
                                ? $"No new emails found. Waiting {addMinutes} minutes... {currentTime}"
                                : $"Yeni e-posta bulunamadı. {addMinutes} dakika bekleniyor... {currentTime}");
                            Thread.Sleep(waitTime);
                        }

                        if (generateExcelReport && emailRecords.Any())
                        {
                            excelService.AppendToDailyExcelReport(emailRecords, language);
                            emailRecords.Clear();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(language == "ENG"
                        ? $"IMAP connection error: {ex.Message}"
                        : $"IMAP bağlantı hatası: {ex.Message}");
                    Console.ResetColor();
                }
                finally
                {
                    client.Disconnect(true);
                    Console.WriteLine(language == "ENG"
                        ? "IMAP connection closed."
                        : "IMAP bağlantısı kapatıldı.");
                }
            }
        }

        private void DeleteEmail(IMailFolder inbox, UniqueId uid)
        {
            try
            {
                inbox.AddFlags(uid, MessageFlags.Deleted, true);
                Console.WriteLine(language == "ENG"
                    ? "Email successfully deleted."
                    : "E-posta başarıyla silindi.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(language == "ENG"
                    ? $"[ERROR] Email deletion error: {ex.Message}"
                    : $"[HATA] E-posta silme hatası: {ex.Message}");
            }
        }
    }
}
