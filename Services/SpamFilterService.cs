using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailKit;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace mail_reader.Services
{
    public class SpamFilterService
    {
        private readonly string filtersFilePath;
        private readonly string language;
        public List<string> SpamKeywords { get; private set; } = new List<string>();
        public List<string> BlacklistEmails { get; private set; } = new List<string>();
        public List<string> WhitelistDomains { get; private set; } = new List<string>();

        public SpamFilterService(string filtersFilePath, string language)
        {
            this.filtersFilePath = filtersFilePath;
            this.language = language;
            LoadFiltersFromJson();
        }

        private void LoadFiltersFromJson()
        {
            try
            {
                if (File.Exists(filtersFilePath))
                {
                    string jsonContent = File.ReadAllText(filtersFilePath);
                    dynamic filters = JsonConvert.DeserializeObject(jsonContent);

                    SpamKeywords = filters.spam_keywords.ToObject<List<string>>();
                    BlacklistEmails = filters.blacklist_emails.ToObject<List<string>>();

                    // whitelist_domains dönüşümü
                    JArray domainsArray = filters.whitelist_domains as JArray;
                        WhitelistDomains = domainsArray
                            .Select(token => token.ToString().Trim().ToLower())
                            .ToList();

                    if (language == "ENG")
                        Console.WriteLine("Whitelist domains loaded: " + string.Join(", ", WhitelistDomains));
                    else
                        Console.WriteLine("Whitelist domainleri yüklendi: " + string.Join(", ", WhitelistDomains));
                }
                else
                {
                    Console.WriteLine(language == "ENG"
                        ? $"[ERROR] JSON file not found: {filtersFilePath}"
                        : $"[HATA] JSON dosyası bulunamadı: {filtersFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(language == "ENG"
                    ? $"[ERROR] Failed to load JSON: {ex.Message}"
                    : $"[HATA] JSON yükleme hatası: {ex.Message}");
            }
        }

        public bool IsWhitelisted(MimeMessage message)
        {
            var sender = message.From.Mailboxes.FirstOrDefault();
            if (sender == null)
                return false;

            string senderEmail = sender.Address.ToLower().Trim();
            string senderDomain = senderEmail.Split('@')[1].Trim();

            if (language == "ENG")
                Console.WriteLine($"Sender: {senderEmail}, Domain: {senderDomain}");
            else
                Console.WriteLine($"Gönderen: {senderEmail}, Domain: {senderDomain}");

            foreach (string trustedDomain in WhitelistDomains)
            {
                if (trustedDomain.Equals(senderDomain, StringComparison.OrdinalIgnoreCase))
                {
                    if (language == "ENG")
                        Console.WriteLine($"Trusted domain: {senderEmail}");
                    else
                        Console.WriteLine($"Güvenilir domain: {senderEmail}");
                    return true;
                }
            }
            return false;
        }

        public bool IsSpam(MimeMessage message)
        {
            string senderEmail = message.From.ToString().ToLower();
            string subject = message.Subject?.ToLower() ?? "";
            string body = message.TextBody?.ToLower() ?? "";

            string spfResult = message.Headers["Received-SPF"];
            Console.WriteLine(language == "ENG"
                ? $"SPF Result: {spfResult ?? "Header not found"}"
                : $"SPF Sonucu: {spfResult ?? "Başlık bulunamadı"}");
            if (!string.IsNullOrEmpty(spfResult) && spfResult.ToLower().Contains("fail"))
            {
                Console.WriteLine(language == "ENG"
                    ? $"SPF failed: {spfResult}"
                    : $"SPF başarısız: {spfResult}");
                return true;
            }

            string dkimResult = message.Headers["Authentication-Results"];
            Console.WriteLine(language == "ENG"
                ? $"DKIM Result: {dkimResult ?? "Header not found"}"
                : $"DKIM Sonucu: {dkimResult ?? "Başlık bulunamadı"}");
            if (!string.IsNullOrEmpty(dkimResult) && dkimResult.ToLower().Contains("dkim=fail"))
            {
                Console.WriteLine(language == "ENG"
                    ? $"DKIM failed: {dkimResult}"
                    : $"DKIM hatası: {dkimResult}");
                return true;
            }

            foreach (string blocked in BlacklistEmails)
            {
                if (senderEmail.Contains(blocked.ToLower()))
                {
                    Console.WriteLine(language == "ENG"
                        ? $"Blacklisted sender: {senderEmail}"
                        : $"Kara liste: {senderEmail}");
                    return true;
                }
            }

            foreach (string spamWord in SpamKeywords)
            {
                if (subject.Contains(spamWord) || body.Contains(spamWord))
                {
                    Console.WriteLine(language == "ENG"
                        ? $"Spam keyword detected: {spamWord} - {message.Subject}"
                        : $"Spam kelimesi bulundu: {spamWord} - {message.Subject}");
                    return true;
                }
            }
            return false;
        }
    }
}
