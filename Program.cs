using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using mail_reader.Services;
using mail_reader;

namespace mail_reader
{
    class Program
    {
        static void Main(string[] args)
        {
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            AppConfig config = configuration.Get<AppConfig>();

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("====================================");
            Console.WriteLine("       IMAP Mail Reader v1.0        ");
            Console.WriteLine("====================================");
            Console.ResetColor();

            Console.Write("Lütfen dil seçiniz (TR/ENG): ");
            string langInput = Console.ReadLine().Trim().ToUpper();
            string language = (langInput == "ENG") ? "ENG" : "TR";

            Console.Write(language == "TR"
                ? "Excel çıktısı almak istiyor musunuz? (E/H): "
                : "Do you want to generate Excel output? (Y/N): ");
            string input = Console.ReadLine().Trim().ToUpper();
            bool generateExcel = (language == "ENG" ? input == "Y" : input == "E");

            EmailService emailService = new EmailService(config, generateExcel, language);
            emailService.ProcessEmails();
        }
    }
}
