namespace mail_reader
{
    public class EmailSettings
    {
        public string UserMail { get; set; }
        public string AppPassword { get; set; }
        public string ImapServer { get; set; }
        public int ImapPort { get; set; }
    }

    public class Paths
    {
        public string AttachmentPath { get; set; }
        public string FiltersFilePath { get; set; }
        public string ExcelOutputPath { get; set; }
    }

    public class TimeSettings
    {
        public int WaitTime { get; set; }
    }

    public class AppConfig
    {
        public EmailSettings EmailSettings { get; set; }
        public Paths Paths { get; set; }
        public TimeSettings TimeSettings { get; set; }
    }
}
