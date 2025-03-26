using System;

namespace mail_reader.Models
{
    public class EmailRecord
    {
        public string Subject { get; set; }
        public string Sender { get; set; }
        public DateTime Date { get; set; }
        public string Body { get; set; }
        public string ImageInfo { get; set; }
    }
}
