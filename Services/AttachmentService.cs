using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using MimeKit;

namespace mail_reader.Services
{
    public class AttachmentService
    {
        private readonly string attachmentPath;
        private readonly string language;
        public List<string> ProcessedImageInfo { get; private set; } = new List<string>();

        public AttachmentService(string attachmentPath, string language)
        {
            this.attachmentPath = attachmentPath;
            this.language = language;
            if (!Directory.Exists(attachmentPath))
                Directory.CreateDirectory(attachmentPath);
        }

        public void ProcessAttachments(MimeMessage message)
        {
            try
            {
                ProcessedImageInfo.Clear();
                foreach (var attachment in message.Attachments)
                {
                    if (attachment is MimePart mimePart)
                    {
                        bool isImage = mimePart.ContentType.MediaType.ToLower() == "image" ||
                            (Path.GetExtension(mimePart.FileName)?.ToLower() is string ext &&
                                (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".gif"));

                        string fileName = mimePart.FileName;
                        string filePath = Path.Combine(attachmentPath, fileName);

                        if (isImage)
                        {
                            string newFileName = $"{Guid.NewGuid()}{Path.GetExtension(fileName)}";
                            filePath = Path.Combine(attachmentPath, newFileName);
                            ProcessedImageInfo.Add($"{newFileName} - {filePath}");
                        }

                        using (var stream = File.Create(filePath))
                        {
                            mimePart.Content.DecodeTo(stream);
                        }
                        Console.WriteLine(language == "ENG"
                            ? $"Attachment downloaded: {fileName} -> {filePath}"
                            : $"Ek indirildi: {fileName} -> {filePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(language == "ENG"
                    ? $"[ERROR] Attachment processing error: {ex.Message}"
                    : $"[HATA] Ek işleme hatası: {ex.Message}");
            }
        }
    }
}
