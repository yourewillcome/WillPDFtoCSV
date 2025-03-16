using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text.RegularExpressions;

namespace PdfToCsvConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("PDF to CSV Converter for Outlook Emails");
            Console.WriteLine("=======================================");

            try
            {
                // Ask for input PDF path
                Console.Write("Enter path to PDF file: ");
                string pdfPath = Console.ReadLine().Trim('"');

                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine("Error: File does not exist.");
                    return;
                }

                // Ask for output CSV path
                Console.Write("Enter path for output CSV file: ");
                string csvPath = Console.ReadLine().Trim('"');

                // Process the PDF and create CSV
                ConvertPdfToCsv(pdfPath, csvPath);

                Console.WriteLine($"Successfully converted to {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static void ConvertPdfToCsv(string pdfPath, string csvPath)
        {
            // Email pattern matching - customize based on your PDF structure
            // This pattern looks for common email elements
            var emailDataList = new List<EmailData>();
            
            // Extract text from PDF
            string pdfText = ExtractTextFromPdf(pdfPath);
            
            // Split the text into lines
            string[] lines = pdfText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            
            // Process the lines to extract email data
            EmailData currentEmail = null;
            
            foreach (var line in lines)
            {
                // Check if line contains a date (potential email header)
                if (Regex.IsMatch(line, @"\b\d{1,2}/\d{1,2}/\d{4}\b") || 
                    Regex.IsMatch(line, @"\b[A-Z][a-z]{2}\s\d{1,2},\s\d{4}\b"))
                {
                    // Save previous email if exists
                    if (currentEmail != null && !string.IsNullOrWhiteSpace(currentEmail.Subject))
                    {
                        emailDataList.Add(currentEmail);
                    }
                    
                    // Start a new email
                    currentEmail = new EmailData();
                    
                    // Try to extract date
                    currentEmail.Date = line.Trim();
                }
                else if (currentEmail != null)
                {
                    // Check if line contains "From:" or "To:"
                    if (line.StartsWith("From:"))
                    {
                        currentEmail.From = line.Substring(5).Trim();
                    }
                    else if (line.StartsWith("To:"))
                    {
                        currentEmail.To = line.Substring(3).Trim();
                    }
                    else if (line.StartsWith("Subject:"))
                    {
                        currentEmail.Subject = line.Substring(8).Trim();
                    }
                    else if (!string.IsNullOrWhiteSpace(line) && 
                             string.IsNullOrWhiteSpace(currentEmail.Subject) &&
                             !string.IsNullOrWhiteSpace(currentEmail.From))
                    {
                        // If we have a From but no Subject yet, this line might be the subject
                        currentEmail.Subject = line.Trim();
                    }
                    else if (!string.IsNullOrWhiteSpace(currentEmail.Subject))
                    {
                        // Append to body if we already have a subject
                        currentEmail.Body += line + Environment.NewLine;
                    }
                }
            }
            
            // Add the last email
            if (currentEmail != null && !string.IsNullOrWhiteSpace(currentEmail.Subject))
            {
                emailDataList.Add(currentEmail);
            }
            
            // Write to CSV
            using (var writer = new StreamWriter(csvPath, false, Encoding.UTF8))
            {
                // Write header
                writer.WriteLine("Date,From,To,Subject,Body");
                
                // Write data
                foreach (var email in emailDataList)
                {
                    writer.WriteLine($"\"{EscapeCsvField(email.Date)}\",\"{EscapeCsvField(email.From)}\",\"{EscapeCsvField(email.To)}\",\"{EscapeCsvField(email.Subject)}\",\"{EscapeCsvField(email.Body)}\"");
                }
            }
        }
        
        static string ExtractTextFromPdf(string pdfPath)
        {
            StringBuilder text = new StringBuilder();
            
            using (PdfReader reader = new PdfReader(pdfPath))
            {
                using (PdfDocument document = new PdfDocument(reader))
                {
                    for (int i = 1; i <= document.GetNumberOfPages(); i++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string pageText = PdfTextExtractor.GetTextFromPage(document.GetPage(i), strategy);
                        text.Append(pageText);
                    }
                }
            }
            
            return text.ToString();
        }
        
        static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
                
            // Replace double quotes with two double quotes
            return field.Replace("\"", "\"\"");
        }
    }
    
    class EmailData
    {
        public string Date { get; set; } = "";
        public string From { get; set; } = "";
        public string To { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
    }
} 