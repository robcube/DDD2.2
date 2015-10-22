using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.OData.ProxyExtensions;
using Microsoft.Office365.OutlookServices;
using System.Speech.Recognition;
using System.IO;
using System.Threading.Tasks;

namespace Wavescribr.Models
{
    public class DisplayMessage
    {
        static bool completed;

        private DateTimeOffset? dateTimeReceived;
        private Recipient from;
        private IReadOnlyList<IAttachment> currentPage;
        public byte[] Attachment { get; set; }
        public string TranscribedText { get; set; }

        public string Attachments { get; set; }
        public string Subject { get; set; }
        public DateTimeOffset DateTimeReceived { get; set; }
        public string From { get; set; }

        public DisplayMessage(string subject, DateTimeOffset? dateTimeReceived,
            Recipient from)
        {
            this.Subject = subject;
            this.DateTimeReceived = (DateTimeOffset)dateTimeReceived;
            this.From = string.Format("{0} ({1})", from.EmailAddress.Name,
                            from.EmailAddress.Address);
        }

        //public DisplayMessage(string subject, DateTimeOffset? dateTimeReceived, 
        //    Recipient from, IList<IAttachment> attachments)
        //{
        //    Subject = subject;
        //    this.DateTimeReceived = (DateTimeOffset)dateTimeReceived;
        //    this.From = string.Format("{0} ({1})", from.EmailAddress.Name,
        //                    from.EmailAddress.Address);
        //    foreach (var attachment in attachments)
        //        Attachments += attachment.Name + " ";
        //}

        public DisplayMessage(string subject, DateTimeOffset? dateTimeReceived, Recipient from, IReadOnlyList<IAttachment> attachments)
        {
            Subject = subject;
            this.DateTimeReceived = (DateTimeOffset)dateTimeReceived;
            this.From = string.Format("{0} ({1})", from.EmailAddress.Name,
                            from.EmailAddress.Address);
            var a = attachments.Where(x => (x.ContentType == "application/octet-stream" || x.ContentType == "audio/x-wav") && !x.IsInline).FirstOrDefault();
            if (a != null)
            {
                this.Attachment = (a as FileAttachment).ContentBytes;
                this.TranscribedText = Task.Run(() => Transcribe(new MemoryStream(Attachment))).Result;
                Attachments = a.Name;
            }
            //{
            //    Attachments += attachment.Name + " ";
            //    //var bytes = ((FileAttachment)attachment).ContentBytes;
            //    //this.TranscibedText = Transcribe(new MemoryStream(bytes));
            //}
        }
    
        private string Transcribe(MemoryStream audioFile)
        {
            using (SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine())
            {
                // Create and load a grammar.
                Grammar dictation = new DictationGrammar();
                dictation.Name = "Dictation Grammar";

                 recognizer.LoadGrammar(dictation);

                // Configure the input to the recognizer.
                recognizer.SetInputToWaveStream(audioFile);

                // Attach event handlers for the results of recognition.
                recognizer.SpeechRecognized +=
                  new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
                recognizer.RecognizeCompleted +=
                  new EventHandler<RecognizeCompletedEventArgs>(recognizer_RecognizeCompleted);

                // Perform recognition on the entire file.
                Console.WriteLine("Starting asynchronous recognition...");
                completed = false;
                recognizer.RecognizeAsync(RecognizeMode.Single);

                // Keep the console window open.
                while (!completed)
                {
                    //Console.ReadLine();
                }
            }
            return TranscribedText;
        }

        // Handle the SpeechRecognized event.
        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result != null && e.Result.Text != null)
            {
                TranscribedText = string.Format("Recognized text =  {0}", e.Result.Text);
            }
            else
            {
                TranscribedText = "Recognized text not available.";
            }
        }

        // Handle the RecognizeCompleted event.
        void recognizer_RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                Console.WriteLine("  Error encountered, {0}: {1}",
                e.Error.GetType().Name, e.Error.Message);
            }
            if (e.Cancelled)
            {
                Console.WriteLine("  Operation cancelled.");
            }
            if (e.InputStreamEnded)
            {
                Console.WriteLine("  End of stream encountered.");
            }
            Console.WriteLine();
            Console.WriteLine("Done. Press ENTER.");
            completed = true;
        }
    }
}