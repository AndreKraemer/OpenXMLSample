// --------------------------------------------------------------------------------------
// <copyright file="MailMergeSample.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 - 2015 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlDemo
{
    internal class MailMergeSample
    {
        /// <summary>
        /// Erstellt für jeden Teilnehmer ein Zertifikat auf Basis der Vorlage Zertifikat.docx
        /// </summary>
        public static void MailMerge()
        {
            const string template = @".\Zertifikat.docx";
            var training = SampleData.GenerateSampleData();

            foreach (var attendee in training.Attendees)
            {
                var destinationFileName = template.Replace(".docx",
                    string.Format(" {0} {1} {2} {3}.docx", training.Title, attendee.FirstName, attendee.LastName,
                        training.From.ToString("yyyy-MM-dd")));
                File.Copy(template, destinationFileName, true);

                using (var document = WordprocessingDocument.Open(destinationFileName, true))
                {
                    var body = document.MainDocumentPart.Document.Body;
                    ReplaceContentControl(body, "Anrede", attendee.Title);
                    ReplaceContentControl(body, "Vorname", attendee.FirstName);
                    ReplaceContentControl(body, "Nachname", attendee.LastName);
                    ReplaceContentControl(body, "Seminartitel", training.Title);
                    ReplaceContentControl(body, "Punkt1", training.Contents[0]);
                    ReplaceContentControl(body, "Punkt2", training.Contents[1]);
                    ReplaceContentControl(body, "Punkt3", training.Contents[2]);
                    ReplaceContentControl(body, "Punkt4", training.Contents[3]);
                    ReplaceContentControl(body, "Punkt5", training.Contents[4]);
                    ReplaceContentControl(body, "Datum", DateTime.Today.ToShortDateString());
                    ReplaceContentControl(body, "Von", training.From.ToShortDateString());
                    ReplaceContentControl(body, "Bis", training.To.ToShortDateString());

                    document.MainDocumentPart.Document.Save();
                    Console.WriteLine("Datei {0} im Programmordner erzeugt", destinationFileName);
                }
            }
        }

        private static void ReplaceContentControl(OpenXmlElement document, string tag, string content)
        {
            var contentControl =
                document.Descendants<SdtRun>().FirstOrDefault(cc => cc.Descendants<Tag>().Any(d => d.Val.HasValue && d.Val.Value == tag));
   
            var run = contentControl.SdtContentRun;
            var parent = contentControl.Parent;

            var replacement = (OpenXmlElement)run.ChildElements.FirstOrDefault().Clone();
            replacement.GetFirstChild<Text>().Text = content;
            
            parent.ReplaceChild(replacement, contentControl);
        }
    }
}