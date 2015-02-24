// --------------------------------------------------------------------------------------
// <copyright file="AttendeeListSample2.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlDemo
{
    internal class DocumentCreationSample
    {
        /// <summary>
        /// Erzeugt ein leeres Dokument mit dem Inhalt "Hallo Welt"
        /// </summary>
        public static void CreateDocument()
        {
            string fileName = Path.Combine(@".\", Path.GetRandomFileName() + ".docx");

            using (
                WordprocessingDocument document = WordprocessingDocument.Create(fileName,
                    WordprocessingDocumentType.Document))
            {
                var text = new Text("Hallo Welt");
                var run = new Run(text);
                var paragraph = new Paragraph(run);
                var body = new Body(paragraph);
                var doc = new Document(body);

                document.AddMainDocumentPart();

                document.MainDocumentPart.Document = doc;

                document.MainDocumentPart.Document.Save();

            }

            Console.WriteLine("Datei {0} erzeugt", fileName);
            Process.Start(fileName);
        }
    }
}