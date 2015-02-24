using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlDemo
{
    internal class AttendeeListSample2
    {
        /// <summary>
        /// Erstellt eine Teilnehmerliste auf basis der Vorlage Teilnehmerliste2.docx
        /// Die Tabelle ist bereits in der Vorlage definiert und wird nur fortgeführt
        /// </summary>
        public static void CreateAttendeeList2()
        {
            const string template = @".\Teilnehmerliste2.docx";
            Training training = SampleData.GenerateSampleData();

            string destinationFileName = template.Replace("2",
                string.Format(" {0} {1}", training.Title, training.From.ToString("yyyy-MM-dd")));
            File.Copy(template, destinationFileName, true);

            using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFileName, true))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

                // Erste Tabelle im Dokument suchen
                var table = doc.Body.Descendants<Table>().First();

                // Die letzte Zeile wird als Vorlage genutzt
                var templateRow = table.Elements<TableRow>().Last();

                foreach (var attendee in training.Attendees)
                {
                    // für jeden Teilnehmer die Vorlagenzeile clonen
                    var tableRow = templateRow.CloneNode(true) as TableRow;

                    SetCellContent(tableRow, 0, attendee.Title);
                    SetCellContent(tableRow, 1, attendee.FirstName);
                    SetCellContent(tableRow, 2, attendee.LastName);

                    table.Append(tableRow);
                }
                table.RemoveChild(templateRow);

                doc.Save();
            }

            Console.WriteLine("Datei {0} erzeugt", destinationFileName);
            Process.Start(destinationFileName);
        }

        private static void SetCellContent(TableRow tableRow, int pos, string text)
        {
            tableRow.Descendants<TableCell>().ElementAt(pos).RemoveAllChildren<Paragraph>();
            tableRow.Descendants<TableCell>().ElementAt(pos).Append(new Paragraph(new Run(new Text(text))));
        }
    }
}