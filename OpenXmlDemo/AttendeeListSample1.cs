// --------------------------------------------------------------------------------------
// <copyright file="AttendeeListSample1.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlDemo
{
    internal class AttendeeListSample1
    {
        /// <summary>
        /// Erzeugt eine Teilnehmerliste auf Basis der Vorlage Teilnerhmerliste1.docx
        /// Die Tabelle wird komplett im Code erzeugt
        /// </summary>
        public static void CreateAttendeeList()
        {
            const string template = @".\Teilnehmerliste1.docx";
            var training = SampleData.GenerateSampleData();

            string destinationFileName = template.Replace("1",
                string.Format("{0} {1}", training.Title, training.From.ToString("yyyy-MM-dd")));
            File.Copy(template, destinationFileName, true);

            using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFileName, true))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

                var table = new Table();

                var borders = new TableBorders
                {
                    TopBorder = new TopBorder { Val = BorderValues.Single, Size = 6 },
                    LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 6 },
                    BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 6 },
                    RightBorder = new RightBorder { Val = BorderValues.Single, Size = 6 },
                    InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Dashed, Size = 6 }
                };

                var tableProperties = new TableProperties();
                tableProperties.Append(borders);
                table.Append(tableProperties);

                var tableRow = new TableRow();

                var tableCell = CreateTableHeader("Titel");
                tableRow.Append(tableCell);

                tableCell = CreateTableHeader("Vorname");
                tableRow.Append(tableCell);

                tableCell = CreateTableHeader("Nachname");
                tableRow.Append(tableCell);


                tableCell = CreateTableHeader("Unterschrift");
                tableRow.Append(tableCell);
                table.Append(tableRow);

                foreach (Person attendee in training.Attendees)
                {
                    tableRow = new TableRow();

                    tableRow.Append(CreateTableCell(attendee.Title));
                    tableRow.Append(CreateTableCell(attendee.FirstName));
                    tableRow.Append(CreateTableCell(attendee.LastName));
                    tableRow.Append(CreateTableCell(""));
                    table.Append(tableRow);
                }

                doc.Body.Append(table);
                doc.Save();
            }

            Console.WriteLine("Datei {0} erzeugt", destinationFileName);
            Process.Start(destinationFileName);
        }

        /// <summary>
        ///  Erzeugt eine neue Überschriftentabellenzelle
        /// </summary>
        /// <param name="content">Spaltenüberschrift</param>
        /// <returns>Die Tabellenzelle</returns>
        private static TableCell CreateTableHeader(string content)
        {
            var text = new Text(content);
            var runProperties = new RunProperties();
            runProperties.Append(new Bold());
            var run = new Run();
            run.Append(runProperties);
            run.Append(text);
            var paragraph = new Paragraph(run);
            var tableCell = new TableCell(paragraph);

            var tcp = new TableCellProperties();
            var tcw = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6000" };
            var shading1 = new Shading()
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = "9CC2E5",
                ThemeFill = ThemeColorValues.Accent1,
                ThemeFillTint = "99"
            };
            tcp.Append(tcw);
            tcp.Append(shading1);
            tableCell.Append(tcp);
            return tableCell;
        }

        /// <summary>
        /// Erzeugt eine neue Tabellenzelle
        /// </summary>
        /// <param name="content">Der Inhalt der Zelle</param>
        /// <returns>Die Zelle</returns>
        private static TableCell CreateTableCell(string content)
        {
            var text = new Text(content);
            var run = new Run(text);
            var paragraph = new Paragraph(run);
            var tableCell = new TableCell(paragraph);

            var tcp = new TableCellProperties();
            var tcw = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "6000" };
            tcp.Append(tcw);
            tableCell.Append(tcp);
            return tableCell;
        }
    }
}