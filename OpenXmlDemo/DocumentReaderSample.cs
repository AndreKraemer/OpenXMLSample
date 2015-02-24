// --------------------------------------------------------------------------------------
// <copyright file="DocumentReaderSample.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlDemo
{
    internal class DocumentReaderSample
    {
        /// <summary>
        /// Liest den kompletten Text einer Datei als Plain Text aus
        /// </summary>
        /// <param name="fileName">Name der Datie</param>
        /// <returns>Den Text</returns>
        public static string ReadTextFromDocument(string fileName)
        {
            string results = null;

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
                results = GetPlainText(docPart.Document.Body);

            }

            return results;
        }

        /// <summary>
        /// Extrahiert den Text aus einem XML Element
        /// </summary>
        /// <param name="rootElement">Das Element</param>
        /// <param name="sb">Optional: ein Stringbuilder</param>
        /// <returns>Den Text</returns>
        public static string GetPlainText(OpenXmlElement rootElement, StringBuilder sb = null)
        {
            if (sb == null)
            {
                sb = new StringBuilder();
            }

            foreach (var childElement in rootElement.Elements())
            {
                switch (childElement.LocalName)
                {
                    case "t": // Text
                        sb.Append(childElement.InnerText);
                        break;
                    case "tab": // Tab 
                        sb.Append("\t");
                        break;
                    case "cr": // Zeilenumbruch
                    case "br": // Seitenumbruch
                        sb.Append(Environment.NewLine);
                        break;
                    case "p":// Absatz 
                        GetPlainText(childElement, sb);
                        sb.AppendLine(Environment.NewLine);
                        break;

                    default:
                        GetPlainText(childElement, sb);
                        break;
                }
            }

            return sb.ToString();
        }
    }
}