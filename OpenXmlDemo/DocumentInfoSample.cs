// --------------------------------------------------------------------------------------
// <copyright file="DocumentInfoSample.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlDemo
{
    internal class DocumentInfoSample
    {
        /// <summary>
        /// Liest die Autoreninformationen aus einer Word Datei aus
        /// </summary>
        /// <param name="fileName">Dateiname einer docx Datei</param>
        public static void DisplayAuthor(string fileName)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
            {
                string author = document.PackageProperties.Creator;
                Console.WriteLine(author);
            }
        }
    }
}