// --------------------------------------------------------------------------------------
// <copyright file="program.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;

namespace OpenXmlDemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string input;

            do
            {
                DisplayMenu();
                input = Console.ReadLine();
                ProcessInput(input);
            } while (input != null && !input.Equals("7", StringComparison.InvariantCultureIgnoreCase));
        }


        private static void DisplayMenu()
        {
            Console.Clear();
            Console.WriteLine("Open XML SDK Demo");
            Console.WriteLine("=================");
            Console.WriteLine("");
            Console.WriteLine("Bitte wählen Sie einen der folgenden Punkte aus:");
            Console.WriteLine("[1]: Autoreninfo aus Datei auslesen");
            Console.WriteLine("[2]: Neues Dokument erzeugen");
            Console.WriteLine("[3]: Plain Text aus Dokument auslesen");
            Console.WriteLine("[4]: Teilnehmerliste aus Vorlage 1 erzeugen");
            Console.WriteLine("[5]: Teilnehmerliste aus Vorlage 2 erzeugen");
            Console.WriteLine("[6]: Zertifikat erstellen");
            Console.WriteLine("[7]: Programm beenden");
        }

        private static void ProcessInput(string input)
        {
            const string fileName = @".\Hallo OpenXML.docx";

            switch (input)
            {
                case "1":
                    DocumentInfoSample.DisplayAuthor(fileName);
                    break;
                case "2":
                    DocumentCreationSample.CreateDocument();
                    break;
                case "3":
                    Console.WriteLine(DocumentReaderSample.ReadTextFromDocument(fileName));
                    break;
                case "4":
                    AttendeeListSample1.CreateAttendeeList();
                    break;
                case "5":
                    AttendeeListSample2.CreateAttendeeList2();
                    break;
                case "6":
                    MailMergeSample.MailMerge();
                    break;
                case "7":
                    break;
                default:
                    Console.WriteLine("Bitte eine gültige Auswahl wählen");
                    break;
            }
            Console.WriteLine("Bitte Eingabetaste drücken");
            Console.ReadLine();
        }

   }
}