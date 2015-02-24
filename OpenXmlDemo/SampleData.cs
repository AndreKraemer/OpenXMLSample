// --------------------------------------------------------------------------------------
// <copyright file="SampleData.cs" company="Andr� Kr�mer - Software, Training & Consulting">
//      Copyright (c) 2014 Andr� Kr�mer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;

namespace OpenXmlDemo
{
    internal class SampleData
    {
        /// <summary>
        /// Legt Demo Daten an
        /// </summary>
        /// <returns>Demo Daten</returns>
        public static Training GenerateSampleData()
        {
            var training = new Training
            {
                Title = "OpenXML SDK",
                From = DateTime.Today.AddDays(-2),
                To = DateTime.Today.AddDays(-1)
            };

            training.Contents.Add("�berblick Open XML SDK");
            training.Contents.Add("Lesen von Dokumenteigenschaften");
            training.Contents.Add("Erstellen von neuen Dokumenten");
            training.Contents.Add("Lesen von bestehenden Dokumenten");
            training.Contents.Add("Ver�ndern bestehender Dokumente");

            training.Attendees.Add(new Person { FirstName = "Wilhelm", LastName = "Brause", Title = "Herr" });
            training.Attendees.Add(new Person { FirstName = "Peter", LastName = "Schmitz", Title = "Herr" });
            training.Attendees.Add(new Person { FirstName = "Laura", LastName = "Buitoni", Title = "Frau" });


            return training;
        }
    }
}