// --------------------------------------------------------------------------------------
// <copyright file="Training.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;


namespace OpenXmlDemo
{
    public class Training
    {

        public Training()
        {
            Contents = new List<string>();
            Attendees = new List<Person>();
        }
        public string Title { get; set; }
        public DateTime From { get; set; }

        public DateTime To { get; set; }

        public List<string> Contents { get; set; }

        public List<Person> Attendees { get; set; }
    }
}