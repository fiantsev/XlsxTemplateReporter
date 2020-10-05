﻿using System;

namespace ExcelReportCreatorProject.Domain.Markers
{
    public class MarkerOptions
    {
        public string Prefix { get; }
        public string Suffix { get; }
        public string Terminator { get; }

        public MarkerOptions(string prefix, string suffix, string terminator)
        {
            Prefix = prefix;
            Suffix = suffix;
            Terminator = terminator;
        }

        public static implicit operator MarkerOptions(string markerTemplate)
        {
            var parts = markerTemplate.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return new MarkerOptions(parts[0], parts[1], parts[2]);
        }
    }
}