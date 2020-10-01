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
            var trimmed = markerTemplate.Trim();
            var prefix = trimmed.Substring(0, trimmed.Length / 2);
            var terminator = trimmed.Substring(trimmed.Length / 2, 1);
            var suffix = trimmed.Substring(trimmed.Length / 2 + 1);
            return new MarkerOptions(prefix, suffix, terminator);
        }
    }
}