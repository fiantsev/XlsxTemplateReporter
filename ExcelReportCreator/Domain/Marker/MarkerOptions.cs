namespace ExcelReportCreatorProject.Domain
{
    public class MarkerOptions
    {
        public string Prefix { get; }
        public string Suffix { get; }

        public MarkerOptions(string prefix, string suffix)
        {
            Prefix = prefix;
            Suffix = suffix;
        }

        public static implicit operator MarkerOptions(string markerTemplate)
        {
            var trimmed = markerTemplate.Trim();
            var prefix = trimmed.Substring(0, trimmed.Length / 2);
            var suffix = trimmed.Substring(trimmed.Length / 2);
            return new MarkerOptions(prefix, suffix);
        }
    }
}