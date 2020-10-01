using System;

namespace ExcelReportCreatorProject.Domain.Markers
{
    public class MarkerRegion
    {
        public Marker StartMarker { get; }
        public Marker EndMarker { get; }
        public bool Collapsed { get; }

        public MarkerRegion(Marker startMarker, Marker endMarker = null)
        {
            if(endMarker == null)
            {
                endMarker = startMarker.Clone();
                endMarker.MarkerType = MarkerType.End;
            }

            if (startMarker.MarkerType != MarkerType.Start || endMarker.MarkerType != MarkerType.End || startMarker.Id != endMarker.Id)
                throw new ArgumentException();

            StartMarker = startMarker;
            EndMarker = endMarker;

            Collapsed = startMarker.Position == endMarker.Position;
        }
    }
}