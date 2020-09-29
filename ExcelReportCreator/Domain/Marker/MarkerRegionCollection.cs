using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportCreatorProject.Domain.Markers
{
    public class MarkerRegionCollection : IEnumerable<MarkerRegion>
    {
        private readonly IEnumerable<Marker> _markers;

        public MarkerRegionCollection(IEnumerable<Marker> markers)
        {
            _markers = markers;
        }

        public IEnumerator<MarkerRegion> GetEnumerator()
        {
            //кэшируем
            var endMarkers = _markers.Where(x => x.MarkerType == MarkerType.End).ToList();
            var startMarkers = _markers.Where(x => x.MarkerType == MarkerType.Start);

            foreach (var startMarker in startMarkers)
            {
                var endMarker = endMarkers.FirstOrDefault(x => x.Id == startMarker.Id);
                var markerRegion = new MarkerRegion(startMarker, endMarker);
                yield return markerRegion;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}