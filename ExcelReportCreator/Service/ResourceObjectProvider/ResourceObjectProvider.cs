using ExcelReportCreatorProject.Domain.ResourceObjects;
using System;

namespace ExcelReportCreatorProject.Service.ResourceObjectProvider
{
    //Ответственность: по Marker.Id вернуть типизированный ResourceObject
    public class ResourceObjectProvider : IResourceObjectProvider
    {
        public Func<string, ResourceObject> Resolve { get; }

        public ResourceObjectProvider(Func<string, ResourceObject> resolve)
        {
            Resolve = resolve;
        }

    }
}