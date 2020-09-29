using ExcelReportCreatorProject.Domain.ResourceObjects;
using System;

namespace ExcelReportCreatorProject.Service.ResourceObjectProvider
{
    public interface IResourceObjectProvider
    {
        Func<string, ResourceObject> Resolve { get; }
    }
}