using ExcelReportCreatorProject.Domain.ResourceObjects;
using System;

namespace ExcelReportCreatorProject.Service.ResourceObjectProvision
{
    public interface IResourceObjectProvider
    {
        Func<string, ResourceObject> Resolve { get; }
    }
}