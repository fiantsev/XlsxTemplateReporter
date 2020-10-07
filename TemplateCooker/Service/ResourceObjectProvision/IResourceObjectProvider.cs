using TemplateCooker.Domain.ResourceObjects;
using System;

namespace TemplateCooker.Service.ResourceObjectProvision
{
    public interface IResourceObjectProvider
    {
        Func<string, ResourceObject> Resolve { get; }
    }
}