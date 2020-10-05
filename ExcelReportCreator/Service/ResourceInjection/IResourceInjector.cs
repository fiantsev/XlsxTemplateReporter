using System;

namespace ExcelReportCreatorProject.Service.ResourceInjection
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; }
    }
}