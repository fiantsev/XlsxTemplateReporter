using System;

namespace ExcelReportCreatorProject.Service.Injection
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; set; }
    }
}