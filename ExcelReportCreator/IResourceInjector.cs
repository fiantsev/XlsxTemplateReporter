using System;

namespace ExcelReportCreatorProject
{
    public interface IResourceInjector
    {
        Action<InjectionContext> Inject { get; set; }
    }
}