using System;

namespace ExcelReportCreatorProject
{
    public class ResourceInjector : IResourceInjector
    {
        public ResourceInjector(Action<IInjectionContext> inject)
        {
            Inject = inject;
        }

        public Action<IInjectionContext> Inject { get; set; }
    }
}