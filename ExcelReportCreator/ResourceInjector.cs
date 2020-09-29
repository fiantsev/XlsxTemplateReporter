﻿using System;

namespace ExcelReportCreatorProject
{
    public class ResourceInjector : IResourceInjector
    {
        public ResourceInjector(Action<InjectionContext> inject)
        {
            Inject = inject;
        }

        public Action<InjectionContext> Inject { get; set; }
    }
}