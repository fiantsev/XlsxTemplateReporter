using System;

namespace ExcelReportCreatorProject
{
    public interface IExcelReportCreator
    {
        void SetInjector(IResourceInjector injector);
        void Execute();
    }
}
