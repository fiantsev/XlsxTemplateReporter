using System.IO;

namespace ExcelReportCreatorProject
{
    public interface IDocumentInjector
    {
        void Inject(Stream workbookStream);
    }
}