using TemplateCooker.Domain.Markers;
using TemplateCooker.Service.ResourceInjection;
using TemplateCooker.Service.ResourceObjectProvision;

namespace TemplateCooker.Service.Creation
{
    public class DocumentInjectorOptions
    {
        public IResourceInjector ResourceInjector { get; set; }
        public IResourceObjectProvider ResourceObjectProvider { get; set; }
        public MarkerOptions MarkerOptions { get; set; }
    }
}