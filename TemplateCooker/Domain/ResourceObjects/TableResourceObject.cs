using System.Collections.Generic;

namespace TemplateCooker.Domain.ResourceObjects
{
    public class TableResourceObject : ResourceObject
    {
        public List<List<object>> Table { get; set; }
    }
}