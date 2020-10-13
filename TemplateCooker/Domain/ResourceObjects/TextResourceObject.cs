using System;

namespace TemplateCooker.Domain.ResourceObjects
{
    public class TextResourceObject : ResourceObject
    {
        public string Text { get; }

        public TextResourceObject(string text)
        {
            if (text == null)
                throw new NullReferenceException();

            Text = text;
        }
    }
}