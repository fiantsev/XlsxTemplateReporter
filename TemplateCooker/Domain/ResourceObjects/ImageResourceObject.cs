using System;

namespace TemplateCooker.Domain.ResourceObjects
{
    public class ImageResourceObject : ResourceObject
    {
        public byte[] Image { get; }

        public ImageResourceObject(byte[] image)
        {
            if (image == null)
                throw new NullReferenceException();

            Image = image;
        }
    }
}