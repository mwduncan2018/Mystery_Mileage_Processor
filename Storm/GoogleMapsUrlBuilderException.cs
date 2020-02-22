using System;
using System.Runtime.Serialization;

namespace Mystery_1051.Storm
{
    [Serializable]
    public class GoogleMapsUrlBuilderException : Exception
    {
        public GoogleMapsUrlBuilderException()
        {
        }

        public GoogleMapsUrlBuilderException(string message) : base(message)
        {
        }

        public GoogleMapsUrlBuilderException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected GoogleMapsUrlBuilderException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}