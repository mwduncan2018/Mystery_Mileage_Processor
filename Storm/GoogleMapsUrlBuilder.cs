using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051.Storm
{
    // Give this class an origin/destination/key
    // and this class will build you the URL
    // for the Google Maps Directions API.
    // This implementation of this class is a Builder Pattern.
    public class GoogleMapsUrlBuilder
    {
        private string _origin;
        private string _destination;
        private string _key;

        public GoogleMapsUrlBuilder WithOrigin(string origin)
        {
            _origin = origin.Replace(" ", "+");
            return this;
        }

        public GoogleMapsUrlBuilder WithDestination(string destination)
        {
            _destination = destination.Replace(" ", "+");
            return this;
        }

        public GoogleMapsUrlBuilder WithKey(string key)
        {
            _key = key;
            return this;
        }

        public string Build()
        {
            if (_origin == null ||
                _destination == null ||
                _key == null)
            {
                var message = "Make sure neither Origin, Destination, or Key are null when calling the Build() method.";
                throw new GoogleMapsUrlBuilderException(message);
            }

            var result =
                "https://maps.googleapis.com/maps/api/directions/json?origin=" +
                _origin +
                "&destination=" +
                _destination +
                "&key=" +
                _key;

            return result;
        }


    }
}
