using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051.Storm
{
    public static class ExtensionMethods
    {
        public static float ConvertMileageStringToFloat(this string mileageString)
        {
            float mileage = 0f;
            if (mileageString.EndsWith(" mi"))
            {
                mileageString = mileageString.Substring(0, mileageString.Length - 3);
                mileage = float.Parse(mileageString);
            }
            return mileage;
        }
    }

}