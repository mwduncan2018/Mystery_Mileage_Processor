using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mystery_1051.Sandbox
{
    public static class Ext_String_01
    {
        public static float DoNotUseThis(this string mileageString)
        {
            float mileage = 0f;
            if (mileageString.EndsWith(" mi"))
            {
                mileageString = mileageString.Substring(mileageString.Length - 3);
                mileage = float.Parse(mileageString);
            }
            return mileage;
        }
    }
}
