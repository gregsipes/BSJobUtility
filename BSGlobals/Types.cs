using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSGlobals
{
   public class Types
    {
        public static object FormatNumber(string inputString)
        {
            if (inputString.Trim() == "" || inputString.Trim() == "?")
                return (object)DBNull.Value;
            else
            {
                inputString = inputString.Replace("$", "").Trim();

                if (inputString.EndsWith("-"))
                    return Decimal.Parse(inputString, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign);
                else
                    return inputString;

            }

        }

        public static object FormatDateTime(string inputString)
        {
            if (inputString.Trim() == "" || inputString.Trim() == "?")
                return (object)DBNull.Value;
            else
            {
                inputString = inputString.Trim();

                DateTime dateTime;

                if (DateTime.TryParse(inputString, out dateTime))
                    return dateTime.ToShortDateString();
                else
                    return (object)DBNull.Value;

            }
        }

        public static object FormatString(string inputString)
        {
            if (inputString.Trim() == "" || inputString.Trim() == "?")
                return (object)DBNull.Value;
            else
            {
                //some strings are parsed by position, so trimming whitespace is problematic
                //  inputString = inputString.Trim();

                return inputString;

            }
        }

    }
}
