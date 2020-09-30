using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSGlobals
{
    public static class Extensions
    {
        public static string Right(this string value, int length)
        {
            string result = "";
            if (value.Length - length > 0)
            {
                result = value.Substring(value.Length - length);
            }
            return (result);
        }

        public static string Right(int startindex, string value)
        {
            int length = value.Length - startindex;
            return (Right(value, length));
        }

        public static string Left(this string value, int length)
        {
            int l = length;
            if (l > value.Length)
            {
                l = value.Length;
            }
            string result = value.Substring(0, l);
            return (result);
        }

        public static string Mid(this string value, int startindex, int length)
        {
            int l = length;
            if (startindex + l > value.Length)
            {
                l = value.Length - startindex;
            }
            string result = value.Substring(startindex, l);
            return (result);
        }

        public static string Deblank(string input)
        {
            // Fastest method that was tested (from Stack Overflow)
            int len = input.Length,
                index = 0,
                i = 0;
            var src = input.ToCharArray();
            bool skip = false;
            char ch;
            for (; i < len; i++)
            {
                ch = src[i];
                switch (ch)
                {
                    case '\u0020':
                    case '\u00A0':
                    case '\u1680':
                    case '\u2000':
                    case '\u2001':
                    case '\u2002':
                    case '\u2003':
                    case '\u2004':
                    case '\u2005':
                    case '\u2006':
                    case '\u2007':
                    case '\u2008':
                    case '\u2009':
                    case '\u200A':
                    case '\u202F':
                    case '\u205F':
                    case '\u3000':
                    case '\u2028':
                    case '\u2029':
                    case '\u0009':
                    case '\u000A':
                    case '\u000B':
                    case '\u000C':
                    case '\u000D':
                    case '\u0085':
                        if (skip) continue;
                        src[index++] = ch;
                        skip = true;
                        continue;
                    default:
                        skip = false;
                        src[index++] = ch;
                        continue;
                }
            }

            return new string(src, 0, index);
        }
    }
}
