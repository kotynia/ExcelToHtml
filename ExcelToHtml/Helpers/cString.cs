using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml.Helpers
{
        public static class cString
        {
            #region String extensions

            public static String SubstringAfter(this System.String value, System.String afterString)
            {
                Int32 position = value.LastIndexOf(afterString);

                if (position == -1)
                    return string.Empty;

                Int32 adjustedPosition = position + afterString.Length;

                if (adjustedPosition >= value.Length)
                    return string.Empty;

                return value.Substring(adjustedPosition);
            }

            public static String SubstringBefore(this System.String value, System.String beforeSubstring)
            {
                Int32 position = value.IndexOf(beforeSubstring);

                if (position == -1)
                    return string.Empty;

                return value.Substring(0, position);
            }

            #endregion String extensions
        }
    

}
