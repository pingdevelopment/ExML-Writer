using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Globalization;

namespace PingDevelopment.ExcelML.Configuration
{
    public class DateTimeDataElement : BaseFormatElement
    {
        /// <summary>
        /// The format string for the numeric data
        /// </summary>
        //[StringValidator(InvalidCharacters = "[^0-9]|[^FN]")]
        [ConfigurationProperty("formatString", DefaultValue = @"[$-409]m/d/yy\ h:mm\ AM/PM;@", IsRequired = false)]
        public String FormatString
        {
            get
            {
                if (this["formatString"] == null) return CultureInfo.CurrentCulture.DateTimeFormat.FullDateTimePattern;
                return (String)this["formatString"];
            }
            set { this["formatString"] = value; }
        }
    }
}
