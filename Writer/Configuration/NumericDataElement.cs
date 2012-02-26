using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace PingDevelopment.ExcelML.Configuration
{
    public class NumericDataElement : BaseFormatElement
    {
        /// <summary>
        /// The format string for the numeric data
        /// </summary>
        //[StringValidator(InvalidCharacters = "[^0-9]|[^FN]")]
        [ConfigurationProperty("formatString", DefaultValue = "General", IsRequired = false)]
        public String FormatString
        {
            get { return (String)this["formatString"]; }
            set { this["formatString"] = value; }
        }
    }
}
