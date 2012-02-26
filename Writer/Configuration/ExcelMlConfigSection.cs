using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace PingDevelopment.ExcelML.Configuration
{
    public class ExcelMlConfigSection : ConfigurationSection
    {
        /// <summary>
        /// Default font style for all cells
        /// </summary>
        [ConfigurationProperty("default")]
        public DefaultFormatElement DefaultFormat
        {
            get { return (DefaultFormatElement)this["default"]; }
            set { this["default"] = value; }
        }

        [ConfigurationProperty("header")]
        public HeaderElement HeaderFormat
        {
            get
            {
                return (HeaderElement)this["header"];
            }
            set
            { this["header"] = value; }
        }

        [ConfigurationProperty("data")]
        public DataRowElement DataFormat
        {
            get
            {
                return (DataRowElement)this["data"];
            }
            set
            { this["data"] = value; }
        }
    }
}
