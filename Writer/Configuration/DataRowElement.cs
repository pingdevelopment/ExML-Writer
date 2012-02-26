using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace PingDevelopment.ExcelML.Configuration
{
    public class DataRowElement : ConfigurationElement
    {
        /// <summary>
        /// Format configuration for numeric data
        /// </summary>
        [ConfigurationProperty("number")]
        public NumericDataElement NumericFormat
        {
            get { return (NumericDataElement)this["number"]; }
            set { this["number"] = value; }
        }

        /// <summary>
        /// Format configuration for string data
        /// </summary>
        [ConfigurationProperty("string")]
        public StringDataElement StringFormat
        {
            get { return (StringDataElement)this["string"]; }
            set { this["string"] = value; }
        }

        /// <summary>
        /// Format configuration for date/time data
        /// </summary>
        [ConfigurationProperty("date")]
        public DateTimeDataElement DateFormat
        {
            get { return (DateTimeDataElement)this["date"]; }
            set { this["date"] = value; }
        }
    }
}
