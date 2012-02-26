using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace PingDevelopment.ExcelML.Configuration
{
    public class BaseFormatElement : ConfigurationElement
    {
        /// <summary>
        /// Flag indicating if text in the column header element should be bold
        /// </summary>
        [ConfigurationProperty("bold", DefaultValue = -1, IsRequired = false)]
        [IntegerValidator(MinValue = -1, MaxValue = 1)]
        public int Bold
        {
            get
            {
                return (Int32)this["bold"];
            }
            set
            {
                this["bold"] = value;
            }
        }

        /// <summary>
        /// Valid values are:
        /// 1. Bottom
        /// 2. Middle
        /// 3. Top
        /// </summary>
        [ConfigurationProperty("verticalAlignment", DefaultValue = "Bottom", IsRequired = false)]
        [StringValidator(InvalidCharacters = "[^A-Z]|[^a-z]")]
        public String VerticalAlignment
        {
            get
            {
                return (String)this["verticalAlignment"];
            }
            set
            {
                this["verticalAlignment"] = value;
            }
        }

        /// <summary>
        /// Valid values are:
        /// 1. Left
        /// 2. Center
        /// 3. Right
        /// </summary>
        [ConfigurationProperty("horizontalAlignment", DefaultValue = "", IsRequired = false)]
        [StringValidator(InvalidCharacters = "[^A-Z]|[^a-z]")]
        public String HorizontalAlignment
        {
            get
            {
                return (String)this["horizontalAlignment"];
            }
            set
            {
                this["horizontalAlignment"] = value;
            }
        }

        /// <summary>
        /// The font for all data
        /// </summary>
        [ConfigurationProperty("font", DefaultValue = "Calibri", IsRequired = false)]
        public String FontName
        {
            get { return (String)this["font"]; }
            set { this["font"] = value; }
        }

        /// <summary>
        /// The font for all data
        /// </summary>
        [ConfigurationProperty("family", DefaultValue = "Swiss", IsRequired = false)]
        public String FontFamily
        {
            get { return (String)this["family"]; }
            set { this["family"] = value; }
        }

        /// <summary>
        /// The font size value for all data
        /// </summary>
        [ConfigurationProperty("size", DefaultValue = "10", IsRequired = false)]
        public String FontSize
        {
            get { return (String)this["size"]; }
            set { this["size"] = value; }
        }

        /// <summary>
        /// The color for all data
        /// </summary>
        [ConfigurationProperty("color", DefaultValue = "A31515", IsRequired = false)]
        [StringValidator(InvalidCharacters = "~!@#$%^&*()[]{}/;'\"|\\GHIJKLMNOPQRSTUVWXYZ", MinLength = 6, MaxLength = 6)]
        public String FontColor
        {
            get { return (String)this["color"]; }
            set { this["color"] = value; }
        }
    }
}
