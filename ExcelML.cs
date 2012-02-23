using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;
using WellsFargo.ExcelML.Configuration;
using System.Configuration;

namespace WellsFargo.ExcelML
{
    public class ExcelML : IDisposable
    {
        #region Private Members
        private string _version = "1.0";
        private string _dateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ";
        private string _author = "ExcelML Library";
        private DateTime _dateCreated = DateTime.Now;
        private DateTime _dateModified = DateTime.Now;
        private string _company = "Wells Fargo &amp; Co.";
        private string _documentVersion = "12.00";

        private DataSet _inputData = null;
        private XmlDocument _outputDocument = null;
        private string _attributeProgId = "Excel.Sheet";
        private ExcelMlConfigSection _configSection = null;

        private string _officeNamespace = "urn:schemas-microsoft-com:office:office";
        private string _excelNamespace = "urn:schemas-microsoft-com:office:excel";
        private string _spreadsheetNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
        private string _htmlNamespace = "http://www.w3.org/TR/REC-html40";

        private string _officePrefix = "o";
        private string _excelPrefix = "x";
        private string _spreadsheetPrefix = "ss";
        private string _htmlPrefix = "html";
        #endregion

        #region Properties
        /// <summary>
        /// XML document Version
        /// </summary>
        public string Version
        {
            get { return _version; }
            set { _version = value; }
        }

        /// <summary>
        /// Format string for DateTime values
        /// </summary>
        public string DateTimeFormat
        {
            get { return _dateTimeFormat; }
            set { _dateTimeFormat = value; }
        }

        /// <summary>
        /// Author of the document. Defaults to ExcelML Library.
        /// </summary>
        public string Author
        {
            get { return _author; }
            set { _author = value; }
        }

        /// <summary>
        /// Date the document was created.
        /// </summary>
        public DateTime DateCreated
        {
            get { return _dateCreated; }
            set { _dateModified = value; }
        }

        /// <summary>
        /// Date the document was last modified
        /// </summary>
        public DateTime DateModified
        {
            get { return _dateModified; }
            set { _dateModified = value; }
        }

        /// <summary>
        /// Company name
        /// </summary>
        /// <remarks>Assumes correct XML encoding on characters such as ampersands and quotes</remarks>
        public string Company
        {
            get { return _company; }
            set { _company = value; }
        }

        /// <summary>
        /// Input DataTable to convert to an Excel XML workbook
        /// </summary>
        public DataSet InputData
        {
            get { return _inputData; }
            set { _inputData = value; }
        }

        /// <summary>
        /// Document generated from InputData
        /// </summary>
        public XmlDocument OutputDocument
        {
            get { return _outputDocument; }
        }

        /// <summary>
        /// Configuration section data
        /// </summary>
        private ExcelMlConfigSection Config
        {
            get { return _configSection; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelML()
            : this(null)
        {
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="input"></param>
        public ExcelML(DataTable input)
        {
            _outputDocument = new XmlDocument();
            _inputData = new DataSet("ExcelMLInput");
            if (input != null) _inputData.Tables.Add(input);
            
            // Get the configuration data if available
            _configSection = (ConfigurationManager.GetSection("excelMl") as ExcelMlConfigSection) ?? new ExcelMlConfigSection();
        }
        #endregion

        /// <summary>
        /// Convert the current input data into an Excel XML formatted document
        /// </summary>
        /// <returns></returns>
        public XmlDocument ConvertDataTableToWorkbook()
        {
            return ConvertDataTableToWorkbook(InputData);
        }

        /// <summary>
        /// Create a workbook using the specified input table
        /// </summary>
        /// <param name="dtInput"></param>
        /// <returns></returns>
        public XmlDocument ConvertDataTableToWorkbook(DataTable dtInput)
        {
            return ConvertDataTableToWorkbook(dtInput, false);
        }

        /// <summary>
        /// Create a workbook using the specified input table and optionally
        /// append the input table to current input data.
        /// </summary>
        /// <param name="dtInput"></param>
        /// <param name="appendInput"></param>
        /// <returns></returns>
        public XmlDocument ConvertDataTableToWorkbook(DataTable dtInput, bool appendInput)
        {
            // Clear the current input if not appending the data
            if (!appendInput) _inputData.Tables.Clear();

            _inputData.Tables.Add(dtInput);
            return ConvertDataTableToWorkbook(InputData);
        }

        /// <summary>
        /// Convert the DataTable object into an Excel XML formatted document and
        /// assign new input data.
        /// </summary>
        /// <param name="dtInput">New InputData data table</param>
        /// <returns></returns>
        public XmlDocument ConvertDataTableToWorkbook(DataSet dtInput)
        {
            // Assign the argument as the new input
            InputData = dtInput;

            // Initialize the base document information
            CreateDocumentHeader();
            CreateWorkbook();

            return _outputDocument;
        }

        #region Private Methods
        /// <summary>
        /// Add the namespacing and global information used for the Excel XML format
        /// </summary>
        private void CreateDocumentHeader()
        {
            string headerString =
                String.Format("<?xml version=\"{0}\"?>", Version) + System.Environment.NewLine +
                String.Format("<?mso-application progid=\"{0}\"?>", _attributeProgId) + System.Environment.NewLine +
                String.Format(
                                "<Workbook xmlns=\"{0}\" xmlns:{1}=\"{2}\" xmlns:{3}=\"{4}\" xmlns:{5}=\"{6}\" xmlns:{7}=\"{8}\"/>",
                                _spreadsheetNamespace,
                                _officePrefix, _officeNamespace,
                                _excelPrefix, _excelNamespace,
                                _spreadsheetPrefix, _spreadsheetNamespace,
                                _htmlPrefix, _htmlNamespace
                             );

            // Add the namespaces
            XmlNamespaceManager nsm = new XmlNamespaceManager(_outputDocument.NameTable);
            nsm.AddNamespace(String.Empty, _spreadsheetNamespace);
            nsm.AddNamespace(_officePrefix, _officeNamespace);
            nsm.AddNamespace(_excelPrefix, _excelNamespace);
            nsm.AddNamespace(_spreadsheetPrefix, _spreadsheetNamespace);
            nsm.AddNamespace(_htmlPrefix, _htmlNamespace);

            // Add the XML declarations
            //_outputDocument.AppendChild(_outputDocument.CreateXmlDeclaration(Version, String.Empty, null));
            //_outputDocument.AppendChild(_outputDocument.CreateNode(XmlNodeType.XmlDeclaration, "mso-application", String.Empty));
            //_outputDocument.LastChild.Attributes.Append(_outputDocument.CreateAttribute("progid"));
            //_outputDocument.LastChild.Attributes["progid"].Value = _attributeProgId;

            // Add the document element
            //_outputDocument.AppendChild(_outputDocument.CreateElement("Workbook", _spreadsheetNamespace));
            //_outputDocument.LastChild.Attributes.Append((XmlAttribute)_outputDocument.CreateNode(XmlNodeType.Attribute, _officePrefix, "xmlns", _officeNamespace));
            //_outputDocument.LastChild.Attributes.Append((XmlAttribute)_outputDocument.CreateNode(XmlNodeType.Attribute, _excelPrefix, "xmlns", _excelNamespace));
            //_outputDocument.LastChild.Attributes.Append((XmlAttribute)_outputDocument.CreateNode(XmlNodeType.Attribute, _spreadsheetPrefix, "xmlns", _spreadsheetNamespace));
            //_outputDocument.LastChild.Attributes.Append((XmlAttribute)_outputDocument.CreateNode(XmlNodeType.Attribute, _htmlPrefix, "xmlns", _htmlNamespace));

            _outputDocument.LoadXml(headerString);
        }

        /// <summary>
        /// Create the root workbook node
        /// </summary>
        private void CreateWorkbook()
        {
            XmlElement wb;
            if (!_outputDocument.DocumentElement.Name.ToLower().Equals("workbook"))
            {
                wb = _outputDocument.CreateElement("Workbook", _spreadsheetNamespace);
            }
            else
            {
                wb = _outputDocument.DocumentElement;
            }

            // Append required namespaces

            // Append the DocumentProperties node
            wb.AppendChild(CreateDocumentProperties());

            // Append the ExcelWorkbook node
            wb.AppendChild(CreateExcelWorkbook());

            // Append the Styles node
            wb.AppendChild(_outputDocument.CreateElement("Styles", _spreadsheetNamespace));

            wb.LastChild.AppendChild(CreateDefaultStyle());
            wb.LastChild.AppendChild(CreateHeaderStyle());
            wb.LastChild.AppendChild(CreateDataStyle());
            wb.LastChild.AppendChild(CreateDateTimeStyle());

            foreach (DataTable t in InputData.Tables)
            {
                wb.AppendChild(CreateWorksheet(t.TableName, t));
            }

            // Don't need to append it if it's already there
            if (!_outputDocument.FirstChild.Name.ToLower().Equals("workbook"))  
                _outputDocument.AppendChild(wb);
        }

        /// <summary>
        /// Create the style attributes for all data cells
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateDataStyle()
        {
            // Add the header format styles from configuration
            XmlElement dataStyle = _outputDocument.CreateElement("Style", _spreadsheetNamespace);
            dataStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "ID", _spreadsheetNamespace));
            dataStyle.Attributes[_spreadsheetPrefix + ":ID"].Value = "xlData";

            dataStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Name", _spreadsheetNamespace));
            dataStyle.Attributes[_spreadsheetPrefix + ":Name"].Value = "DataCell";

            dataStyle.AppendChild(_outputDocument.CreateElement("Alignment", _spreadsheetNamespace));
            
            // Add the default vertical alignment
            //dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Vertical", _spreadsheetNamespace));
            //dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Vertical"].Value = Config.DataFormat.StringFormat.VerticalAlignment;

            // Add the default horizontal alignment
            //dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Horizontal", _spreadsheetNamespace));
            //dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Horizontal"].Value = Config.DataFormat.StringFormat.HorizontalAlignment;

            // Add the numeric data format settings
            dataStyle.AppendChild(_outputDocument.CreateElement("NumberFormat", _spreadsheetNamespace));
            if (Config.DataFormat.NumericFormat.FormatString.Length > 0)
            {
                // - Numeric format
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Format", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Format"].Value = Config.DataFormat.NumericFormat.FormatString;
            }

            // Add the data font settings
            dataStyle.AppendChild(_outputDocument.CreateElement("Font", _spreadsheetNamespace));
            if (Config.DataFormat.StringFormat.FontName.Length > 0)
            {
                // - Font name
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "FontName", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":FontName"].Value = Config.DataFormat.StringFormat.FontName;
            }

            if (Config.DataFormat.StringFormat.FontFamily.Length > 0)
            {
                // - Font family
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Family", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Family"].Value = Config.DataFormat.StringFormat.FontFamily;
            }

            if (Config.DataFormat.StringFormat.FontSize.Length > 0)
            {
                // - Size (points)
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Size", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Size"].Value = Config.DataFormat.StringFormat.FontSize;
            }

            if (Config.DataFormat.StringFormat.FontColor.Length > 0)
            {
                // - Color (RGB)
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Color", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Color"].Value = String.Format("#{0}", Config.DataFormat.StringFormat.FontColor);
            }

            if (Config.DataFormat.StringFormat.Bold != -1)
            {
                // Add the header bolding
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Bold", _spreadsheetNamespace));
                //style.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = "1";
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = Config.DataFormat.StringFormat.Bold.ToString("F0");
            }
            return dataStyle;
        }

        /// <summary>
        /// Create the style attributes for all data cells
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateDateTimeStyle()
        {
            // Add the header format styles from configuration
            XmlElement dataStyle = _outputDocument.CreateElement("Style", _spreadsheetNamespace);
            dataStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "ID", _spreadsheetNamespace));
            dataStyle.Attributes[_spreadsheetPrefix + ":ID"].Value = "xlDateTime";

            dataStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Name", _spreadsheetNamespace));
            dataStyle.Attributes[_spreadsheetPrefix + ":Name"].Value = "DateTimeCell";

            dataStyle.AppendChild(_outputDocument.CreateElement("Alignment", _spreadsheetNamespace));
            if (Config.DataFormat.DateFormat.VerticalAlignment.Length > 0)
            {
                // Add the default vertical alignment
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Vertical", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Vertical"].Value = Config.DataFormat.DateFormat.VerticalAlignment;
            }

            if (Config.DataFormat.DateFormat.HorizontalAlignment.Length > 0)
            {
                // Add the default horizontal alignment
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Horizontal", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Horizontal"].Value = Config.DataFormat.DateFormat.HorizontalAlignment;
            }

            // Add the numeric data format settings
            dataStyle.AppendChild(_outputDocument.CreateElement("NumberFormat", _spreadsheetNamespace));
            if (Config.DataFormat.DateFormat.FormatString.Length > 0)
            {
                // - Font name
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Format", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Format"].Value = Config.DataFormat.DateFormat.FormatString;
            }

            // Add the data font settings
            dataStyle.AppendChild(_outputDocument.CreateElement("Font", _spreadsheetNamespace));
            if (Config.DataFormat.StringFormat.FontName.Length > 0)
            {
                // - Font name
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "FontName", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":FontName"].Value = Config.DataFormat.StringFormat.FontName;
            }

            if (Config.DataFormat.StringFormat.FontFamily.Length > 0)
            {
                // - Font family
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Family", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Family"].Value = Config.DataFormat.StringFormat.FontFamily;
            }

            if (Config.DataFormat.StringFormat.FontSize.Length > 0)
            {
                // - Size (points)
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Size", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Size"].Value = Config.DataFormat.StringFormat.FontSize;
            }

            if (Config.DataFormat.StringFormat.FontColor.Length > 0)
            {
                // - Color (RGB)
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Color", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Color"].Value = String.Format("#{0}", Config.DataFormat.StringFormat.FontColor);
            }

            if (Config.DataFormat.DateFormat.Bold != -1)
            {
                // Add the header bolding
                dataStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Bold", _spreadsheetNamespace));
                dataStyle.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = Config.DataFormat.DateFormat.Bold.ToString("F0");
            }
            return dataStyle;
        }

        /// <summary>
        /// Create the style attributes for all column header cells
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateHeaderStyle()
        {
            // Add the header format styles from configuration
            XmlElement headerStyle = _outputDocument.CreateElement("Style", _spreadsheetNamespace);
            headerStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "ID", _spreadsheetNamespace));
            headerStyle.Attributes[_spreadsheetPrefix + ":ID"].Value = "xlHeader";

            headerStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Name", _spreadsheetNamespace));
            headerStyle.Attributes[_spreadsheetPrefix + ":Name"].Value = "HeaderCell";

            headerStyle.AppendChild(_outputDocument.CreateElement("Alignment", _spreadsheetNamespace));

            // Add the default vertical alignment
            if (Config.HeaderFormat.VerticalAlignment.Length > 0)
            {
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Vertical", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Vertical"].Value = Config.HeaderFormat.VerticalAlignment;
            }

            // Add the default horizontal alignment
            if (Config.HeaderFormat.HorizontalAlignment.Length > 0)
            {
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Horizontal", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Horizontal"].Value = Config.HeaderFormat.HorizontalAlignment;
            }

            // Add the header font settings
            headerStyle.AppendChild(_outputDocument.CreateElement("Font", _spreadsheetNamespace));
            if (Config.HeaderFormat.FontName.Length > 0)
            {
                // - Font name/family
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "FontName", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":FontName"].Value = Config.HeaderFormat.FontName;
            }

            if (Config.HeaderFormat.FontFamily.Length > 0)
            {
                // - Font name/family
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Family", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Family"].Value = Config.HeaderFormat.FontFamily;
            }

            if (Config.HeaderFormat.FontSize.Length > 0)
            {
                // - Size (points)
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Size", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Size"].Value = Config.HeaderFormat.FontSize;
            }

            if (Config.HeaderFormat.FontColor.Length > 0)
            {
                // - Color (RGB)
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Color", _spreadsheetNamespace));
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Color"].Value = String.Format("#{0}", Config.HeaderFormat.FontColor);
            }

            if (Config.HeaderFormat.Bold != -1)
            {
                // Add the header bolding
                headerStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Bold", _spreadsheetNamespace));
                //style.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = "1";
                headerStyle.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = Config.HeaderFormat.Bold.ToString("F0");
            }

            return headerStyle;
        }

        /// <summary>
        /// Create the style attributes used by default
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateDefaultStyle()
        {
            XmlElement defaultStyle = _outputDocument.CreateElement("Style", _spreadsheetNamespace);
            defaultStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "ID", _spreadsheetNamespace));
            defaultStyle.Attributes[_spreadsheetPrefix + ":ID"].Value = "Default";

            defaultStyle.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Name", _spreadsheetNamespace));
            defaultStyle.Attributes[_spreadsheetPrefix + ":Name"].Value = "Normal";

            // Add the default alignment settings
            defaultStyle.AppendChild(_outputDocument.CreateElement("Alignment", _spreadsheetNamespace));
            if (Config.DefaultFormat.VerticalAlignment.Length > 0)
            {
                // Add the default vertical alignment
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Vertical", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Vertical"].Value = Config.DefaultFormat.VerticalAlignment;
            }

            if (Config.DefaultFormat.HorizontalAlignment.Length > 0)
            {
                // Add the default horizontal alignment
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Horizontal", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Horizontal"].Value = Config.DefaultFormat.HorizontalAlignment;
            }

            // Add the default font settings
            defaultStyle.AppendChild(_outputDocument.CreateElement("Font", _spreadsheetNamespace));
            if (Config.DefaultFormat.FontName.Length > 0)
            {
                // - Font name
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "FontName", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":FontName"].Value = Config.DefaultFormat.FontName;
            }

            if (Config.DefaultFormat.FontFamily.Length > 0)
            {
                // - Font family
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Family", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Family"].Value = Config.DefaultFormat.FontFamily;
            }

            if (Config.DefaultFormat.FontSize.Length > 0)
            {
                // - Size (points)
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Size", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Size"].Value = Config.DefaultFormat.FontSize;
            }

            if (Config.DefaultFormat.FontColor.Length > 0)
            {
                // - Color (RGB)
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Color", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Color"].Value = String.Format("#{0}", Config.DefaultFormat.FontColor);
            }

            if (Config.DefaultFormat.Bold != -1)
            {
                // Add the default bolding
                defaultStyle.LastChild.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Bold", _spreadsheetNamespace));
                defaultStyle.LastChild.Attributes[_spreadsheetPrefix + ":Bold"].Value = Config.DefaultFormat.Bold.ToString("F0");
            }

            return defaultStyle;
        }

        /// <summary>
        /// Create the DocumentProperties element
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateDocumentProperties()
        {
            XmlElement dp = _outputDocument.CreateElement("DocumentProperties", _officeNamespace);

            // Author element
            dp.AppendChild(_outputDocument.CreateElement("Author", _officeNamespace));
            dp.LastChild.InnerText = Author;

            // LastAuthor element
            dp.AppendChild(_outputDocument.CreateElement("LastAuthor", _officeNamespace));
            dp.LastChild.InnerText = Author;

            // Created element (date created)
            dp.AppendChild(_outputDocument.CreateElement("Created", _officeNamespace));
            dp.LastChild.InnerText = DateCreated.ToString(_dateTimeFormat);

            // LastSaved element (date modified)
            dp.AppendChild(_outputDocument.CreateElement("LastSaved", _officeNamespace));
            dp.LastChild.InnerText = DateModified.ToString(_dateTimeFormat);

            // Company element, only if not blank
            if (Company.Trim().Length > 0)
            {
                dp.AppendChild(_outputDocument.CreateElement("Company", _officeNamespace));
                dp.LastChild.InnerText = Company;
            }

            // Version element (document version)
            dp.AppendChild(_outputDocument.CreateElement("Version", _officeNamespace));
            dp.LastChild.InnerText = _documentVersion;

            return dp;
        }

        /// <summary>
        /// Create the ExcelWorkbook element
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateExcelWorkbook()
        {
            XmlElement ewb = _outputDocument.CreateElement("ExcelWorkbook", _excelNamespace);
            ewb.AppendChild(_outputDocument.CreateElement("ProtectStructure", _excelNamespace));
            ewb.LastChild.InnerText = "False";
            ewb.AppendChild(_outputDocument.CreateElement("ProtectWindows", _excelNamespace));
            ewb.LastChild.InnerText = "False"; 
            return ewb;
        }

        /// <summary>
        /// Create the worksheet of the specified name
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private XmlElement CreateWorksheet(string sheetName, DataTable t)
        {
            XmlElement ws = _outputDocument.CreateElement("Worksheet", _spreadsheetNamespace);
            ws.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Name", _spreadsheetNamespace));
            ws.Attributes[_spreadsheetPrefix + ":Name"].Value = sheetName;
            ws.AppendChild(CreateTable(t));

            // Add the options for this worksheet
            ws.AppendChild(CreateWorksheetOptions());

            return ws;
        }

        /// <summary>
        /// Create a default set of worksheet options for this element
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateWorksheetOptions()
        {
            XmlElement wso = _outputDocument.CreateElement("WorksheetOptions", _spreadsheetNamespace);
            
            return wso;
        }

        /// <summary>
        /// Create the table element used to contain rows
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateTable(DataTable t)
        {
            XmlElement tbl = _outputDocument.CreateElement("Table", _spreadsheetNamespace);
            tbl.AppendChild(CreateColumnHeaders(t.Columns));

            // Append all the DataRows as row elements
            foreach (DataRow dr in t.Rows)
            {
                tbl.AppendChild(CreateRow(dr));
            }

            return tbl;
        }

        /// <summary>
        /// Create a row of column headers in the first row of the worksheet
        /// </summary>
        /// <returns></returns>
        private XmlElement CreateColumnHeaders(DataColumnCollection columns)
        {
            XmlElement row = _outputDocument.CreateElement("Row", _spreadsheetNamespace);

            foreach (DataColumn c in columns)
            {
                XmlElement cell = _outputDocument.CreateElement("Cell", _spreadsheetNamespace);
                cell.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "StyleID", _spreadsheetNamespace));
                cell.Attributes[_spreadsheetPrefix + ":StyleID"].Value = "xlHeader";
                
                XmlElement cellData = _outputDocument.CreateElement("Data", _spreadsheetNamespace);
                cellData.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Type", _spreadsheetNamespace));
                cellData.Attributes[_spreadsheetPrefix + ":Type"].Value = "String";
                cellData.InnerText = c.ColumnName.ToUpper();

                cell.AppendChild(cellData);
                row.AppendChild(cell);
            }

            return row;
        }

        /// <summary>
        /// Convert the DataRow object to a Row element
        /// </summary>
        /// <param name="dr"></param>
        /// <returns></returns>
        private XmlElement CreateRow(DataRow dr)
        {
            XmlElement row = _outputDocument.CreateElement("Row", _spreadsheetNamespace);
            //XmlElement row = _outputDocument.CreateElement("tr");

            foreach (DataColumn dc in dr.Table.Columns)
            {
                // Determine the data type for the cell data
                string type = "String";
                switch (dc.DataType.Name.ToLower())
                {
                    case "decimal":
                    case "double":
                    case "int16":
                    case "int32":
                    case "int64":
                        //type = "Number";
                        type = "Number";
                        break;

                    case "datetime":
                        type = "DateTime";
                        break;
                }
                
                XmlElement cell = _outputDocument.CreateElement("Cell", _spreadsheetNamespace);
                cell.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "StyleID", _spreadsheetNamespace));
                cell.Attributes[_spreadsheetPrefix + ":StyleID"].Value = "xlData";

                XmlElement cellData = _outputDocument.CreateElement("Data", _spreadsheetNamespace);
                cellData.Attributes.Append(_outputDocument.CreateAttribute(_spreadsheetPrefix, "Type", _spreadsheetNamespace));
                cellData.Attributes[_spreadsheetPrefix + ":Type"].Value = type;

                if (dr[dc] is DateTime)
                {
                    cellData.InnerText = ((DateTime)dr[dc]).ToString(_dateTimeFormat);
                    cell.Attributes[_spreadsheetPrefix + ":StyleID"].Value = "xlDateTime";
                }
                else if (dr[dc] != DBNull.Value)
                {
                    cellData.InnerText = dr[dc].ToString();
                }

                cell.AppendChild(cellData);
                row.AppendChild(cell);
            }

            return row;
        }
        #endregion

        #region IDisposable Members

        /// <summary>
        /// Destroy input and output data and force garbage collection
        /// </summary>
        public void Dispose()
        {
            _inputData = null;
            _outputDocument = null;
            GC.Collect();
        }

        #endregion
    }
}
