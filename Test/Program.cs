using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PingDevelopment.ExcelML;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = @"Data Source=DatabaseServer;Initial Catalog=DatabaseName;Persist Security Info=False;User ID=Username;Password=Password";
            using (SqlConnection db = new SqlConnection(connectionString))
            {
                db.Open();
                string sql = "SELECT TOP 100 * FROM [TableName]";
                using (SqlCommand cmd = new SqlCommand(sql, db))
                {
                    using (DataTable dt = new DataTable("testData"))
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                            using (ExcelML eml = new ExcelML(dt))
                            {
                                using (StreamWriter sw = new StreamWriter(@"C:\TestXml.xml"))
                                {
                                    XmlDocument xDoc = eml.ConvertDataTableToWorkbook();
                                    sw.Write(xDoc.OuterXml);
                                    sw.Close();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
