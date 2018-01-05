using CannyData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Proj96333
{
    public partial class Util
    {

        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        public static DataSet ExcelToDS(string file)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                //refs:https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbconnection.getoledbschematable(v=vs.110).aspx
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in schemaTable.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    //create a new table corresponding to a sheet in excel
                    DataTable sheetDataTable = new DataTable();
                    sheetDataTable.TableName = sheetName;

                    //remarks https://msdn.microsoft.com/en-us/library/system.data.oledb.oledbdataadapter(v=vs.110).aspx#Remarks
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(sheetDataTable);

                    //add to total dataset
                    ds.Tables.Add(sheetDataTable);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        
        public static void UpdateExcel(string originalFilePath)
        {
            DataSet ds = new DataSet();
            string connectionString = GetConnectionString(originalFilePath);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    //here table has data now,so we do it
                    //列项
                    foreach (DataRow row in dt.Rows)
                    {
                        var regCode = row["注册代码"].ToString();
                        var startTime = row["故障发生时间"].ToString();
                        var endTime = row["故障恢复时间"].ToString();

                        var etor_bianhao = DbHelper.GetCannyEtorBianhaoByRegCodeFromLocalDb(regCode);
                        var ret = DbHelper.GetDtuStatusFromCannyDbByDtuBianhao(etor_bianhao, startTime);

                        if (ret != null)

                        {
                            var isTrap = ret.Item1;
                            var solution = ret.Item2;

                            //var regCodeInt = int.Parse(regCode.Substring(14));
                            //string trap = (0 == regCodeInt % 2) ? "Y" : "N";

                            cmd.CommandText = $"Update [{sheetName}] set 是否困人 = '{isTrap}',处置信息='{solution}' where 注册代码='{regCode}'";
                            cmd.ExecuteNonQuery();
                        }
                    }
                    
                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }
        }
        
    }
}
