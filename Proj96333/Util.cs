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
    public class Util
    {

        public static string SelectPath()
        {
            string path = string.Empty;
            //var openFileDialog = new Microsoft.Win32.OpenFileDialog()
            //{
            //    Filter = "Files (*.xls)|*.xls|(*.xlsx)|*.xlsx"//如果需要筛选txt文件（"Files (*.txt)|*.txt"）
            //    //Filter = "Files (全部文件)|*.*"
            //};
            //var result = openFileDialog.ShowDialog();
            //if (result == true)
            //{
            //    path = openFileDialog.FileName;
            //}
            return path;
        }


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

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }


         

        public static void DSToExcel(string originalFilePath, DataSet oldds)
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

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            //return ds;


            OleDbConnection myConn = new OleDbConnection(connectionString);
            string strCom = "select * from [Sheet1$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            System.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);
            //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。 
            //获取insert语句中保留字符（起始位置）
            builder.QuotePrefix = "[";
            //获取insert语句中保留字符（结束位置） 
            builder.QuoteSuffix = "]";

            DataSet newds = new DataSet();

            myCommand.Fill(newds, "Table1");
            for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
            {
                //在这里不能使用ImportRow方法将一行导入到news中，因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
                // 在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState != Added               
                DataRow nrow = newds.Tables["Table1"].NewRow();
                for (int j = 0; j < newds.Tables[0].Columns.Count; j++)
                {
                    nrow[j] = oldds.Tables[0].Rows[i][j];
                }
                newds.Tables["Table1"].Rows.Add(nrow);
            }

            myCommand.Update(newds, "Table1");
            myConn.Close();
        }

        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="oldds">需要导入的数据</param>
        /// <param name="TableName">表名</param>
        /// <param name="strCon">连接字符串</param>
        public static void ExcuteSQL(DataSet oldds, string TableName, string filePath)
        {
            string connectionString = GetConnectionString(filePath);
            //连接
            OleDbConnection myConn = new OleDbConnection(connectionString);

            string strCom = "select * from [" + TableName + "$]";

            try
            {
                myConn.Open();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);

                System.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);

                //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。   
                //获取insert语句中保留字符（起始位置）  
                builder.QuotePrefix = "[";

                //获取insert语句中保留字符（结束位置）   
                builder.QuoteSuffix = "]";

                DataSet newds = new DataSet();
                //获得表结构
                DataTable ndt = oldds.Tables[0].Clone();
                //清空数据
                //ndt.Rows.Clear();

                ndt.TableName = TableName;
                newds.Tables.Add(ndt);

                //myCommand.Fill(newds, TableName);

                for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
                {
                    //在这里不能使用ImportRow方法将一行导入到news中，
                    //因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
                    //在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState!=Added     
                    DataRow nrow = newds.Tables[0].NewRow();
                    for (int j = 0; j < oldds.Tables[0].Columns.Count; j++)
                    {
                        nrow[j] = oldds.Tables[0].Rows[i][j];
                    }
                    newds.Tables[0].Rows.Add(nrow);
                }
                //插入数据
                myCommand.Update(newds, TableName);
            }
            finally
            {
                myConn.Close();
            }
        }

        public static void Up()
        {
            string path = "d:\\tmp\\1.xlsx";

            try
            {
                string connectionString = GetConnectionString(path);
                //连接
                using (OleDbConnection MyConnection = new OleDbConnection(connectionString))
                {
                    OleDbCommand myCommand = new OleDbCommand();

                    MyConnection.Open();
                    myCommand.Connection = MyConnection;

                    myCommand.CommandText = "Update [Sheet1$] set 是否困人 = 'Y' where 注册代码='30103205842016070110'";
                    myCommand.ExecuteNonQuery();
                    MyConnection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void DF()
        {

            string xlsFile = @"D:\\tmp\\AdapterTest.xlsx";
            string xlsSheet = @"Sheet2$";

            // HDR=Yes means that the first row in the range is the header row (or field names) by default.
            // If the first range does not contain headers, you can specify HDR=No in the extended properties in your connection string.
            //string connectionstring = string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;'", xlsFile);
            string connectionstring = GetConnectionString(xlsFile);

            // Create connection
            OleDbConnection oleDBConnection = new OleDbConnection(connectionstring);

            // Create the dataadapter with the select to get all rows in in the xls
            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT FirstName, LastName, Age FROM [" + xlsSheet + "]", oleDBConnection);

            // Since there is no pk in Excel, using a command builder will not help here.
            //OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
            // Create the dataset and fill it by using the adapter.

            DataSet dataset = new DataSet();

            //adapter.Fill(dataset);//, "PersonTable");
            adapter.Fill(dataset, "Sheet2");


            // Time to create the INSERT/UPDATE commands for the Adapter,
            // the way to do this is to use parameterized commands.

            // *** INSERT COMMAND ***
            adapter.InsertCommand = new OleDbCommand("INSERT INTO [" + xlsSheet + "] (FirstName, LastName, Age) VALUES (?, ?, ?)", oleDBConnection);

            adapter.InsertCommand.Parameters.Add("@FirstName", OleDbType.VarChar, 255).SourceColumn = "FirstName";
            adapter.InsertCommand.Parameters.Add("@LastName", OleDbType.Char, 255).SourceColumn = "LastName";
            adapter.InsertCommand.Parameters.Add("@Age", OleDbType.Char, 255).SourceColumn = "Age";
            
            // *** UPDATE COMMAND ***
            adapter.UpdateCommand = new OleDbCommand("UPDATE [" + xlsSheet + "] SET FirstName = ?, LastName = ?, Age = ?" +
                                                        " WHERE FirstName = ? AND LastName = ? AND Age = ?", oleDBConnection);

            adapter.UpdateCommand.Parameters.Add("@FirstName", OleDbType.Char, 255).SourceColumn = "FirstName";
            adapter.UpdateCommand.Parameters.Add("@LastName", OleDbType.Char, 255).SourceColumn = "LastName";
            adapter.UpdateCommand.Parameters.Add("@Age", OleDbType.Char, 255).SourceColumn = "Age";

            // For Updates, we need to provide the old values so that we only update the corresponding row.
            adapter.UpdateCommand.Parameters.Add("@OldFirstName", OleDbType.Char, 255, "FirstName").SourceVersion = DataRowVersion.Original;
            adapter.UpdateCommand.Parameters.Add("@OldLastName", OleDbType.Char, 255, "LastName").SourceVersion = DataRowVersion.Original;
            adapter.UpdateCommand.Parameters.Add("@OldAge", OleDbType.Char, 255, "Age").SourceVersion = DataRowVersion.Original;

            // Insert a new row
            DataRow newPersonRow = dataset.Tables[0].NewRow();

            newPersonRow["FirstName"] = "New";
            newPersonRow["LastName"] = "Person";
            newPersonRow["Age"] = "100";

            dataset.Tables[0].Rows.Add(newPersonRow);

            // Updates the first row
            dataset.Tables[0].Rows[0]["FirstName"] = "Updated";
            dataset.Tables[0].Rows[0]["LastName"] = "Person";
            dataset.Tables[0].Rows[0]["Age"] = "55";


            // Call update on the adapter to save all the changes to the dataset
            adapter.Update(dataset);
        }

        public static void WriteExcelFile()
        {
            string xlsFile = @"D:\\tmp\\AdapterTest.xlsx";
            string xlsSheet = @"Sheet1$";

            string connectionstring = GetConnectionString(xlsFile);

            //string connectionstring = GetConnectionString();
            using (OleDbConnection conn = new OleDbConnection(connectionstring))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = "CREATE TABLE [table1] (id INT, name VARCHAR, datecol DATE );";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(1,'AAAA','2014-01-01');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(2, 'BBBB','2014-01-03');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(3, 'CCCC','2014-01-03');";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE [table1] SET name = 'DDDD' WHERE id = 3;";
                cmd.ExecuteNonQuery();

                conn.Close();
            }
        }
    }
}
