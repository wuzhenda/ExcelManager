using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Proj96333
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "d:\\tmp\\1.xlsx";
            string path2 = "d:\\tmp\\1_tr.xlsx";

            //Util.WriteExcelFile();
            //Util.DF();
            Util.Up();
            Console.ReadLine();
            return;


            var dt = Util.ExcelToDS(path);

            GLog.D("" + dt);

            //Handle(dt);

            //Util.DSToExcel(path, dt);
            Util.ExcuteSQL(dt,"Sheet1", path2);

            Console.ReadLine();
        }

        static void Handle(DataSet ds)
        {
            DataTable dt = ds.Tables[0];

            //Header项
            //foreach (DataColumn dc in dt.Rows[1].Table.Columns)
            //{
            //    //MessageBox.Show(row[dc].ToString());
            //    GLog.D(dc.ColumnName);
            //}

            ////遍历
            ////列项
            //foreach (DataRow row in dt.Rows)
            //{
            //    //横项
            //    foreach (DataColumn dc in row.Table.Columns)
            //    {
            //        GLog.D(row[dc].ToString());
            //    }
            //}


            //列项
            foreach (DataRow row in dt.Rows)
            {
                var regCode = row["注册代码"].ToString();
                var startTime= row["故障发生时间"].ToString();
                var endTime = row["故障恢复时间"].ToString();

                //横项
                foreach (DataColumn dc in row.Table.Columns)
                {
                    if (dc.ColumnName.Equals("是否困人"))
                    {
                        GLog.D(row[dc].ToString());
                    }
                    if (dc.ColumnName.Equals("处置信息"))
                    {
                        GLog.D(row[dc].ToString());
                    }
                }
            }
        }//static void Handle(DataSet ds)

    }
}
