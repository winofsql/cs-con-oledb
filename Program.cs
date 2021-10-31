using System;
using System.Data.OleDb;

namespace cs_con_oledb
{
    class Program
    {
        static void Main(string[] args)
        {
            OleDbConnection myConAccess;

            // *************************************
            // System.Data.OleDb
            // *************************************
            myConAccess = new OleDbConnection();
            myConAccess.ConnectionString =
                string
                    .Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};",
                    @"\app\workspace\subject-1031\cs-con-oledb\販売管理.accdb");

            // 接続を開く
            try
            {
                myConAccess.Open();

                myConAccess.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("接続エラーです:" + ex.Message);
            }
        }
    }
}
