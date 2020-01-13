using System;
using System.Data.SqlClient;

namespace DLLCreateScheduleExcel.Data
{
    class ConnectionDataBase
    {
        private SqlConnection conn;
        private string strCnx;
        private SqlDataReader dr;

        public void Connect()
        {
            this.strCnx = "Server=OSASHOMOLOG;Initial Catalog=DB_Tgestiona;user id=sa;pwd=osasAdminHomologa";
            this.conn = new SqlConnection(strCnx);
            this.conn.Open();
        }

        public string Query(string query)
        {
            var json = "";
            SqlCommand comm = new SqlCommand(query, this.conn);
            this.dr = comm.ExecuteReader();

            for (int i = 0; i < dr.FieldCount; i++)
            {

                Console.WriteLine("{0} ", dr.GetName(i));
            }
            Console.WriteLine();

            while (dr.Read())
            {
                for (int i = 0; i < dr.FieldCount; i++)
                {
                    json = dr[i].ToString();
                    Console.WriteLine("{0} ", dr[i]);
                }
            }
            this.dr.Close();
            this.conn.Close();
            return json;
        }

        //public void CloseConnection()
        //{
        //    this.dr.Close();
        //    this.conn.Close();
        //}
    }
}
