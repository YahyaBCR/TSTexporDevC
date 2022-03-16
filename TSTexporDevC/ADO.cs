using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace TSTexporDevC
{
    class ADO
    {
        public SqlConnection cnx = new SqlConnection(@"Data Source=DESKTOP-SVQ4VSE\SQLEXPRESS01;Initial Catalog=ExportStgrDevC;Integrated Security=True");
        public SqlCommand cmd = new SqlCommand();
        public SqlDataReader dr;
        public DataTable dt = new DataTable();
        public SqlDataAdapter dap;


        public void connect()
        {
            if (cnx.State == ConnectionState.Closed || cnx.State == ConnectionState.Broken)
            {
                cnx.Open();
            }
        }

        public void deconnect()
        {
            if (cnx.State == ConnectionState.Open)
            {
                cnx.Close();
            }
        }
    }
}
