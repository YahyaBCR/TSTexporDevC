using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Odbc;

namespace TSTexporDevC
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            d.connect();

        }
        System.Data.DataSet DtSet;
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.OleDb.OleDbConnection Myconnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                Myconnection = new System.Data.OleDb.OleDbConnection("provider =Microsoft.Jet.OLEDB.4.0;Data Source = 'C:\\Users\\dell\\Desktop\\StgrExp.xls'; Extended Properties = Excel 8.0");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Feuil1$]", Myconnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];
               
                Myconnection.Close();

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        ADO d = new ADO();
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                d.cmd.CommandText = "insert into Stagiaire values(" + dataGridView1.Rows[i].Cells[0].Value+",'"+dataGridView1.Rows[i].Cells[1].Value+"','"+dataGridView1.Rows[i].Cells[1].Value+"')";
                d.cmd.Connection = d.cnx;
                d.cmd.ExecuteNonQuery();
            }
        }
    }
}