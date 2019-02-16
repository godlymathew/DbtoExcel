using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace WindowsFormsApp1
{
    public partial class frmMain : Form
    {

        string connectionString = "Data Source={IP/Name}; Initial Catalog = {DatabaseName}}; Persist Security Info=True;User ID = sa; Password=******";
        public frmMain()
        {
            InitializeComponent();
        }

        private void SaveExcel(DataTable dataTable, string newFile)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(newFile);
                ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                pck.SaveAs( new System.IO.FileInfo("d:\\Export\\" + newFile + ".xlsx"));
            }
        }

        private DataTable GetData(string tableName)
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM " + tableName, sqlCon);
                sqlCon.Open();
                var reader = sqlCommand.ExecuteReader();

                DataTable dt = new DataTable();
                dt.Load(reader);
                return dt;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void BtnStartProcess_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < listBox1.Items.Count; i++)
            {
                var fn = listBox1.Items[i].ToString();
                var dt = GetData(fn);
                if (dt != null)
                {
                    SaveExcel(dt, fn);
                }
            }
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM sys.Tables", sqlCon);
                sqlCon.Open();
                var reader = sqlCommand.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        listBox1.Items.Add(reader[0].ToString());
                    }
                }
            }
        }
    }
}
