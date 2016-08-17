using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace PosgrIQ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                var query = "SELECT * FROM Profesores";
                var command = new OleDbCommand(query, conection);

                /*
                var reader = command.ExecuteReader();

                int count = 0;
                while (reader.Read())
                {
                    count++;
                }
                 * 

                MessageBox.Show(count.ToString());*/

                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.DataSource = dt;

                conection.Close();
            }
            catch
            {


            }

        }
    }
}
