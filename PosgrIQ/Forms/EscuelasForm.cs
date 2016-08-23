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
using GemBox.Spreadsheet;

namespace PosgrIQ
{
    public partial class EscuelasForm : Form
    {
        #region variables de clase

        public MainForm padre;

        #endregion

        public EscuelasForm()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarEscuelasForm();
        }

        private void EscuelasForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Escuelas";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dt = new DataTable();
                da.Fill(dt);

                
                // se enlaza el datatable con el datagrid
                dataGridEscuelas.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridEscuelas);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
