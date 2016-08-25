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
    public partial class SemestresForm : Form
    {
        #region variables de clase

        public MainForm padre;

        #endregion

        public SemestresForm()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarSemestresForm();
        }

        private void SemestresForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtSemestres;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Semestres";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtSemestres = new DataTable();
                da.Fill(dtSemestres);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Año", typeof(int));
                dt.Columns.Add("Periodo", typeof(int));
                dt.Columns.Add("Limite Tomar Qualify", typeof(DateTime));
                dt.Columns.Add("Limite Propuesta Doctorado", typeof(DateTime));
                dt.Columns.Add("Limite Propuesta Maestria", typeof(DateTime));
                dt.Columns.Add("Limite Pedir Qualify", typeof(DateTime));
                dt.Columns.Add("Limite Tema", typeof(DateTime));

                // se llena el nuevo datatable
                for (int i = 0; i < dtSemestres.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del semestre
                    fila[0] = Convert.ToInt32(dtSemestres.Rows[i][0]);

                    // nombre del semestre
                    fila[1] = dtSemestres.Rows[i][1];

                    // año del semestre
                    fila[2] = Convert.ToInt32(dtSemestres.Rows[i][2]);

                    // periodo del semestre
                    fila[3] = Convert.ToInt32(dtSemestres.Rows[i][3]);

                    // tomar qualify
                    fila[4] = Convert.ToDateTime(dtSemestres.Rows[i][4]);

                    // propuesta doct
                    fila[5] = Convert.ToDateTime(dtSemestres.Rows[i][5]);

                    // propusta maes
                    fila[6] = Convert.ToDateTime(dtSemestres.Rows[i][6]);

                    // pedir qualify
                    fila[7] = Convert.ToDateTime(dtSemestres.Rows[i][7]);

                    // tema
                    fila[8] = Convert.ToDateTime(dtSemestres.Rows[i][8]);

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridSemestres.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                // MainForm.ReescalarDataGridView(ref dataGridSemestres);

                dataGridSemestres.Sort(dataGridSemestres.Columns[0], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
