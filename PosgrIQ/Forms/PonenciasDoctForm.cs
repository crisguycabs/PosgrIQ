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
    public partial class PonenciasDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public PonenciasDoctForm()
        {
            InitializeComponent();
        }

        private void PonenciasDoctForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtEstudiantesDoct, dtPonenciasDoct;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                // se pide la informacion de las ponencias de doctorado
                query = "SELECT * FROM PonenciasDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtPonenciasDoct = new DataTable();
                da.Fill(dtPonenciasDoct);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Revista", typeof(string));
                dt.Columns.Add("Fecha", typeof(string));
                dt.Columns.Add("Alcance", typeof(string));
                
                // se llena el nuevo datatable
                for (int i = 0; i < dtPonenciasDoct.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtPonenciasDoct.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesDoct.Select("codigo=" + dtPonenciasDoct.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // titulo de la publicacion
                    fila[2] = dtPonenciasDoct.Rows[i][2];

                    // revista de la publicacion
                    fila[3] = dtPonenciasDoct.Rows[i][3];

                    // fecha de la publicacion
                    fila[4] = Convert.ToString(dtPonenciasDoct.Rows[i][4].ToString());

                    // alcance de la publicacion (Nacional o Internacional)
                    fila[5] = dtPonenciasDoct.Rows[i][5];                    

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridPonencias.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPonencias);

                dataGridPonencias.Sort(dataGridPonencias.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Ponencias Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarPonenciasDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddPonenciasDoctForm agregar = new AddPonenciasDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.PonenciasDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridPonencias.Rows.Count < 1) return;

            AddPonenciasDoctForm agregar = new AddPonenciasDoctForm();
            agregar.padre = this.padre;
            agregar.modo = false;

            if (agregar.ShowDialog() == DialogResult.OK) this.PonenciasDoctForm_Load(sender, e);
        }
    }
}
