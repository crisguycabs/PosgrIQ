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
    public partial class PublicacionesDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public PublicacionesDoctForm()
        {
            InitializeComponent();
        }

        private void PublicacionesDoctForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtEstudiantesDoct, dtPublicacionesDoct;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                // se pide la informacion de las publicaciones de doctorado
                query = "SELECT * FROM PublicacionesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtPublicacionesDoct = new DataTable();
                da.Fill(dtPublicacionesDoct);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Revista", typeof(string));
                dt.Columns.Add("Fecha", typeof(string));
                dt.Columns.Add("Alcance", typeof(string));
                dt.Columns.Add("Categoria", typeof(string));
                
                // se llena el nuevo datatable
                for (int i = 0; i < dtPublicacionesDoct.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtPublicacionesDoct.Rows[i][0];

                    // nombre del estudiante
                    fila[1] = dtPublicacionesDoct.Rows[i][1];

                    // titulo de la publicacion
                    fila[2] = dtPublicacionesDoct.Rows[i][2];

                    // revista de la publicacion
                    fila[3] = dtPublicacionesDoct.Rows[i][3];

                    // fecha de la publicacion
                    fila[4] = Convert.ToString(dtPublicacionesDoct.Rows[i][4].ToString());

                    // alcance de la publicacion (Nacional o Internacional)
                    fila[5] = dtPublicacionesDoct.Rows[i][5];

                    // categoria de la revista
                    fila[6] = dtPublicacionesDoct.Rows[i][6];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridPublicaciones.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPublicaciones);

                dataGridPublicaciones.Sort(dataGridPublicaciones.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.padre.CerrarPublicacionesDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddPublicacionesDoctForm agregar = new AddPublicacionesDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.PublicacionesDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridPublicaciones.Rows.Count < 1) return;

            AddPublicacionesDoctForm agregar = new AddPublicacionesDoctForm();
            agregar.padre = this.padre;
            agregar.modo = false;

            if (agregar.ShowDialog() == DialogResult.OK) this.PublicacionesDoctForm_Load(sender, e);
        }
    }
}
