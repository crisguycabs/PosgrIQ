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
    public partial class PropuestaDoctForm: Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public PropuestaDoctForm()
        {
            InitializeComponent();
        }

        private void PropuestaDoctForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtPropuestasDoct, dtProfesores, dtConceptos, dtEstudiantesDoct;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtPropuestasDoct = new DataTable();
                da.Fill(dtPropuestasDoct);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                // se pide la informacion de los conceptos
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtConceptos = new DataTable();
                da.Fill(dtConceptos);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));
                dt.Columns.Add("Titulo", typeof(string));
                dt.Columns.Add("Fecha Entrega", typeof(string));
                dt.Columns.Add("Calificador 1", typeof(string));
                dt.Columns.Add("Calificador 2", typeof(string));
                dt.Columns.Add("Calificador 3", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 2", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 3", typeof(string));
                dt.Columns.Add("Fecha Entrega Correcciones", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 2", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 3", typeof(string));
                dt.Columns.Add("Fecha Sustentacion", typeof(string));
                dt.Columns.Add("Concepto Final", typeof(string));
                

                // se llena el nuevo datatable
                for (int i = 0; i < dtPropuestasDoct.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtPropuestasDoct.Rows[i][0];

                    // nombre del estudiante
                    fila[1] = dtPropuestasDoct.Rows[i][1];

                    // titulo de la propuesta
                    fila[2] = dtPropuestasDoct.Rows[i][2];

                    // fecha de entrega
                    fila[3] = dtPropuestasDoct.Rows[i][7];

                    // calificador 1
                    fila[4] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][4].ToString())[0][1];

                    // calificador 2
                    fila[5] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][5].ToString())[0][1];

                    // calificador 3
                    fila[6] = dtProfesores.Select("codigo=" + dtPropuestasDoct.Rows[i][6].ToString())[0][1];

                    // concepto 1 calificador 1
                    fila[7] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][8].ToString())[0][1];

                    // concepto 1 calificador 2
                    fila[8] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][9].ToString())[0][1];

                    // concepto 1 calificador 3
                    fila[9] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][10].ToString())[0][1];

                    // fecha de entrega de correcciones
                    fila[10] = dtPropuestasDoct.Rows[i][14];

                    // concepto 2 calificador 1
                    fila[11] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][15].ToString())[0][1];

                    // concepto 2 calificador 2
                    fila[12] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][16].ToString())[0][1];

                    // concepto 2 calificador 3
                    fila[13] = dtConceptos.Select("codigo=" + dtPropuestasDoct.Rows[i][17].ToString())[0][1];

                    // fecha sustentacion
                    fila[14] = dtPropuestasDoct.Rows[i][21];

                    // concepto sustentacion
                    fila[15] = dtPropuestasDoct.Rows[i][22];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridPropuestas.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPropuestas);

                dataGridPropuestas.Sort(dataGridPropuestas.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarPropuestasDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        private void btnMod_Click(object sender, EventArgs e)
        {

        }
    }
}
