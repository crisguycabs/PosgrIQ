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
    public partial class MatriculaDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public MatriculaDoctForm()
        {
            InitializeComponent();
        }

        public void MatriculaDoctForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtMatriculaDoct, dtEstudiantesDoct, dtSemestre;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM MatriculaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtMatriculaDoct = new DataTable();
                da.Fill(dtMatriculaDoct);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                // se pide la informacion de los semestres
                query = "SELECT * FROM Semestres ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtSemestre = new DataTable();
                da.Fill(dtSemestre);

                // se crea una tabla para mostrar la data
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Estudiante", typeof(string));

                // 8 semestres normales
                for (int i = 1; i <= 8; i++)
                {
                    dt.Columns.Add("Semestre "+i.ToString(), typeof(string));
                    dt.Columns.Add("Promedio " + i.ToString(), typeof(double));
                    dt.Columns.Add("Beca " + i.ToString(), typeof(string));
                }

                // 4 semestres adicionales
                for (int i = 9; i <= 12; i++)
                {
                    dt.Columns.Add("Semestre " + i.ToString(), typeof(string));
                    dt.Columns.Add("Promedio " + i.ToString(), typeof(double));                    
                }

                // se llena el nuevo datatable
                for (int i = 0; i < dtMatriculaDoct.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();
                    int j;

                    // codigo de la matricula
                    fila[0] = dtMatriculaDoct.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesDoct.Select("codigo=" + dtMatriculaDoct.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // semestre 1
                    j=2;
                    seleccionado=dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 2
                    j = 5;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 3
                    j = 8;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 4
                    j = 11;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 5
                    j = 14;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 6
                    j = 17;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 7
                    j = 20;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 8
                    j = 23;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);

                    // semestre 9
                    j = 26;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);

                    // semestre 10
                    j = 28;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);

                    // semestre 11
                    j = 30;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);

                    // semestre 12
                    j = 32;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);                    

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridMatricula.DataSource = dt;

                conection.Close();

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridMatricula);

                dataGridMatricula.Sort(dataGridMatricula.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }                
        }
        
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarMatriculaDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddMatriculaDoctForm agregar = new AddMatriculaDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.MatriculaDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridMatricula.Rows.Count < 1) return;

            AddMatriculaDoctForm modificar = new AddMatriculaDoctForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridMatricula.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.MatriculaDoctForm_Load(sender, e);
        }
    }
}
