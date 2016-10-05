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
    public partial class MatriculaMaesForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public MatriculaMaesForm()
        {
            InitializeComponent();
        }

        private void MatriculaMaesForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtMatriculaMaes, dtEstudiantesMaes, dtSemestre;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM MatriculaMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtMatriculaMaes = new DataTable();
                da.Fill(dtMatriculaMaes);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtEstudiantesMaes = new DataTable();
                da.Fill(dtEstudiantesMaes);

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

                // 4 semestres normales
                for (int i = 1; i <= 4; i++)
                {
                    dt.Columns.Add("Semestre " + i.ToString(), typeof(string));
                    dt.Columns.Add("Promedio " + i.ToString(), typeof(double));
                    dt.Columns.Add("Beca " + i.ToString(), typeof(string));
                }

                // 4 semestres adicionales
                for (int i = 5; i <= 8; i++)
                {
                    dt.Columns.Add("Semestre " + i.ToString(), typeof(string));
                    dt.Columns.Add("Promedio " + i.ToString(), typeof(double));
                }

                // se llena el nuevo datatable
                for (int i = 0; i < dtMatriculaMaes.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();
                    int j;

                    // codigo de la matricula
                    fila[0] = dtMatriculaMaes.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesMaes.Select("codigo=" + dtMatriculaMaes.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // semestre 1
                    j = 2;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaMaes.Rows[i][j + 2]);

                    // semestre 2
                    j = 5;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaMaes.Rows[i][j + 2]);

                    // semestre 3
                    j = 8;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaMaes.Rows[i][j + 2]);

                    // semestre 4
                    j = 11;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);
                    fila[j + 2] = Convert.ToString(dtMatriculaMaes.Rows[i][j + 2]);

                    // semestre 5
                    j = 14;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);

                    // semestre 6
                    j = 16;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);

                    // semestre 7
                    j = 18;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);

                    // semestre 8
                    j = 20;
                    seleccionado = dtSemestre.Select("codigo=" + dtMatriculaMaes.Rows[i][j]);
                    fila[j] = Convert.ToString(seleccionado[0][1]);
                    fila[j + 1] = Convert.ToDouble(dtMatriculaMaes.Rows[i][j + 1]);

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
                MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarMatriculaMaesForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddMatriculaMaesForm agregar = new AddMatriculaMaesForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.MatriculaMaesForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridMatricula.Rows.Count < 1) return;

            AddMatriculaMaesForm modificar = new AddMatriculaMaesForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridMatricula.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.MatriculaMaesForm_Load(sender, e);
        }
    }
}
