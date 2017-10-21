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

        /// <summary>
        /// contiene los indices de las filas que coinciden con la busqueda
        /// </summary>
        public List<int> busqueda = null;

        /// <summary>
        /// indice del elemento de la busqueda que esta seleccionado
        /// </summary>
        public int iBusqueda = -1;

        #endregion

        public MatriculaDoctForm()
        {
            InitializeComponent();
        }

        public void MatriculaDoctForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtMatriculaDoct, dtEstudiantesDoct, dtSemestre;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM MatriculaDoct ORDER BY codigo ASC";
                conection.Open(); 
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtMatriculaDoct = new DataTable();
                da.Fill(dtMatriculaDoct);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                System.Threading.Thread.Sleep(1000);
                padre.CloseWaiting();

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                conection.Open(); 
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                System.Threading.Thread.Sleep(1000);
                padre.CloseWaiting();

                // se pide la informacion de los semestres
                query = "SELECT * FROM Semestres ORDER BY codigo ASC";
                conection.Open(); 
                command = new OleDbCommand(query, conection);
                da = new OleDbDataAdapter(command);
                dtSemestre = new DataTable();
                da.Fill(dtSemestre);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                padre.ShowWaiting("Espere un momento mientras PosgrIQ procesa...");
                System.Threading.Thread.Sleep(1000);
                padre.CloseWaiting();

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
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 3
                    j = 8;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 4
                    j = 11;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 5
                    j = 14;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 6
                    j = 17;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 7
                    j = 20;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 8
                    j = 23;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                        fila[j + 2] = Convert.ToString(dtMatriculaDoct.Rows[i][j + 2]);
                    }

                    // semestre 9
                    j = 26;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    }

                    // semestre 10
                    j = 28;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    }

                    // semestre 11
                    j = 30;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    }

                    // semestre 12
                    j = 32;
                    if (Convert.ToInt32(dtMatriculaDoct.Rows[i][j]) > 0)
                    {
                        seleccionado = dtSemestre.Select("codigo=" + dtMatriculaDoct.Rows[i][j]);
                        fila[j] = Convert.ToString(seleccionado[0][1]);
                        fila[j + 1] = Convert.ToDouble(dtMatriculaDoct.Rows[i][j + 1]);
                    }

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridMatricula.DataSource = dt;

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

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            // se vacia la lista de indices de la busqueda
            busqueda = new List<int>();

            if (txtSearch.Text.Length < 1) return;

            // se prepara la cadena de busqueda
            string searchValue = txtSearch.Text;

            dataGridMatricula.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                /*foreach (DataGridViewRow row in dataGridEstudiantes.Rows)
                {
                    if (row.Cells[1].Value.ToString().Contains(searchValue))
                    {
                        // row.Selected = true;
                        // dataGridEstudiantes.FirstDisplayedScrollingRowIndex = dataGridEstudiantes.SelectedRows[0].Index;
                        // break;
                        
                    }
                }*/

                for (int i = 0; i < dataGridMatricula.Rows.Count; i++)
                {
                    if (dataGridMatricula.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                    {
                        busqueda.Add(i);
                    }
                }

                // se selecciona la primera fila encontrada, si existe una
                if (busqueda.Count > 0)
                {
                    iBusqueda = 0;
                    dataGridMatricula.Rows[busqueda[0]].Selected = true;
                    dataGridMatricula.FirstDisplayedScrollingRowIndex = dataGridMatricula.SelectedRows[0].Index;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;
            
            // se aumenta en 1 el indice de busqueda seleccionado
            iBusqueda++;
            if (iBusqueda >= busqueda.Count) iBusqueda = busqueda.Count - 1;

            dataGridMatricula.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridMatricula.FirstDisplayedScrollingRowIndex = dataGridMatricula.SelectedRows[0].Index;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;

            iBusqueda--;
            if (iBusqueda <= 0) iBusqueda = 0;

            dataGridMatricula.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridMatricula.FirstDisplayedScrollingRowIndex = dataGridMatricula.SelectedRows[0].Index;
        }

        private void dataGridEstudiantes_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
