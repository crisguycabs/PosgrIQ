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
    public partial class TesisDoctForm : Form
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

        public TesisDoctForm()
        {
            InitializeComponent();
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridTesis.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtTesisDoct;
                OleDbDataAdapter da;

                // se pide la informacion de la propuesta de doct
                query = "SELECT * FROM TesisDoct WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtTesisDoct = new DataTable();
                da.Fill(dtTesisDoct);

                conection.Close();

                try
                {
                    switch (cmbVer.SelectedIndex)
                    {
                        case 0: // propuesta
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][3]));
                            break;
                        case 1: // concepto 1 calificador 1
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][10]));
                            break;
                        case 2: // concepto 1 calificador 2
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][11]));
                            break;
                        case 3: // concepto 1 calificador 3
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][12]));
                            break;
                        case 4: // concepto 1 calificador 4
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][13]));
                            break;
                        case 5: // concepto 1 calificador 5
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][14]));
                            break;
                        case 6: // concepto 2 calificador 1
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][21]));
                            break;
                        case 7: // concepto 2 calificador 2
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][22]));
                            break;
                        case 8: // concepto 2 calificador 3
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][23]));
                            break;
                        case 9: // concepto 2 calificador 4
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][24]));
                            break;
                        case 10: // concepto 2 calificador 5
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][25]));
                            break;
                        case 11: // sustentacion
                            System.Diagnostics.Process.Start(padre.sourceONE + "\\Soportes\\" + Convert.ToString(dtTesisDoct.Rows[0][33]));
                            break;
                    }
                }
                catch
                {
                    MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarTesisDoctForm();
        }

        public void TesisDoctForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtTesisDoct, dtProfesores, dtConceptos, dtEstudiantesDoct;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM TesisDoct ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtTesisDoct = new DataTable();
                da.Fill(dtTesisDoct);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                // se pide la informacion de los conceptos
                query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

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
                dt.Columns.Add("Calificador 4", typeof(string));
                dt.Columns.Add("Calificador 5", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 2", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 3", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 4", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 5", typeof(string));
                dt.Columns.Add("Fecha Entrega Correcciones", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 2", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 3", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 4", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 5", typeof(string));
                dt.Columns.Add("Fecha Sustentacion", typeof(string));
                dt.Columns.Add("Concepto Final", typeof(string));

                // se llena el nuevo datatable
                for (int i = 0; i < dtTesisDoct.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtTesisDoct.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesDoct.Select("codigo=" + dtTesisDoct.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // titulo de la propuesta
                    fila[2] = dtTesisDoct.Rows[i][2];

                    // fecha de entrega
                    fila[3] = Convert.ToString(dtTesisDoct.Rows[i][9].ToString());

                    // calificador 1
                    fila[4] = dtProfesores.Select("codigo=" + dtTesisDoct.Rows[i][4].ToString())[0][1];

                    // calificador 2
                    fila[5] = dtProfesores.Select("codigo=" + dtTesisDoct.Rows[i][5].ToString())[0][1];

                    // calificador 3
                    fila[6] = dtProfesores.Select("codigo=" + dtTesisDoct.Rows[i][6].ToString())[0][1];

                    // calificador 4
                    fila[7] = dtProfesores.Select("codigo=" + dtTesisDoct.Rows[i][7].ToString())[0][1];

                    // calificador 5
                    fila[8] = dtProfesores.Select("codigo=" + dtTesisDoct.Rows[i][8].ToString())[0][1];

                    // concepto 1 calificador 1
                    fila[9] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][10].ToString())[0][1];

                    // concepto 1 calificador 2
                    fila[10] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][11].ToString())[0][1];

                    // concepto 1 calificador 3
                    fila[11] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][12].ToString())[0][1];

                    // concepto 1 calificador 4
                    fila[12] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][13].ToString())[0][1];

                    // concepto 1 calificador 5
                    fila[13] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][14].ToString())[0][1];

                    // fecha de entrega de correcciones
                    fila[14] = dtTesisDoct.Rows[i][20];

                    // concepto 2 calificador 1
                    fila[15] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][21].ToString())[0][1];

                    // concepto 2 calificador 2
                    fila[16] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][22].ToString())[0][1];

                    // concepto 2 calificador 3
                    fila[17] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][23].ToString())[0][1];

                    // concepto 2 calificador 4
                    fila[18] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][24].ToString())[0][1];

                    // concepto 2 calificador 5
                    fila[19] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][25].ToString())[0][1];

                    // fecha sustentacion
                    fila[20] = dtTesisDoct.Rows[i][31];

                    // concepto sustentacion
                    fila[21] = dtConceptos.Select("codigo=" + dtTesisDoct.Rows[i][32].ToString())[0][1];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridTesis.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridTesis);

                dataGridTesis.Sort(dataGridTesis.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddTesisDoctForm agregar = new AddTesisDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.TesisDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridTesis.Rows.Count < 1) return;

            AddTesisDoctForm modificar = new AddTesisDoctForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridTesis.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.TesisDoctForm_Load(sender, e);
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            // se vacia la lista de indices de la busqueda
            busqueda = new List<int>();

            if (txtSearch.Text.Length < 1) return;

            // se prepara la cadena de busqueda
            string searchValue = txtSearch.Text;

            dataGridTesis.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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

                for (int i = 0; i < dataGridTesis.Rows.Count; i++)
                {
                    if (dataGridTesis.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                    {
                        busqueda.Add(i);
                    }
                }

                // se selecciona la primera fila encontrada, si existe una
                if (busqueda.Count > 0)
                {
                    iBusqueda = 0;
                    dataGridTesis.Rows[busqueda[0]].Selected = true;
                    dataGridTesis.FirstDisplayedScrollingRowIndex = dataGridTesis.SelectedRows[0].Index;
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

            dataGridTesis.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridTesis.FirstDisplayedScrollingRowIndex = dataGridTesis.SelectedRows[0].Index;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return;

            iBusqueda--;
            if (iBusqueda <= 0) iBusqueda = 0;

            dataGridTesis.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridTesis.FirstDisplayedScrollingRowIndex = dataGridTesis.SelectedRows[0].Index;
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
