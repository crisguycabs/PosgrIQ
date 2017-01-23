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
    public partial class PropuestaMaesForm : Form
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

        public PropuestaMaesForm()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarPropuestasMaesForm();
        }

        public void PropuestaMaesForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtPropuestasMaes, dtProfesores, dtConceptos, dtEstudiantesMaes;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM PropuestaMaes ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtPropuestasMaes = new DataTable();
                da.Fill(dtPropuestasMaes);

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtEstudiantesMaes = new DataTable();
                da.Fill(dtEstudiantesMaes);

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
                dt.Columns.Add("Concepto 1 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 1 Calificador 2", typeof(string));
                dt.Columns.Add("Fecha Entrega Correcciones", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 1", typeof(string));
                dt.Columns.Add("Concepto 2 Calificador 2", typeof(string));
                dt.Columns.Add("Fecha Sustentacion", typeof(string));
                dt.Columns.Add("Concepto Final", typeof(string));


                // se llena el nuevo datatable
                for (int i = 0; i < dtPropuestasMaes.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtPropuestasMaes.Rows[i][0];

                    // nombre del estudiante
                    DataRow[] seleccionado = dtEstudiantesMaes.Select("codigo=" + dtPropuestasMaes.Rows[i][1]);
                    fila[1] = Convert.ToString(seleccionado[0][1]);

                    // titulo de la propuesta
                    fila[2] = dtPropuestasMaes.Rows[i][2];

                    // fecha de entrega
                    fila[3] = Convert.ToString(dtPropuestasMaes.Rows[i][6].ToString());

                    // calificador 1
                    fila[4] = dtProfesores.Select("codigo=" + dtPropuestasMaes.Rows[i][4].ToString())[0][1];

                    // calificador 2
                    fila[5] = dtProfesores.Select("codigo=" + dtPropuestasMaes.Rows[i][5].ToString())[0][1];

                    // concepto 1 calificador 1
                    fila[6] = dtConceptos.Select("codigo=" + dtPropuestasMaes.Rows[i][7].ToString())[0][1];

                    // concepto 1 calificador 2
                    fila[7] = dtConceptos.Select("codigo=" + dtPropuestasMaes.Rows[i][8].ToString())[0][1];

                    // fecha de entrega de correcciones
                    fila[8] = dtPropuestasMaes.Rows[i][11];

                    // concepto 2 calificador 1
                    fila[9] = dtConceptos.Select("codigo=" + dtPropuestasMaes.Rows[i][12].ToString())[0][1];

                    // concepto 2 calificador 2
                    fila[10] = dtConceptos.Select("codigo=" + dtPropuestasMaes.Rows[i][13].ToString())[0][1];

                    // fecha sustentacion
                    fila[11] = dtPropuestasMaes.Rows[i][16];

                    // concepto sustentacion
                    fila[12] = dtConceptos.Select("codigo=" + dtPropuestasMaes.Rows[i][17].ToString())[0][1];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridPropuestas.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridPropuestas);

                dataGridPropuestas.Sort(dataGridPropuestas.Columns[1], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridPropuestas.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtPropuestaDoct;
                OleDbDataAdapter da;

                // se pide la informacion de la propuesta de doct
                query = "SELECT * FROM PropuestaMaes WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                dtPropuestaDoct = new DataTable();
                da.Fill(dtPropuestaDoct);

                conection.Close();

                try
                {
                    switch (cmbVer.SelectedIndex)
                    {
                        case 0: // propuesta
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][3]));
                            break;
                        case 1: // concepto 1 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][9]));
                            break;
                        case 2: // concepto 1 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][10]));
                            break;
                        case 3: // concepto 2 calificador 1
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][14]));
                            break;
                        case 4: // concepto 2 calificador 2
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][15]));
                            break;
                        case 5: // sustentacion
                            System.Diagnostics.Process.Start(Convert.ToString(dtPropuestaDoct.Rows[0][18]));
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
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddPropuestaMaesForm agregar = new AddPropuestaMaesForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.PropuestaMaesForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridPropuestas.Rows.Count < 1) return;

            AddPropuestaMaesForm modificar = new AddPropuestaMaesForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridPropuestas.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.PropuestaMaesForm_Load(sender, e);
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            // se vacia la lista de indices de la busqueda
            busqueda = new List<int>();

            if (txtSearch.Text.Length < 1) return;

            // se prepara la cadena de busqueda
            string searchValue = txtSearch.Text;

            dataGridPropuestas.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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

                for (int i = 0; i < dataGridPropuestas.Rows.Count; i++)
                {
                    if (dataGridPropuestas.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                    {
                        busqueda.Add(i);
                    }
                }

                // se selecciona la primera fila encontrada, si existe una
                if (busqueda.Count > 0)
                {
                    iBusqueda = 0;
                    dataGridPropuestas.Rows[busqueda[0]].Selected = true;
                    dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
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

            dataGridPropuestas.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            // failsafe
            if (busqueda.Count <= 0) return; 
            
            iBusqueda--;
            if (iBusqueda <= 0) iBusqueda = 0;

            dataGridPropuestas.Rows[busqueda[iBusqueda]].Selected = true;
            dataGridPropuestas.FirstDisplayedScrollingRowIndex = dataGridPropuestas.SelectedRows[0].Index;
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
