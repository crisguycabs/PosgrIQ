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
    public partial class EstudiantesDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public EstudiantesDoctForm()
        {
            InitializeComponent();
        }

        public void EstudiantesDoctForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtEstudiantesDoct, dtProfesores, dtCondicion, dtReglamentos, dtConceptos;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorad
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

                // se pide la informacion de la condicion del estudiante
                query = "SELECT * FROM Condicion ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtCondicion = new DataTable();
                da.Fill(dtCondicion);

                // se pide la informacion de los reglamentos de posgrado
                query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtReglamentos = new DataTable();
                da.Fill(dtReglamentos);

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
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Correo", typeof(string));
                dt.Columns.Add("Cedula", typeof(string));
                dt.Columns.Add("Ciudad", typeof(string));
                dt.Columns.Add("Condicion", typeof(string));
                dt.Columns.Add("Nivel", typeof(int));
                dt.Columns.Add("Director", typeof(string));
                dt.Columns.Add("Codirector1", typeof(string));
                dt.Columns.Add("Codirector2", typeof(string));
                dt.Columns.Add("Reglamento", typeof(string));
                dt.Columns.Add("Tema", typeof(string));
                dt.Columns.Add("Fecha Entrega Tema", typeof(string));
                dt.Columns.Add("Concepto Tema", typeof(string));
                dt.Columns.Add("Solicito Qualify", typeof(string));
                dt.Columns.Add("Aprobo Qualify", typeof(string));
                dt.Columns.Add("Observaciones", typeof(string));

                // se llena el nuevo datatable
                for (int i = 0; i < dtEstudiantesDoct.Rows.Count; i++)
                {   
                    DataRow fila = dt.NewRow();

                    // codigo del estudiante
                    fila[0] = dtEstudiantesDoct.Rows[i][0];

                    // nombre del estudiante
                    fila[1] = dtEstudiantesDoct.Rows[i][1];

                    // correo del estudiante
                    fila[2] = dtEstudiantesDoct.Rows[i][2];

                    // cedula del estudiante
                    fila[3] = dtEstudiantesDoct.Rows[i][3];

                    // ciudad de la cedula del estudiante
                    fila[4] = dtEstudiantesDoct.Rows[i][4];

                    // condicion
                    fila[5] = dtCondicion.Select("codigo=" + dtEstudiantesDoct.Rows[i][5].ToString())[0][1];

                    // nivel del estudiante
                    fila[6] = dtEstudiantesDoct.Rows[i][6];

                    // director. es posible que no haya informacion
                    if (string.IsNullOrWhiteSpace(dtEstudiantesDoct.Rows[i][7].ToString())) fila[7] = "";
                    else fila[7] = dtProfesores.Select("codigo=" + dtEstudiantesDoct.Rows[i][7].ToString())[0][1];

                    // codirector1. es posible que no haya informacion
                    if (string.IsNullOrWhiteSpace(dtEstudiantesDoct.Rows[i][8].ToString())) fila[8] = "";
                    else fila[8] = dtProfesores.Select("codigo=" + dtEstudiantesDoct.Rows[i][8].ToString())[0][1];

                    // codirector2. es posible que no haya informacion
                    if (string.IsNullOrWhiteSpace(dtEstudiantesDoct.Rows[i][9].ToString())) fila[9] = "";
                    else fila[9] = dtProfesores.Select("codigo=" + dtEstudiantesDoct.Rows[i][9].ToString())[0][1];

                    // reglamento
                    fila[10] = dtReglamentos.Select("codigo=" + dtEstudiantesDoct.Rows[i][10].ToString())[0][1];

                    // tema
                    fila[11] = dtEstudiantesDoct.Rows[i][11];

                    // fecha de entrega del tema
                    fila[12] = dtEstudiantesDoct.Rows[i][12];

                    // concepto del tema, sensible a valores vacios
                    if (string.IsNullOrWhiteSpace(dtEstudiantesDoct.Rows[i][13].ToString())) fila[13] = "";
                    else fila[13] = dtConceptos.Select("codigo=" + dtEstudiantesDoct.Rows[i][13].ToString())[0][1];

                    // solicito qualify
                    fila[14] = dtEstudiantesDoct.Rows[i][15];

                    // aprobo qualify
                    fila[15] = dtEstudiantesDoct.Rows[i][16];

                    // observaciones
                    fila[16] = dtEstudiantesDoct.Rows[i][17];

                    dt.Rows.Add(fila);
                }

                // se enlaza el datatable con el datagrid
                dataGridEstudiantes.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridEstudiantes);

                dataGridEstudiantes.Sort(dataGridEstudiantes.Columns[0], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarEstudiantesDoctForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddEstudiantesDoctForm agregar = new AddEstudiantesDoctForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.EstudiantesDoctForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridEstudiantes.Rows.Count < 1) return;
            
            AddEstudiantesDoctForm modificar = new AddEstudiantesDoctForm();
            modificar.padre = this.padre;
            modificar.modo = false;
            modificar.codigo = Convert.ToInt32(dataGridEstudiantes.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.EstudiantesDoctForm_Load(sender, e);
        }

        private void btnVerTema_Click(object sender, EventArgs e)
        {
            int codigo = Convert.ToInt32(dataGridEstudiantes.SelectedRows[0].Cells[0].Value);
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dtEstudiantesDoct;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorad
                query = "SELECT * FROM EstudiantesDoct WHERE codigo=" + codigo.ToString() + " ORDER BY codigo ASC";

                conection.Open();

                command = new OleDbCommand(query, conection);

                conection.Close();

                da = new OleDbDataAdapter(command);
                dtEstudiantesDoct = new DataTable();
                da.Fill(dtEstudiantesDoct);                

                try
                {
                    System.Diagnostics.Process.Start(Convert.ToString(dtEstudiantesDoct.Rows[0][14]));
                }
                catch
                {
                    MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
