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
    public partial class AddEstudiantesDoctForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia al MainForm padre
        /// </summary>
        public MainForm padre;

        /// <summary>
        /// Valores posibles: 'true' para agregar, 'false' para modificar
        /// </summary>
        public bool modo;

        /// <summary>
        /// Codigo del profesor a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Profesores antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtProfesores;

        /// <summary>
        /// Guarda la informacion de la tabla Condicion antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtCondicion;

        /// <summary>
        /// Guarda la informacion de la tabla Reglamentos antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtReglamentos;

        /// <summary>
        /// Guarda la informacion de la tabla Conceptos antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtConceptos;

        /// <summary>
        /// Guarda la información de la tabla EstudiantesDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtEstudiantes;

        #endregion

        public AddEstudiantesDoctForm()
        {
            InitializeComponent();
        }

        private void LlenarProfesores()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                conection.Close();

                cmbDirector.Items.Clear();
                cmbCodirector1.Items.Clear();
                cmbCodirector2.Items.Clear();

                foreach (DataRow row in dtProfesores.Rows)
                {
                    cmbDirector.Items.Add(row[1]);
                    cmbCodirector1.Items.Add(row[1]);
                    cmbCodirector2.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LlenarCondicion()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Condicion ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtCondicion = new DataTable();
                da.Fill(dtCondicion);

                conection.Close();

                cmbCondicion.Items.Clear();

                foreach (DataRow row in dtCondicion.Rows)
                {
                    cmbCondicion.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Condicion", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LlenarReglamentos()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtReglamentos = new DataTable();
                da.Fill(dtReglamentos);

                conection.Close();

                cmbReglamentos.Items.Clear();

                foreach (DataRow row in dtReglamentos.Rows)
                {
                    cmbReglamentos.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Reglamentos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LlenarConceptos()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                string query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtConceptos = new DataTable();
                da.Fill(dtConceptos);

                conection.Close();

                cmbConceptoTema.Items.Clear();

                foreach (DataRow row in dtConceptos.Rows)
                {
                    cmbConceptoTema.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Reglamentos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddEstudiantesDoctForm_Load(object sender, EventArgs e)
        {
            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los estudiantes de doctorado
                query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtEstudiantes = new DataTable();
                da.Fill(dtEstudiantes);

                conection.Close();

                // se llenan los comboBox
                LlenarConceptos();
                LlenarCondicion();
                LlenarProfesores();
                LlenarReglamentos();

                switch (modo)
                {
                    case true: // se agrega un nuevo estudiante de doctorado

                        numCod.Value = dtEstudiantes.Rows.Count + 1;
                        txtNombre.Text = "";
                        txtCorreo.Text = "";
                        numCedula.Value = 0;
                        txtCiudad.Text = "";
                        cmbCondicion.SelectedIndex = 0;
                        cmbDirector.SelectedIndex = -1;
                        cmbCodirector1.SelectedIndex = -1;
                        cmbCodirector2.SelectedIndex = -1;
                        cmbReglamentos.SelectedIndex = cmbReglamentos.Items.Count - 1;
                        txtTema.Text = "";
                        dateTema.Value = DateTime.Today;
                        cmbConceptoTema.SelectedIndex = -1;
                        txtRutaTema.Text = "";
                        cmbSolicitarQualify.SelectedIndex = 0;
                        cmbAprobarQualify.SelectedIndex = 0;

                        this.Text = "AGREGAR ESTUDIANTE DOCTORADO";
                        
                        break;

                    case false: // se modifica un estudiante de doctorado

                        DataRow[] seleccionado = dtEstudiantes.Select("codigo=" + codigo.ToString());

                        // codigo
                        numCod.Value = codigo;

                        // nombre
                        txtNombre.Text = Convert.ToString(seleccionado[0][1]);

                        // correo
                        txtCorreo.Text = Convert.ToString(seleccionado[0][2]);

                        // cedula
                        numCedula.Value = Convert.ToInt32(seleccionado[0][3]);

                        // ciudad
                        txtCiudad.Text = Convert.ToString(seleccionado[0][4]);

                        // condicion
                        cmbCondicion.SelectedIndex = Convert.ToInt32(seleccionado[0][5]) - 1;

                        // nivel
                        txtNivel.Text = Convert.ToString(seleccionado[0][6]);

                        // director
                        for (int i = 0; i < dtProfesores.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(seleccionado[0][7]) == Convert.ToInt32(dtProfesores.Rows[i][0])) cmbDirector.SelectedIndex=i;
                        }

                        // codirector 1
                        for (int i = 0; i < dtProfesores.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(seleccionado[0][8]) == Convert.ToInt32(dtProfesores.Rows[i][0])) this.cmbCodirector1.SelectedIndex = i;
                        }

                        // codirector 2
                        for (int i = 0; i < dtProfesores.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(seleccionado[0][9]) == Convert.ToInt32(dtProfesores.Rows[i][0])) this.cmbCodirector2.SelectedIndex = i;
                        }

                        // reglamento
                        cmbReglamentos.SelectedIndex = Convert.ToInt32(seleccionado[0][10]) - 1;

                        // tema
                        txtTema.Text = Convert.ToString(seleccionado[0][11]);
                        if ((!string.IsNullOrEmpty(txtTema.Text)) & (!string.IsNullOrWhiteSpace(txtTema.Text))) chkTema.Checked = true;

                        // fecha
                        dateTema.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][12]));

                        // concepto
                        cmbConceptoTema.SelectedIndex = Convert.ToInt32(seleccionado[0][13]) - 1;

                        // ruta
                        txtRutaTema.Text = Convert.ToString(seleccionado[0][14]);

                        // solicito qualify
                        if (Convert.ToString(seleccionado[0][15]) == "No") cmbSolicitarQualify.SelectedIndex = 0;
                        else cmbSolicitarQualify.SelectedIndex = 1;

                        // aprobo qualify
                        if (Convert.ToString(seleccionado[0][16]) == "No") this.cmbAprobarQualify.SelectedIndex = 0;
                        else cmbSolicitarQualify.SelectedIndex = 1;

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR ESTUDIANTE DOCTORADO";

                        break;
                }

            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void chkTema_CheckedChanged(object sender, EventArgs e)
        {
            txtTema.Enabled = chkTema.Checked;
            dateTema.Enabled = chkTema.Checked;
            cmbConceptoTema.Enabled = chkTema.Checked;
            txtRutaTema.Enabled = chkTema.Checked;
            btnRutaTema.Enabled = chkTema.Checked;
            btnVerArchivoTema.Enabled = chkTema.Checked;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        private void btnAddReglamento_Click(object sender, EventArgs e)
        {
            AddReglamentosForm agregar = new AddReglamentosForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarReglamentos();
        }
    }
}
