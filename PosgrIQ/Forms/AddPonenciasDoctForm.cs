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
    public partial class AddPonenciasDoctForm : Form
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
        /// Codigo de la publicacion a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla EstudiantesDoctorado antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtEstudiantes;

        /// <summary>
        /// Guarda la informacion de la tabla PublicacionesDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtPonencias;

        #endregion

        public AddPonenciasDoctForm()
        {
            InitializeComponent();
        }

        private void LlenarEstudiantes()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtEstudiantes = new DataTable();
                da.Fill(dtEstudiantes);

                conection.Close();

                this.cmbEstudiante.Items.Clear();

                foreach (DataRow row in dtEstudiantes.Rows)
                {
                    cmbEstudiante.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddPonenciasDoctForm_Load(object sender, EventArgs e)
        {
            cmbAlcance.SelectedIndex = 0;

            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de las propuestas de doctorado
                query = "SELECT * FROM PonenciasDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtPonencias = new DataTable();
                da.Fill(dtPonencias);

                conection.Close();

                // se llenan los combobox
                LlenarEstudiantes();

                switch (modo)
                {
                    case true: // se agrega una nueva publicacion de doctorado

                        numCod.Value = dtPonencias.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtTitulo.Text = "";
                        txtRevista.Text = "";
                        cmbAlcance.SelectedIndex = 0;
                        dateAceptado.Value = DateTime.Today;

                        this.Text = "AGREGAR PONENCIA DE DOCTORADO";

                        break;

                    case false: // se modifica una nueva publicacion de doctorado

                        DataRow[] seleccionado = dtPonencias.Select("codigo=" + codigo.ToString());

                        // codigo
                        numCod.Value = codigo;

                        // estudiante
                        cmbEstudiante.SelectedIndex = Convert.ToInt32(seleccionado[0][1]) - 1;

                        // titulo
                        txtTitulo.Text = Convert.ToString(seleccionado[0][2]);

                        // revista
                        txtRevista.Text = Convert.ToString(seleccionado[0][3]);

                        // fecha aceptado
                        dateAceptado.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][4]));

                        // alcance
                        if (Convert.ToString(seleccionado[0][5]) == "Nacional") cmbAlcance.SelectedIndex = 0;
                        else cmbAlcance.SelectedIndex = 1;                        

                        this.Text = "MODIFICAR PONENCIAS DE DOCTORADO";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Ponencias de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad

            DataRow[] busqueda;

            // existe una propuesta con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                busqueda = dtPonencias.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una ponencia con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no se ha escogido un estudiante
            if (cmbEstudiante.SelectedIndex < 0)
            {
                MessageBox.Show("No ha seleccionado un estudiante de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // titulo vacio
            if (string.IsNullOrWhiteSpace(txtTitulo.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre de la ponencia de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtTitulo.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo de la ponencia de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // revista vacia
            if (string.IsNullOrWhiteSpace(txtRevista.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre del evento de la ponencia de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtRevista.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo del evento de la ponencia de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // alcance vacio
            if (cmbAlcance.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha ingresado el alcance de la ponencia de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true: // se agrega la publicacion
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO PublicacionesDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += numCod.Value.ToString();

                        query += ", estudiante";
                        query2 += ", " + (cmbEstudiante.SelectedIndex + 1).ToString();

                        query += ", nombre";
                        query2 += ", '" + txtTitulo.Text + "'";

                        query += ", revista";
                        query2 += ", '" + txtRevista.Text + "'";

                        query += ", fecha";
                        query2 += ", '" + MainForm.Fecha2Texto(this.dateAceptado.Value) + "'";

                        query += ", alcance";
                        if (cmbAlcance.SelectedIndex == 1) query2 += ", 'Nacional'";
                        else query2 += ", 'Internacional'";

                        query += ")";
                        query2 += ")";
                        query += query2;

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        //this.DialogResult = DialogResult.OK;
                        padre.ponenciasDoctForm.PonenciasDoctForm_Load(sender, e);
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Ponencias de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false: // se modifica la publicacion

                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE PropuestaDoct SET ";
                        query += "codigo=" + numCod.Value.ToString();
                        query += ", estudiante=" + (this.cmbEstudiante.SelectedIndex + 1).ToString();
                        query += ", nombre='" + txtTitulo.Text + "'";
                        query += ", revista='" + txtRevista.Text + "'";
                        query += ", fecha='" + MainForm.Fecha2Texto(dateAceptado.Value) + "'";

                        if (cmbAlcance.SelectedIndex == 1) query += ", alcance='Nacional'";
                        else query += ", alcance='Internacional'";

                        query += " WHERE codigo=" + codigo.ToString();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Ponencias de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;
            }
        }
    }
}
