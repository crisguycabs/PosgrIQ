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
    public partial class AddPublicacionesDoctForm : Form
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
        public DataTable dtPublicaciones;        

        #endregion

        public AddPublicacionesDoctForm()
        {
            InitializeComponent();
        }

        private void AddPublicacionesDoctForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

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
                query = "SELECT * FROM PublicacionesDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtPublicaciones = new DataTable();
                da.Fill(dtPublicaciones);

                conection.Close();

                // se llenan los combobox
                LlenarEstudiantes();

                switch (modo)
                {
                    case true: // se agrega una nueva publicacion de doctorado

                        numCod.Value = dtPublicaciones.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtTitulo.Text = "";
                        txtRevista.Text = "";
                        cmbAlcance.SelectedIndex = 0;
                        txtCategoria.Text = "";
                        dateAceptado.Value = DateTime.Today;

                        this.Text = "AGREGAR PUBLICACION DE DOCTORADO";

                        break;

                    case false: // se modifica una nueva publicacion de doctorado

                        DataRow[] seleccionado = dtPublicaciones.Select("codigo=" + codigo.ToString());

                        // codigo
                        numCod.Value = codigo;

                        // estudiante
                        // se selecciona el indice en el cmbEstudiante segun el codigo de estudiante en la propuesta
                        string est = seleccionado[0][1].ToString();
                        for (int i = 0; i < dtEstudiantes.Rows.Count; i++)
                        {
                            if (dtEstudiantes.Rows[i][0].ToString() == est)
                            {
                                cmbEstudiante.SelectedIndex = i;
                                break;
                            }
                        }

                        // titulo
                        txtTitulo.Text = Convert.ToString(seleccionado[0][2]);

                        // revista
                        txtRevista.Text = Convert.ToString(seleccionado[0][3]);

                        // fecha aceptado
                        dateAceptado.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][4]));

                        // alcance
                        if (Convert.ToString(seleccionado[0][5]) == "Nacional") cmbAlcance.SelectedIndex = 0;
                        else cmbAlcance.SelectedIndex = 1;

                        // categoria
                        txtCategoria.Text = Convert.ToString(seleccionado[0][6]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR PUBLICACION DE DOCTORADO";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Publicaciones de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                busqueda = dtPublicaciones.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una publicacion con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("No se ha ingresado el nombre de la publicacion de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtTitulo.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo de la publicacion de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // revista vacia
            if (string.IsNullOrWhiteSpace(txtRevista.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre de la revista de la publicacion de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtRevista.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo de la revista de la publicacion de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // categoria vacia
            if (string.IsNullOrWhiteSpace(txtCategoria.Text))
            {
                MessageBox.Show("No se ha ingresado la categoria de la publicacion de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }           

            // alcance vacio
            if (cmbAlcance.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha ingresado el alcance de la publicacion de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true: // se agrega la publicacion
                    try
                    {
                        // se prepara la cadena SQL
                        query = "INSERT INTO PublicacionesDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += numCod.Value.ToString();

                        query += ", estudiante";
                        query2 += ", " + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        query += ", nombre";
                        query2 += ", '" + txtTitulo.Text + "'";

                        query += ", revista";
                        query2 += ", '" + txtRevista.Text + "'";

                        query += ", fecha";
                        query2 += ", '" + MainForm.Fecha2Texto(this.dateAceptado.Value) + "'";

                        query += ", alcance";
                        if (cmbAlcance.SelectedIndex == 1) query2 += ", 'Nacional'";
                        else query2 += ", 'Internacional'";

                        query += ", categoria";
                        query2 += ", '" + txtCategoria.Text + "'";

                        query += ")";
                        query2 += ")";
                        query += query2;

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        txtCategoria.Text = "";
                        txtRevista.Text = "";
                        txtTitulo.Text = "";
                        cmbAlcance.SelectedIndex = -1;
                        cmbEstudiante.SelectedIndex = -1;
                        numCod.Value = 0;

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.publicacionesDoctForm.PublicacionesDoctForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Publicaciones de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false: // se modifica la publicacion

                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE PropuestaDoct SET ";
                        query += "codigo=" + numCod.Value.ToString();
                        query += ", estudiante=" + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();
                        query += ", nombre='" + txtTitulo.Text + "'";
                        query += ", revista='" + txtRevista.Text + "'";
                        query += ", fecha='" + MainForm.Fecha2Texto(dateAceptado.Value) + "'";

                        if (cmbAlcance.SelectedIndex == 1) query += ", alcance='Nacional'";
                        else query += ", alcance='Internacional'";

                        query += ", categoria='" + txtCategoria.Text + "'";

                        query += " WHERE codigo=" + codigo.ToString();

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Publicaciones de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;
            }
        }
    }
}
