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
    public partial class AddPropuestaDoctForm : Form
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
        /// Codigo de la propuesta a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Estudiantes antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtEstudiantes;

        /// <summary>
        /// Guarda la informacion de la tabla Profesores antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtProfesores;

        /// <summary>
        /// Guarda la informacion de la tabla Conceptos antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtConceptos;

        /// <summary>
        /// Guarda la informacion de la tabla PropuestasDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtPropuestas;

        #endregion

        public AddPropuestaDoctForm()
        {
            InitializeComponent();
        }

        private void LlenarProfesores()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                conection.Close();

                cmbCalificador1.Items.Clear();
                cmbCalificador2.Items.Clear();
                cmbCalificador3.Items.Clear();

                foreach (DataRow row in dtProfesores.Rows)
                {
                    cmbCalificador1.Items.Add(row[1]);
                    cmbCalificador2.Items.Add(row[1]);
                    cmbCalificador3.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void LlenarConceptos()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Conceptos ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtConceptos = new DataTable();
                da.Fill(dtConceptos);

                conection.Close();

                cmbConcepto1Calificador1.Items.Clear();
                cmbConcepto1Calificador2.Items.Clear();
                cmbConcepto1Calificador3.Items.Clear();
                cmbConcepto2Calificador1.Items.Clear();
                cmbConcepto2Calificador2.Items.Clear();
                cmbConcepto2Calificador3.Items.Clear();
                cmbSustentacion.Items.Clear();

                foreach (DataRow row in dtConceptos.Rows)
                {
                    cmbConcepto1Calificador1.Items.Add(row[1]);
                    cmbConcepto1Calificador2.Items.Add(row[1]);
                    cmbConcepto1Calificador3.Items.Add(row[1]);
                    cmbConcepto2Calificador1.Items.Add(row[1]);
                    cmbConcepto2Calificador2.Items.Add(row[1]);
                    cmbConcepto2Calificador3.Items.Add(row[1]);
                    cmbSustentacion.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Conceptos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddPropuestaDoctForm_Load(object sender, EventArgs e)
        {
            label5.BackColor = label25.BackColor = label26.BackColor = Color.DarkRed;
            
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
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtPropuestas = new DataTable();
                da.Fill(dtPropuestas);

                conection.Close();

                // se llenan los combobox
                LlenarConceptos();
                LlenarEstudiantes();
                LlenarProfesores();

                switch (modo)
                {
                    case true: // se agrega una nueva propuesta de doctorado

                        numCod.Value = dtPropuestas.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtPropuesta.Text = "";
                        txtRutaPropuesta.Text = "";
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbCalificador3.SelectedIndex = -1;
                        datePropuesta.Value = DateTime.Today;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador3.SelectedIndex = -1;
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        txtRutaConcepto1Calificador3.Text = "";
                        dateCorrecciones.Value = DateTime.Today;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador3.SelectedIndex = -1;
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        txtRutaConcepto2Calificador3.Text = "";
                        dateSustentacion.Value = DateTime.Today;
                        cmbSustentacion.SelectedIndex = -1;
                        txtRutaSustentacion.Text = "";

                        this.Text = "AGREGAR PROPUESTA DE DOCTORADO";

                        break;
                    case false: // se modifica una propuesta de doctorado

                        DataRow[] seleccionado = dtPropuestas.Select("codigo=" + codigo.ToString());

                        // codigo
                        numCod.Value = codigo;

                        // estudiante
                        cmbEstudiante.SelectedIndex = Convert.ToInt32(seleccionado[0][1]) - 1;

                        // titulo
                        txtPropuesta.Text = Convert.ToString(seleccionado[0][2]);

                        // ruta
                        txtRutaPropuesta.Text = Convert.ToString(seleccionado[0][3]);

                        // calificador1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][4].ToString())) cmbCalificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][4]) - 1;

                        // calificador2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][5].ToString())) cmbCalificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][5]) - 1;

                        // calificador3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][6].ToString())) cmbCalificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][6]) - 1;

                        // entrega
                        datePropuesta.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][7]));

                        // concepto 1 calificador 1
                        cmbConcepto1Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][8]) - 1;

                        // concepto 1 calificador 2
                        cmbConcepto1Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][9]) - 1;

                        // concepto 1 calificador 3
                        cmbConcepto1Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][10]) - 1;

                        // ruta concepto 1 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][11].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][11]);

                        // ruta concepto 1 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][12].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][12]);

                        // ruta concepto 1 calificador 3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][13].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][13]);

                        // correcciones
                        dateCorrecciones.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][14]));

                        // concepto 2 calificador 1
                        cmbConcepto2Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][15]) - 1;

                        // concepto 2 calificador 2
                        cmbConcepto2Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][16]) - 1;

                        // concepto 2 calificador 3
                        cmbConcepto2Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][17]) - 1;

                        // ruta concepto 2 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][18].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][18]);

                        // ruta concepto 2 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][19].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][19]);

                        // ruta concepto 2 calificador 3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][20].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][20]);

                        // fecha sustentacion
                        dateSustentacion.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][21]));

                        // concepto sustentacion
                        cmbSustentacion.SelectedIndex = Convert.ToInt32(seleccionado[0][22]) - 1;

                        // ruta concepto sustentacion
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][23].ToString())) txtRutaSustentacion.Text = Convert.ToString(seleccionado[0][23]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR PROPUESTA DE DOCTORADO";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                busqueda = dtPropuestas.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una propuesta con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (string.IsNullOrWhiteSpace(txtPropuesta.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if(txtPropuesta.Text.Length>=255)
                {
                    MessageBox.Show("El titulo de la propuesta de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
                }
            }

            // ruta vacia
            if (string.IsNullOrWhiteSpace(txtRutaPropuesta.Text))
            {
                MessageBox.Show("No se ha seleccionado el documento de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no ha calificador1
            if (cmbCalificador1.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 1 de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // calificador 1 no es doctor
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador1.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 1 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no ha calificador2
            if (cmbCalificador2.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 2 de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // calificador 2 no es doctor
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador2.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 2 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            // no ha calificador3
            if (cmbCalificador3.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 3 de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // calificador 3 no es doctor
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador3.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 3 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if ((cmbCalificador1.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador1.SelectedIndex == cmbCalificador2.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 1 y calificador 2", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador1.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador1.SelectedIndex == cmbCalificador3.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 1 y calificador 3", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex == cmbCalificador2.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 3", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }           

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega la propuesta
                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO PropuestaDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += numCod.Value.ToString();

                        query += ", estudiante";
                        query2 += ", " + (cmbEstudiante.SelectedIndex + 1).ToString();

                        query += ", titulo";
                        query2 += ", '" + txtPropuesta.Text + "'";

                        query += ", ruta";
                        query2 += ", '" + txtRutaPropuesta.Text + "'";

                        query += ", calificador1";
                        query2 += ", " + (cmbCalificador1.SelectedIndex + 1).ToString();

                        query += ", calificador2";
                        query2 += ", " + (cmbCalificador2.SelectedIndex + 1).ToString();

                        query += ", calificador3";
                        query2 += ", " + (cmbCalificador3.SelectedIndex + 1).ToString();

                        query += ", entrega1";
                        query2 += ", '" + MainForm.Fecha2Texto(this.datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador1";
                            query2 += ", " + (cmbConcepto1Calificador1.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador1";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador2";
                            query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador2";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador3.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador3";
                            query2 += ", " + (cmbConcepto1Calificador3.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador3";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
                        {
                            query += ", rutaconcepto1calificador1";
                            query2 += ", '" + txtRutaConcepto1Calificador1.Text + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
                        {
                            query += ", rutaconcepto1calificador2";
                            query2 += ", '" + txtRutaConcepto1Calificador2.Text + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text))
                        {
                            query += ", rutaconcepto1calificador3";
                            query2 += ", '" + txtRutaConcepto1Calificador3.Text + "'";
                        }

                        query += ", correcciones";
                        query2 += ", '" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                        if (cmbConcepto2Calificador1.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador1";
                            query2 += ", " + (cmbConcepto1Calificador1.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador1";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador2.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador2";
                            query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador2";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador3.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador3";
                            query2 += ", " + (cmbConcepto2Calificador3.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador3";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text))
                        {
                            query += ", rutaconcepto2calificador1";
                            query2 += ", '" + txtRutaConcepto2Calificador1.Text + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text))
                        {
                            query += ", rutaconcepto2calificador2";
                            query2 += ", '" + txtRutaConcepto2Calificador2.Text + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text))
                        {
                            query += ", rutaconcepto2calificador3";
                            query2 += ", '" + txtRutaConcepto2Calificador3.Text + "'";
                        }

                        query += ", sustentacion";
                        query2 +=" ,'" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                        if (cmbSustentacion.SelectedIndex >= 0)
                        {
                            query += ", conceptofinal";
                            query2 += ", " + (cmbSustentacion.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", conceptofinal";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text))
                        {
                            query += ", rutaconceptofinal";
                            query2 += ", '" + txtRutaSustentacion.Text + "'";
                        }

                        query += ")";
                        query2 += ")";
                        query += query2;

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.propuestaDoctForm.PropuestaDoctForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false:

                    // se modifica la propuesta

                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE PropuestaDoct SET ";
                        query += "codigo=" + numCod.Value.ToString();
                        query += ", estudiante=" + (this.cmbEstudiante.SelectedIndex + 1).ToString();
                        query += ", titulo='" + txtPropuesta.Text + "'";
                        query += ", ruta='" + txtRutaPropuesta.Text + "'";
                        query += ", calificador1=" + (cmbCalificador1.SelectedIndex + 1).ToString();
                        query += ", calificador2=" + (cmbCalificador2.SelectedIndex + 1).ToString();
                        query += ", calificador3=" + (cmbCalificador3.SelectedIndex + 1).ToString();
                        query += ", entrega1='" + MainForm.Fecha2Texto(datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0) query += ", concepto1calificador1=" + (cmbConcepto1Calificador1.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0) query += ", concepto1calificador2=" + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador3.SelectedIndex >= 0) query += ", concepto1calificador3=" + (cmbConcepto1Calificador3.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text)) query += ", rutaconcepto1calificador1='" + txtRutaConcepto1Calificador1.Text + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text)) query += ", rutaconcepto1calificador2='" + txtRutaConcepto1Calificador2.Text + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text)) query += ", rutaconcepto1calificador3='" + txtRutaConcepto1Calificador3.Text + "'";

                        query += ", correcciones='" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                        if (cmbConcepto2Calificador1.SelectedIndex >= 0) query += ", concepto2calificador1=" + (cmbConcepto2Calificador1.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador2.SelectedIndex >= 0) query += ", concepto2calificador2=" + (cmbConcepto2Calificador2.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador3.SelectedIndex >= 0) query += ", concepto2calificador3=" + (cmbConcepto2Calificador3.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text)) query += ", rutaconcepto2calificador1='" + txtRutaConcepto2Calificador1.Text + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text)) query += ", rutaconcepto2calificador2='" + txtRutaConcepto2Calificador2.Text + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text)) query += ", rutaconcepto2calificador3='" + txtRutaConcepto2Calificador3.Text + "'";

                        query += ", sustentacion='" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                        if (cmbSustentacion.SelectedIndex >= 0) query += ", conceptofinal=" + (cmbSustentacion.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text)) query += ", rutaconceptofinal='" + txtRutaSustentacion.Text + "'";

                        query += " WHERE codigo=" + codigo.ToString();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;                    
            }
        }

        private void btnSelConcepto1Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador1.Text = open.FileName;
            }
        }

        private void btnRutaTema_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaPropuesta.Text = open.FileName;
            }
        }

        private void btnSelConcepto1Calificador2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto1Calificador3_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador3.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador1.Text = open.FileName;
            }
        }

        private void btnSelConcepto3Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto4Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador3.Text = open.FileName;
            }
        }

        private void btnSelSustentacion_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                this.txtRutaSustentacion.Text = open.FileName;
            }
        }

        private void btnVerArchivoTema_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaPropuesta.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto1Calificador1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto1Calificador1.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto1Calificador2_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto1Calificador2.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto1Calificador3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto1Calificador3.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto2Calificador1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto2Calificador1.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto2Calificador2_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto2Calificador2.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto2Calificador3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto2Calificador3.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerSustentacion_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaSustentacion.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddProfesor_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarProfesores();
        }
    }
}
