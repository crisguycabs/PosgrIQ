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
    public partial class AddTesisDoctForm : Form
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
        /// Guarda la informacion de la tabla TesisDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtTesis;

        #endregion

        public AddTesisDoctForm()
        {
            InitializeComponent();
        }

        private void LlenarProfesores()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                string query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                
                conection.Open();

                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                conection.Close();

                cmbCalificador1.Items.Clear();
                cmbCalificador2.Items.Clear();
                cmbCalificador3.Items.Clear();
                cmbCalificador4.Items.Clear();
                cmbCalificador5.Items.Clear();

                foreach (DataRow row in dtProfesores.Rows)
                {
                    cmbCalificador1.Items.Add(row[1]);
                    cmbCalificador2.Items.Add(row[1]);
                    cmbCalificador3.Items.Add(row[1]);
                    cmbCalificador4.Items.Add(row[1]);
                    cmbCalificador5.Items.Add(row[1]);
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
                cmbConcepto1Calificador4.Items.Clear();
                cmbConcepto1Calificador5.Items.Clear();
                cmbConcepto2Calificador1.Items.Clear();
                cmbConcepto2Calificador2.Items.Clear();
                cmbConcepto2Calificador3.Items.Clear();
                cmbConcepto2Calificador4.Items.Clear();
                cmbConcepto2Calificador5.Items.Clear();
                cmbSustentacion.Items.Clear();

                foreach (DataRow row in dtConceptos.Rows)
                {
                    cmbConcepto1Calificador1.Items.Add(row[1]);
                    cmbConcepto1Calificador2.Items.Add(row[1]);
                    cmbConcepto1Calificador3.Items.Add(row[1]);
                    cmbConcepto1Calificador4.Items.Add(row[1]);
                    cmbConcepto1Calificador5.Items.Add(row[1]);
                    cmbConcepto2Calificador1.Items.Add(row[1]);
                    cmbConcepto2Calificador2.Items.Add(row[1]);
                    cmbConcepto2Calificador3.Items.Add(row[1]);
                    cmbConcepto2Calificador4.Items.Add(row[1]);
                    cmbConcepto2Calificador5.Items.Add(row[1]);
                    cmbSustentacion.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Conceptos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void AddTesisDoctForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

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
                query = "SELECT * FROM TesisDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtTesis = new DataTable();
                da.Fill(dtTesis);

                conection.Close();

                // se llenan los combobox
                LlenarConceptos();
                LlenarEstudiantes();
                LlenarProfesores();

                switch (modo)
                {
                    case true: // se agrega una nueva propuesta de doctorado

                        codigo = dtTesis.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtTesis.Text = "";
                        txtRutaTesis.Text = "";
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbCalificador3.SelectedIndex = -1;
                        cmbCalificador4.SelectedIndex = -1;
                        cmbCalificador5.SelectedIndex = -1;
                        datePropuesta.Value = DateTime.Today;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador3.SelectedIndex = -1;
                        cmbConcepto1Calificador4.SelectedIndex = -1;
                        cmbConcepto1Calificador5.SelectedIndex = -1;
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        txtRutaConcepto1Calificador3.Text = "";
                        txtRutaConcepto1Calificador4.Text = "";
                        txtRutaConcepto1Calificador5.Text = "";
                        dateCorrecciones.Value = DateTime.Today;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador3.SelectedIndex = -1;
                        cmbConcepto2Calificador4.SelectedIndex = -1;
                        cmbConcepto2Calificador5.SelectedIndex = -1;
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        txtRutaConcepto2Calificador3.Text = "";
                        txtRutaConcepto2Calificador4.Text = "";
                        txtRutaConcepto2Calificador5.Text = "";
                        dateSustentacion.Value = DateTime.Today;
                        cmbSustentacion.SelectedIndex = -1;
                        txtRutaSustentacion.Text = "";

                        this.Text = "AGREGAR TESIS DE DOCTORADO";

                        break;
                    case false: // se modifica una propuesta de doctorado

                        DataRow[] seleccionado = dtTesis.Select("codigo=" + codigo.ToString());

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
                        txtTesis.Text = Convert.ToString(seleccionado[0][2]);

                        // ruta
                        txtRutaTesis.Text = Convert.ToString(seleccionado[0][3]);

                        // calificador1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][4].ToString())) cmbCalificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][4]) - 1;

                        // calificador2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][5].ToString())) cmbCalificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][5]) - 1;

                        // calificador3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][6].ToString())) cmbCalificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][6]) - 1;

                        // calificador4
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][7].ToString())) cmbCalificador4.SelectedIndex = Convert.ToInt32(seleccionado[0][7]) - 1;

                        // calificador5
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][8].ToString())) cmbCalificador5.SelectedIndex = Convert.ToInt32(seleccionado[0][8]) - 1;

                        // entrega
                        datePropuesta.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][9]));

                        // concepto 1 calificador 1
                        cmbConcepto1Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][10]) - 1;

                        // concepto 1 calificador 2
                        cmbConcepto1Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][11]) - 1;

                        // concepto 1 calificador 3
                        cmbConcepto1Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][12]) - 1;

                        // concepto 1 calificador 4
                        cmbConcepto1Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][13]) - 1;

                        // concepto 1 calificador 5
                        cmbConcepto1Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][14]) - 1;

                        // ruta concepto 1 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][15].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][15]);

                        // ruta concepto 1 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][16].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][16]);

                        // ruta concepto 1 calificador 3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][17].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][17]);

                        // ruta concepto 1 calificador 4
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][18].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][18]);

                        // ruta concepto 1 calificador 5
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][19].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][19]);

                        // correcciones
                        dateCorrecciones.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][20]));

                        // concepto 2 calificador 1
                        cmbConcepto2Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][21]) - 1;

                        // concepto 2 calificador 2
                        cmbConcepto2Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][22]) - 1;

                        // concepto 2 calificador 3
                        cmbConcepto2Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][23]) - 1;

                        // concepto 2 calificador 4
                        cmbConcepto2Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][24]) - 1;

                        // concepto 2 calificador 5
                        cmbConcepto2Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][25]) - 1;

                        // ruta concepto 2 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][26].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][26]);

                        // ruta concepto 2 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][27].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][27]);

                        // ruta concepto 2 calificador 3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][28].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][28]);

                        // ruta concepto 2 calificador 4
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][29].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][29]);

                        // ruta concepto 2 calificador 5
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][30].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][30]);

                        // fecha sustentacion
                        dateSustentacion.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][31]));

                        // concepto sustentacion
                        cmbSustentacion.SelectedIndex = Convert.ToInt32(seleccionado[0][32]) - 1;

                        // ruta concepto sustentacion
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][33].ToString())) txtRutaSustentacion.Text = Convert.ToString(seleccionado[0][33]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR TESIS DE DOCTORADO";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad

            DataRow[] busqueda;

            // existe una propuesta con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                busqueda = dtTesis.Select("codigo=" + codigo.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una tesis con el codigo " + codigo.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (string.IsNullOrWhiteSpace(txtTesis.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtTesis.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo de la tesis de doctorado es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // ruta vacia
            if (string.IsNullOrWhiteSpace(txtRutaTesis.Text))
            {
                MessageBox.Show("No se ha seleccionado el documento de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string destino = "";
            // se intenta mover el archivo de la propuesta. Si no se puede, se cancela todo
            if (!txtRutaTesis.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
            {
                try
                {
                    destino = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_Tesis.pdf";
                    System.IO.File.Copy(txtRutaTesis.Text, padre.sourceONE + "\\Soportes\\" + destino, true);
                }
                catch
                {
                    MessageBox.Show("No se tiene acceso al archivo de la tesis. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no ha calificador1
            if (cmbCalificador1.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 1 de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("No se ha seleccionado el calificador 2 de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("No se ha seleccionado el calificador 3 de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            // no ha calificador4
            if (cmbCalificador4.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 4 de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // calificador 4 no es doctor
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador4.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 4 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no ha calificador5
            if (cmbCalificador5.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 5 de la tesis de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // calificador 5 no es doctor
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador5.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 5 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            if ((cmbCalificador1.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador1.SelectedIndex == cmbCalificador4.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 1 y calificador 4", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador1.SelectedIndex >= 0) & (cmbCalificador5.SelectedIndex >= 0) & (cmbCalificador1.SelectedIndex == cmbCalificador5.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 1 y calificador 5", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex == cmbCalificador3.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 3", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex == cmbCalificador4.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 4", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador5.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex == cmbCalificador5.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 5", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex == cmbCalificador4.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 3 y calificador 4", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador5.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex == cmbCalificador5.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 3 y calificador 5", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador5.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex == cmbCalificador5.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 4 y calificador 5", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string destinoC1C1 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC1C1.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador1.Text.Contains("TesisDoctorado")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador1.Text, padre.sourceONE + "\\Soportes\\" + destinoC1C1, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 1. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C2 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC1C2.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador2.Text.Contains("TesisDoctorado")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador2.Text, padre.sourceONE + "\\Soportes\\" + destinoC1C2, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 2. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C3 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC1C3.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador3.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador3.Text, padre.sourceONE + "\\Soportes\\" + destinoC1C3, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 3. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C4 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC1C4.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador4.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador4.Text, padre.sourceONE + "\\Soportes\\" + destinoC1C4, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 4. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C5 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC1C5.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador5.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador5.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador5.Text, padre.sourceONE + "\\Soportes\\" + destinoC1C5, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 5. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            // - 

            string destinoC2C1 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC2C1.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador1.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador1.Text, padre.sourceONE + "\\Soportes\\" + destinoC2C1, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 2 Calificador 1. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C2 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC2C2.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador2.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador2.Text, padre.sourceONE + "\\Soportes\\" + destinoC2C2, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Conceto 1 Calificador 2. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C3 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC2C3.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador3.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador3.Text, padre.sourceONE + "\\Soportes\\" + destinoC2C3, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Conceto 2 Calificador 3. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C4 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC2C4.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador4.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador4.Text, padre.sourceONE + "\\Soportes\\" + destinoC2C4, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Conceto 2 Calificador 4. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C5 = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_TC2C5.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador5.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador5.Text, padre.sourceONE + "\\Soportes\\" + destinoC2C5, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Conceto 2 Calificador 5. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            // --

            string destinoSust = "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_Sustentacion.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text) & chkSustentacion.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaSustentacion.Text.Contains("TesisDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaSustentacion.Text, padre.sourceONE + "\\Soportes\\" + destinoSust, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo de la Sustentacion. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega la propuesta
                    try
                    {
                        // se prepara la cadena SQL
                        query = "INSERT INTO TesisDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += codigo.ToString();

                        query += ", estudiante";
                        query2 += ", " + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        query += ", titulo";
                        query2 += ", '" + txtTesis.Text + "'";

                        query += ", ruta";
                        query2 += ", '" + "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_Tesis.pdf" + "'";

                        query += ", calificador1";
                        query2 += ", " + (cmbCalificador1.SelectedIndex + 1).ToString();

                        query += ", calificador2";
                        query2 += ", " + (cmbCalificador2.SelectedIndex + 1).ToString();

                        query += ", calificador3";
                        query2 += ", " + (cmbCalificador3.SelectedIndex + 1).ToString();

                        query += ", calificador4";
                        query2 += ", " + (cmbCalificador4.SelectedIndex + 1).ToString();

                        query += ", calificador5";
                        query2 += ", " + (cmbCalificador5.SelectedIndex + 1).ToString();

                        query += ", entrega1";
                        query2 += ", '" + MainForm.Fecha2Texto(this.datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador1";
                            query2 += ", " + (cmbConcepto1Calificador1.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador1";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador2";
                            query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador2";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador3.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador3";
                            query2 += ", " + (cmbConcepto1Calificador3.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador3";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador4.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador4";
                            query2 += ", " + (cmbConcepto1Calificador4.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador4";
                            query2 += ", 1";
                        }

                        if (cmbConcepto1Calificador5.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador5";
                            query2 += ", " + (cmbConcepto1Calificador5.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador5";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
                        {
                            query += ", rutaconcepto1calificador1";
                            query2 += ", '" + destinoC1C1 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
                        {
                            query += ", rutaconcepto1calificador2";
                            query2 += ", '" + destinoC1C2 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text))
                        {
                            query += ", rutaconcepto1calificador3";
                            query2 += ", '" + destinoC1C3 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text))
                        {
                            query += ", rutaconcepto1calificador4";
                            query2 += ", '" + destinoC1C4 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador5.Text))
                        {
                            query += ", rutaconcepto1calificador5";
                            query2 += ", '" + destinoC1C5 + "'";
                        }

                        query += ", correcciones";
                        query2 += ", '" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                        if (cmbConcepto2Calificador1.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador1";
                            query2 += ", " + (cmbConcepto1Calificador1.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador1";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador2.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador2";
                            query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador2";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador3.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador3";
                            query2 += ", " + (cmbConcepto2Calificador3.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador3";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador4.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador4";
                            query2 += ", " + (cmbConcepto2Calificador4.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador4";
                            query2 += ", 1";
                        }

                        if (cmbConcepto2Calificador5.SelectedIndex >= 0)
                        {
                            query += ", concepto2calificador5";
                            query2 += ", " + (cmbConcepto2Calificador5.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto2calificador5";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text))
                        {
                            query += ", rutaconcepto2calificador1";
                            query2 += ", '" + destinoC2C1 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text))
                        {
                            query += ", rutaconcepto2calificador2";
                            query2 += ", '" + destinoC2C2 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text))
                        {
                            query += ", rutaconcepto2calificador3";
                            query2 += ", '" + destinoC2C3 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text))
                        {
                            query += ", rutaconcepto2calificador4";
                            query2 += ", '" + destinoC2C4 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador5.Text))
                        {
                            query += ", rutaconcepto2calificador5";
                            query2 += ", '" + destinoC2C5 + "'";
                        }

                        query += ", sustentacion";
                        query2 += " ,'" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                        if (cmbSustentacion.SelectedIndex >= 0)
                        {
                            query += ", conceptofinal";
                            query2 += ", " + (cmbSustentacion.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", conceptofinal";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text))
                        {
                            query += ", rutaconceptofinal";
                            query2 += ", '" + destinoSust + "'";
                        }

                        query += ")";
                        query2 += ")";
                        query += query2;

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.txtRutaConcepto1Calificador1.Text = "";
                        this.txtRutaConcepto1Calificador2.Text = "";
                        this.txtRutaConcepto1Calificador3.Text = "";
                        this.txtRutaConcepto1Calificador4.Text = "";
                        this.txtRutaConcepto1Calificador5.Text = "";
                        this.txtRutaConcepto2Calificador1.Text = "";
                        this.txtRutaConcepto2Calificador2.Text = "";
                        this.txtRutaConcepto2Calificador3.Text = "";
                        this.txtRutaConcepto2Calificador4.Text = "";
                        this.txtRutaConcepto2Calificador5.Text = "";
                        this.txtRutaSustentacion.Text = "";
                        this.txtRutaTesis.Text = "";
                        this.txtTesis.Text = "";
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbCalificador3.SelectedIndex = -1;
                        cmbCalificador4.SelectedIndex = -1;
                        cmbCalificador5.SelectedIndex = -1;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador3.SelectedIndex = -1;
                        cmbConcepto1Calificador4.SelectedIndex = -1;
                        cmbConcepto1Calificador5.SelectedIndex = -1;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador3.SelectedIndex = -1;
                        cmbConcepto2Calificador4.SelectedIndex = -1;
                        cmbConcepto2Calificador5.SelectedIndex = -1;
                        cmbEstudiante.SelectedIndex = -1;
                        cmbSustentacion.SelectedIndex = -1;
                        codigo++;
                        
                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.tesisDoctForm.TesisDoctForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false:

                    // se modifica la propuesta

                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE TesisDoct SET ";
                        query += "codigo=" + codigo.ToString();
                        query += ", estudiante=" + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();
                        query += ", titulo='" + txtTesis.Text + "'";
                        query += ", ruta='" + "TesisDoctorado\\" + cmbEstudiante.Text.Replace(" ", "") + "_Tesis.pdf" + "'";
                        query += ", calificador1=" + (cmbCalificador1.SelectedIndex + 1).ToString();
                        query += ", calificador2=" + (cmbCalificador2.SelectedIndex + 1).ToString();
                        query += ", calificador3=" + (cmbCalificador3.SelectedIndex + 1).ToString();
                        query += ", calificador4=" + (cmbCalificador4.SelectedIndex + 1).ToString();
                        query += ", calificador5=" + (cmbCalificador5.SelectedIndex + 1).ToString();
                        query += ", entrega1='" + MainForm.Fecha2Texto(datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0) query += ", concepto1calificador1=" + (cmbConcepto1Calificador1.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0) query += ", concepto1calificador2=" + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador3.SelectedIndex >= 0) query += ", concepto1calificador3=" + (cmbConcepto1Calificador3.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador4.SelectedIndex >= 0) query += ", concepto1calificador4=" + (cmbConcepto1Calificador4.SelectedIndex + 1).ToString();

                        if (cmbConcepto1Calificador5.SelectedIndex >= 0) query += ", concepto1calificador5=" + (cmbConcepto1Calificador5.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text)) query += ", rutaconcepto1calificador1='" + destinoC1C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text)) query += ", rutaconcepto1calificador2='" + destinoC1C2 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text)) query += ", rutaconcepto1calificador3='" + destinoC1C3 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text)) query += ", rutaconcepto1calificador4='" + destinoC1C4 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador5.Text)) query += ", rutaconcepto1calificador5='" + destinoC1C5 + "'";

                        query += ", correcciones='" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                        if (cmbConcepto2Calificador1.SelectedIndex >= 0) query += ", concepto2calificador1=" + (cmbConcepto2Calificador1.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador2.SelectedIndex >= 0) query += ", concepto2calificador2=" + (cmbConcepto2Calificador2.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador3.SelectedIndex >= 0) query += ", concepto2calificador3=" + (cmbConcepto2Calificador3.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador4.SelectedIndex >= 0) query += ", concepto2calificador2=" + (cmbConcepto2Calificador4.SelectedIndex + 1).ToString();

                        if (cmbConcepto2Calificador5.SelectedIndex >= 0) query += ", concepto2calificador3=" + (cmbConcepto2Calificador5.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text)) query += ", rutaconcepto2calificador1='" + destinoC2C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text)) query += ", rutaconcepto2calificador2='" + destinoC2C2 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text)) query += ", rutaconcepto2calificador3='" + destinoC2C3 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text)) query += ", rutaconcepto2calificador4='" + destinoC2C4 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador5.Text)) query += ", rutaconcepto2calificador5='" + destinoC2C5 + "'";

                        query += ", sustentacion='" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                        if (cmbSustentacion.SelectedIndex >= 0) query += ", conceptofinal=" + (cmbSustentacion.SelectedIndex + 1).ToString();

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text)) query += ", rutaconceptofinal='" + txtRutaSustentacion.Text + "'";

                        query += " WHERE codigo=" + codigo.ToString();

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Tesis de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
            }
        }

        private void btnAddProfesor_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarProfesores();
        }

        private void btnRutaTesis_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaTesis.Text = open.FileName;
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

        private void btnSelConcepto1Calificador4_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador4.Text = open.FileName;
            }
        }

        private void btnSelConcepto1Calificador5_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador5.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador3_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador3.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador4_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador4.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador5_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador5.Text = open.FileName;
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

        private void btnVerConcepto1Calificador4_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto1Calificador4.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto1Calificador5_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto1Calificador5.Text);
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

        private void btnVerConcepto2Calificador4_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto2Calificador4.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerConcepto2Calificador5_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaConcepto2Calificador5.Text);
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

        private void btnVerArchivoTesis_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaTesis.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkCorrecciones_CheckedChanged(object sender, EventArgs e)
        {
            dateCorrecciones.Enabled = chkCorrecciones.Checked;

            cmbConcepto2Calificador1.Enabled = chkCorrecciones.Checked;
            cmbConcepto2Calificador2.Enabled = chkCorrecciones.Checked;
            cmbConcepto2Calificador3.Enabled = chkCorrecciones.Checked;
            cmbConcepto2Calificador4.Enabled = chkCorrecciones.Checked;
            cmbConcepto2Calificador5.Enabled = chkCorrecciones.Checked;

            txtRutaConcepto2Calificador1.Enabled = chkCorrecciones.Checked;
            txtRutaConcepto2Calificador2.Enabled = chkCorrecciones.Checked;
            txtRutaConcepto2Calificador3.Enabled = chkCorrecciones.Checked;
            txtRutaConcepto2Calificador4.Enabled = chkCorrecciones.Checked;
            txtRutaConcepto2Calificador5.Enabled = chkCorrecciones.Checked;

            btnSelConcepto2Calificador1.Enabled = chkCorrecciones.Checked;
            btnSelConcepto2Calificador2.Enabled = chkCorrecciones.Checked;
            btnSelConcepto2Calificador3.Enabled = chkCorrecciones.Checked;
            btnSelConcepto2Calificador4.Enabled = chkCorrecciones.Checked;
            btnSelConcepto2Calificador5.Enabled = chkCorrecciones.Checked;

            btnVerConcepto2Calificador1.Enabled = chkCorrecciones.Checked;
            btnVerConcepto2Calificador2.Enabled = chkCorrecciones.Checked;
            btnVerConcepto2Calificador3.Enabled = chkCorrecciones.Checked;
            btnVerConcepto2Calificador4.Enabled = chkCorrecciones.Checked;
            btnVerConcepto2Calificador5.Enabled = chkCorrecciones.Checked;
        }

        private void chkSustentacion_CheckedChanged(object sender, EventArgs e)
        {
            dateSustentacion.Enabled = chkCorrecciones.Checked;
            cmbSustentacion.Enabled = chkCorrecciones.Checked;
            txtRutaSustentacion.Enabled = chkCorrecciones.Checked;
            btnSelSustentacion.Enabled = chkCorrecciones.Checked;
            btnVerSustentacion.Enabled = chkCorrecciones.Checked;
        }

        private void cmbEstudiante_SelectedIndexChanged(object sender, EventArgs e)
        {
            // se selecciono un estudiante, por tanto se debe buscar la PROPUESTA en la tabla de propuestas

            if (cmbEstudiante.SelectedIndex < 0) return; // no hacer nada en caso que se resetee el comboBox

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;
            OleDbDataAdapter da;

            // se crea la cadena de busqueda
            query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";

            try
            {
                conection.Open();
                command = new OleDbCommand(query, conection);

                command.ExecuteNonQuery();

                da = new OleDbDataAdapter(command);
                DataTable dtPropuestas = new DataTable();
                da.Fill(dtPropuestas);               

                conection.Close();

                // se busca si existe alguna propuesta a nombre del estudiante seleccionado
                DataRow[] seleccion = dtPropuestas.Select("estudiante=" + this.dtEstudiantes.Rows[this.cmbEstudiante.SelectedIndex][0].ToString());
                if (seleccion.Length > 0)
                {
                    txtTesis.Text = Convert.ToString(seleccion[0][2]);
                }
                else txtTesis.Text = "";
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }        
    }
}
