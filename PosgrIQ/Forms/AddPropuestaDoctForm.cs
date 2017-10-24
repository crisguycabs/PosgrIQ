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
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                cmbCalificador1.Items.Clear();
                cmbCalificador2.Items.Clear();
                cmbCalificador3.Items.Clear();
                cmbCalificador4.Items.Clear();

                foreach (DataRow row in dtProfesores.Rows)
                {
                    cmbCalificador1.Items.Add(row[1]);
                    cmbCalificador2.Items.Add(row[1]);
                    cmbCalificador3.Items.Add(row[1]);
                    cmbCalificador4.Items.Add(row[1]);
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
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

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
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                cmbConcepto1Calificador1.Items.Clear();
                cmbConcepto1Calificador2.Items.Clear();
                cmbConcepto1Calificador3.Items.Clear();
                cmbConcepto1Calificador4.Items.Clear();
                cmbConcepto2Calificador1.Items.Clear();
                cmbConcepto2Calificador2.Items.Clear();
                cmbConcepto2Calificador3.Items.Clear();
                cmbConcepto2Calificador4.Items.Clear();
                cmbSustentacion.Items.Clear();

                foreach (DataRow row in dtConceptos.Rows)
                {
                    cmbConcepto1Calificador1.Items.Add(row[1]);
                    cmbConcepto1Calificador2.Items.Add(row[1]);
                    cmbConcepto1Calificador3.Items.Add(row[1]);
                    cmbConcepto1Calificador4.Items.Add(row[1]);
                    cmbConcepto2Calificador1.Items.Add(row[1]);
                    cmbConcepto2Calificador2.Items.Add(row[1]);
                    cmbConcepto2Calificador3.Items.Add(row[1]);
                    cmbConcepto2Calificador4.Items.Add(row[1]);
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
                query = "SELECT * FROM PropuestaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtPropuestas = new DataTable();
                da.Fill(dtPropuestas);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se llenan los combobox
                LlenarConceptos();
                LlenarEstudiantes();
                LlenarProfesores();

                switch (modo)
                {
                    case true: // se agrega una nueva propuesta de doctorado

                        codigo = dtPropuestas.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtPropuesta.Text = "";
                        txtRutaPropuesta.Text = "";
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbCalificador3.SelectedIndex = -1;
                        cmbCalificador4.SelectedIndex = -1;
                        datePropuesta.Value = DateTime.Today;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador3.SelectedIndex = -1;
                        cmbConcepto1Calificador4.SelectedIndex = -1;
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        txtRutaConcepto1Calificador3.Text = "";
                        txtRutaConcepto1Calificador4.Text = "";
                        dateCorrecciones.Value = DateTime.Today;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador3.SelectedIndex = -1;
                        cmbConcepto2Calificador4.SelectedIndex = -1;
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        txtRutaConcepto2Calificador3.Text = "";
                        txtRutaConcepto2Calificador4.Text = "";
                        dateSustentacion.Value = DateTime.Today;
                        cmbSustentacion.SelectedIndex = -1;
                        txtRutaSustentacion.Text = "";

                        this.Text = "AGREGAR PROPUESTA DE DOCTORADO";

                        break;
                    case false: // se modifica una propuesta de doctorado

                        DataRow[] seleccionado = dtPropuestas.Select("codigo=" + codigo.ToString());

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
                        //cmbEstudiante.SelectedIndex = Convert.ToInt32(seleccionado[0][1]) - 1;

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

                        // calificador4
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][7].ToString())) cmbCalificador4.SelectedIndex = Convert.ToInt32(seleccionado[0][7]) - 1;

                        // entrega
                        datePropuesta.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][8]));

                        // concepto 1 calificador 1
                        cmbConcepto1Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][9]);

                        // concepto 1 calificador 2
                        cmbConcepto1Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][10]);

                        // concepto 1 calificador 3
                        cmbConcepto1Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][11]);

                        // concepto 1 calificador 4
                        cmbConcepto1Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][12]);

                        // ruta concepto 1 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][11].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][13]);

                        // ruta concepto 1 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][12].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][14]);

                        // ruta concepto 1 calificador 3
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][13].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][15]);

                        // ruta concepto 1 calificador 4
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][13].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][16]);

                        // correcciones
                        // se revisa si hay fecha de correcciones
                        string fecha = Convert.ToString(seleccionado[0][17]);
                        if (fecha.Length > 0)
                        {
                            // hay una fecha registrada, por tanto hay correcciones
                            chkCorrecciones.Checked = true;

                            dateCorrecciones.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][17]));

                            // concepto 2 calificador 1
                            cmbConcepto2Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][18]);

                            // concepto 2 calificador 2
                            cmbConcepto2Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][19]);

                            // concepto 2 calificador 3
                            cmbConcepto2Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][20]);

                            // concepto 2 calificador 4
                            cmbConcepto2Calificador3.SelectedIndex = Convert.ToInt32(seleccionado[0][21]);

                            // ruta concepto 2 calificador 1
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][18].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][22]);

                            // ruta concepto 2 calificador 2
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][19].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][23]);

                            // ruta concepto 2 calificador 3
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][20].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][24]);

                            // ruta concepto 2 calificador 4
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][20].ToString())) txtRutaConcepto1Calificador3.Text = Convert.ToString(seleccionado[0][25]);
                        }

                        // fecha sustentacion
                        // se revisa si hay fecha de sustentacion
                        fecha = Convert.ToString(seleccionado[0][26]);
                        if (fecha.Length > 0)
                        {
                            // hay una fecha registrada, por tanto hay sustentacion
                            chkSustentacion.Checked = true;

                            dateSustentacion.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][26]));

                            // concepto sustentacion
                            cmbSustentacion.SelectedIndex = Convert.ToInt32(seleccionado[0][27]);

                            // ruta concepto sustentacion
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][28].ToString())) txtRutaSustentacion.Text = Convert.ToString(seleccionado[0][28]);
                        }
                        else
                        {
                            // no hay una fecha registrada, por tanto no hay correcciones
                            chkSustentacion.Checked = false;
                        }

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
                busqueda = dtPropuestas.Select("codigo=" + codigo.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una propuesta con el codigo " + codigo.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (txtPropuesta.Text.Length >= 255)
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

            // se revisa si ya existe el folder para alojar la informacion del estudiante. Si no existe, se crea
            busqueda = dtEstudiantes.Select("nombre='" + cmbEstudiante.Items[cmbEstudiante.SelectedIndex] + "'");
            string codEst = Convert.ToString(busqueda[0][0]);
            string folderEst = "Est_" + codEst.ToString();
            if (!System.IO.Directory.Exists(padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + folderEst))
            {
                System.IO.Directory.CreateDirectory(padre.sourceONE + "\\Soportes\\PropuestasMaestria\\" + folderEst);
            }

            // se intenta mover el archivo de la propuesta. Si no se puede, se cancela todo            
            string destino = folderEst + "\\Propuesta_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaPropuesta.Text);
            if (!System.IO.Path.GetFileName(txtRutaPropuesta.Text).Contains("ac1017")) // no contienen la cadena => no es necesario verificar
            {
                try
                {
                    System.IO.File.Copy(txtRutaPropuesta.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destino, true);
                }
                catch
                {
                    MessageBox.Show("No se tiene acceso al archivo de la propuesta. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
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

            // calificador 4 no es doctor
            if (cmbCalificador4.SelectedIndex >= 0)
            {
                busqueda = dtProfesores.Select("codigo=" + (cmbCalificador4.SelectedIndex + 1).ToString());
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El calificador 4 no tiene titulo de doctor, no se puede asignar como calificador", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            if ((cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex == cmbCalificador2.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 3", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex == cmbCalificador2.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 2 y calificador 4", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if ((cmbCalificador4.SelectedIndex >= 0) & (cmbCalificador3.SelectedIndex >= 0) & (cmbCalificador4.SelectedIndex == cmbCalificador3.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 3 y calificador 4", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string destinoC1C1 = folderEst + "\\Concepto1Cal1_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto1Calificador1.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto1Calificador1.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador1.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC1C1, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 1. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C2 = folderEst + "\\Concepto1Cal2_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto1Calificador2.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto1Calificador2.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador2.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC1C2, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 2. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C3 = folderEst + "\\Concepto1Cal3_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto1Calificador3.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto1Calificador3.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador3.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC1C3, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 3. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC1C4 = folderEst + "\\Concepto1Cal4_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto1Calificador4.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto1Calificador4.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto1Calificador4.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC1C4, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 1 Calificador 4. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            // - 

            string destinoC2C1 = folderEst + "\\Concepto2Cal1_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto2Calificador1.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto2Calificador1.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador1.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC2C1, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 2 Calificador 1. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C2 = folderEst + "\\Concepto2Cal2_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto2Calificador2.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto2Calificador2.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador2.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC2C2, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 2 Calificador 2. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C3 = folderEst + "\\Concepto2Cal3_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto2Calificador3.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto2Calificador3.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador3.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC2C3, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 2 Calificador 3. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoC2C4 = folderEst + "\\Concepto2Cal4_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaConcepto2Calificador4.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaConcepto2Calificador4.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaConcepto2Calificador4.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoC2C4, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo Concepto 2 Calificador 4. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            string destinoActa = folderEst + "\\Acta_" + codEst.ToString() + "_ac1017" + System.IO.Path.GetExtension(txtRutaSustentacion.Text);
            if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!System.IO.Path.GetFileName(txtRutaSustentacion.Text).Contains("ac1017")) // no contienen la cadena => es necesario verificar
                {
                    try
                    {
                        System.IO.File.Copy(txtRutaSustentacion.Text, padre.sourceONE + "\\Soportes\\PropuestasDoctorado\\" + destinoActa, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo del Acta de Sustentacion. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        query = "INSERT INTO PropuestaDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += codigo.ToString();

                        query += ", estudiante";
                        query2 += ", " + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        query += ", titulo";
                        query2 += ", '" + txtPropuesta.Text + "'";

                        query += ", ruta";
                        query2 += ", 'PropuestasDoctorado\\" + destino + "'";

                        query += ", calificador1";
                        query2 += ", " + (dtProfesores.Rows[cmbCalificador1.SelectedIndex][0]).ToString();

                        query += ", calificador2";
                        query2 += ", " + (dtProfesores.Rows[cmbCalificador2.SelectedIndex][0]).ToString();

                        query += ", calificador3";
                        query2 += ", " + (dtProfesores.Rows[cmbCalificador3.SelectedIndex][0]).ToString();

                        if (cmbCalificador4.SelectedIndex >= 0)
                        {
                            query += ", calificador4";
                            query2 += ", " + (dtProfesores.Rows[cmbCalificador4.SelectedIndex][0]).ToString();
                        }
                        //else
                        //{
                        //    query += ", calificador4";
                        //    query2 += ", 0";
                        //}

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

                        if (cmbConcepto1Calificador4.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador4";
                            query2 += ", " + (cmbConcepto1Calificador4.SelectedIndex + 1).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador4";
                            query2 += ", 1";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
                        {
                            query += ", rutaconcepto1calificador1";
                            query2 += ", 'PropuestasDoctorado\\" + destinoC1C1 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
                        {
                            query += ", rutaconcepto1calificador2";
                            query2 += ", 'PropuestasDoctorado\\" + destinoC1C2 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text))
                        {
                            query += ", rutaconcepto1calificador3";
                            query2 += ", 'PropuestasDoctorado\\" + destinoC1C3 + "'";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text))
                        {
                            query += ", rutaconcepto1calificador4";
                            query2 += ", 'PropuestasDoctorado\\" + destinoC1C4 + "'";
                        }

                        if (chkCorrecciones.Checked)
                        {
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
                                query2 += ", 0";
                            }

                            if (cmbConcepto2Calificador2.SelectedIndex >= 0)
                            {
                                query += ", concepto2calificador2";
                                query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();
                            }
                            else
                            {
                                query += ", concepto2calificador2";
                                query2 += ", 0";
                            }

                            if (cmbConcepto2Calificador3.SelectedIndex >= 0)
                            {
                                query += ", concepto2calificador3";
                                query2 += ", " + (cmbConcepto2Calificador3.SelectedIndex + 1).ToString();
                            }
                            else
                            {
                                query += ", concepto2calificador3";
                                query2 += ", 0";
                            }

                            if (cmbConcepto2Calificador4.SelectedIndex >= 0)
                            {
                                query += ", concepto2calificador4";
                                query2 += ", " + (cmbConcepto2Calificador4.SelectedIndex + 1).ToString();
                            }
                            else
                            {
                                query += ", concepto2calificador4";
                                query2 += ", 0";
                            }

                            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text))
                            {
                                query += ", rutaconcepto2calificador1";
                                query2 += ", 'PropuestasDoctorado\\" + destinoC2C1 + "'";
                            }

                            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text))
                            {
                                query += ", rutaconcepto2calificador2";
                                query2 += ", 'PropuestasDoctorado\\" + destinoC2C2 + "'";
                            }

                            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text))
                            {
                                query += ", rutaconcepto2calificador3";
                                query2 += ", 'PropuestasDoctorado\\" + destinoC2C3 + "'";
                            }

                            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador4.Text))
                            {
                                query += ", rutaconcepto2calificador4";
                                query2 += ", 'PropuestasDoctorado\\" + destinoC2C4 + "'";
                            }
                        }
                        else
                        {
                            query += ", concepto2calificador1";
                            query2 += ", 0";

                            query += ", concepto2calificador2";
                            query2 += ", 0";

                            query += ", concepto2calificador3";
                            query2 += ", 0";

                            query += ", concepto2calificador4";
                            query2 += ", 0";
                        }

                        if (chkSustentacion.Checked)
                        {
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
                                query2 += ", 0";
                            }

                            if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text))
                            {
                                query += ", rutaconceptofinal";
                                query2 += ", 'PropuestasDoctorado\\" + destinoActa + "'";
                            }
                        }
                        else
                        {
                            query += ", conceptofinal";
                            query2 += ", 0";
                        }

                        query += ")";
                        query2 += ")";
                        query += query2;

                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        txtPropuesta.Text = "";
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        txtRutaConcepto1Calificador3.Text = "";
                        txtRutaConcepto1Calificador4.Text = "";
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        txtRutaConcepto2Calificador3.Text = "";
                        txtRutaConcepto2Calificador4.Text = "";
                        txtRutaPropuesta.Text = "";
                        txtRutaSustentacion.Text = "";
                        codigo++;
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbCalificador3.SelectedIndex = -1;
                        cmbCalificador4.SelectedIndex = -1;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador3.SelectedIndex = -1;
                        cmbConcepto1Calificador4.SelectedIndex = -1;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador3.SelectedIndex = -1;
                        cmbConcepto2Calificador4.SelectedIndex = -1;
                        cmbEstudiante.SelectedIndex = -1;
                        cmbSustentacion.SelectedIndex = -1;

                        //this.DialogResult = DialogResult.OK;
                        try
                        {
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
                        // se prepara la cadena SQL
                        query = "UPDATE PropuestaDoct SET ";
                        query += "codigo=" + codigo.ToString();

                        query += ", estudiante=" + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        query += ", titulo='" + txtPropuesta.Text + "'";
                        query += ", ruta='PropuestasDoctorado\\" + destino + "'";
                        query += ", calificador1=" + (cmbCalificador1.SelectedIndex + 1).ToString();
                        query += ", calificador2=" + (cmbCalificador2.SelectedIndex + 1).ToString();
                        query += ", calificador3=" + (cmbCalificador3.SelectedIndex + 1).ToString();
                        if (cmbCalificador4.SelectedIndex >= 0) query += ", calificador4=" + (cmbCalificador4.SelectedIndex + 1).ToString();
                        query += ", entrega1='" + MainForm.Fecha2Texto(datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0) query += ", concepto1calificador1=" + (cmbConcepto1Calificador1.SelectedIndex + 1).ToString();
                        else query += ", concepto1calificador1=0";

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0) query += ", concepto1calificador2=" + (cmbConcepto1Calificador2.SelectedIndex + 1).ToString();
                        else query += ", concepto1calificador2=0";

                        if (cmbConcepto1Calificador3.SelectedIndex >= 0) query += ", concepto1calificador3=" + (cmbConcepto1Calificador3.SelectedIndex + 1).ToString();
                        else query += ", concepto1calificador3=0";

                        if (cmbConcepto1Calificador4.SelectedIndex >= 0) query += ", concepto1calificador4=" + (cmbConcepto1Calificador4.SelectedIndex + 1).ToString();
                        else query += ", concepto1calificador4=0";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC1C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC1C2 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador3.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC1C3 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador4.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC1C4 + "'";

                        if (chkCorrecciones.Checked)
                        {
                            // existen correcciones, se guardan en la base de datos
                            query += ", correcciones='" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                            if (cmbConcepto2Calificador1.SelectedIndex >= 0) query += ", concepto2calificador1=" + (cmbConcepto2Calificador1.SelectedIndex + 1).ToString();
                            else query += ", concepto2calificador1=0";

                            if (cmbConcepto2Calificador2.SelectedIndex >= 0) query += ", concepto2calificador2=" + (cmbConcepto2Calificador2.SelectedIndex + 1).ToString();
                            else query += ", concepto2calificador2=0";

                            if (cmbConcepto2Calificador3.SelectedIndex >= 0) query += ", concepto2calificador3=" + (cmbConcepto2Calificador3.SelectedIndex + 1).ToString();
                            else query += ", concepto2calificador3=0";

                            if (cmbConcepto2Calificador4.SelectedIndex >= 0) query += ", concepto2calificador4=" + (cmbConcepto2Calificador4.SelectedIndex + 1).ToString();
                            else query += ", concepto2calificador4=0";
                        }
                        else
                        {
                            // no existen correcciones, por tanto se guardan ceros
                            query += ", concepto2calificador1=0";
                            query += ", concepto2calificador2=0";
                            query += ", concepto2calificador3=0";
                            query += ", concepto2calificador4=0";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC2C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC2C2 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC2C3 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador3.Text)) query += ", rutaconcepto1calificador1='PropuestasDoctorado\\" + destinoC2C4 + "'";

                        if (chkSustentacion.Checked)
                        {
                            // existe sustentacion, se guardan en la base de datos
                            query += ", sustentacion='" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                            if (cmbSustentacion.SelectedIndex >= 0) query += ", conceptofinal=" + (cmbSustentacion.SelectedIndex + 1).ToString();
                            else query += ", conceptofinal=0";
                        }
                        else query += ", conceptofinal=0";

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text)) query += ", rutaconceptofinal='PropuestasDoctorado\\" + destinoActa + "'";

                        query += " WHERE codigo=" + codigo.ToString();

                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

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
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador1.Text = open.FileName;
            }
        }

        private void btnRutaTema_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaPropuesta.Text = open.FileName;
            }
        }

        private void btnSelConcepto1Calificador2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto1Calificador3_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto1Calificador3.Text = open.FileName;
            }
        }

        private void btnSelConcepto2Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador1.Text = open.FileName;
            }
        }

        private void btnSelConcepto3Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador2.Text = open.FileName;
            }
        }

        private void btnSelConcepto4Calificador1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador3.Text = open.FileName;
            }
        }

        private void btnSelSustentacion_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                this.txtRutaSustentacion.Text = open.FileName;
            }
        }

        private void btnVerArchivoTema_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaPropuesta.Text);
        }

        private void btnVerConcepto1Calificador1_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto1Calificador1.Text);
        }

        private void btnVerConcepto1Calificador2_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto1Calificador2.Text);
        }

        private void btnVerConcepto1Calificador3_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto1Calificador3.Text);
        }

        private void btnVerConcepto2Calificador1_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto2Calificador1.Text);
        }

        private void btnVerConcepto2Calificador2_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto2Calificador2.Text);
        }

        private void btnVerConcepto2Calificador3_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto2Calificador3.Text);
        }

        private void btnVerSustentacion_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaSustentacion.Text);
        }

        private void btnAddProfesor_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarProfesores();
        }

        private void cmbEstudiante_SelectedIndexChanged(object sender, EventArgs e)
        {
            // se selecciono un estudiante, por tanto se debe buscar el TEMA en la tabla de estudiantes

            if (cmbEstudiante.SelectedIndex < 0) return; // no hacer nada en caso que se resetee el comboBox

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;
            OleDbDataAdapter da;

            // se crea la cadena de busqueda
            query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";

            try
            {
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();
                command = new OleDbCommand(query, conection);

                command.ExecuteNonQuery();

                da = new OleDbDataAdapter(command);
                dtEstudiantes = new DataTable();
                da.Fill(dtEstudiantes);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                txtPropuesta.Text = Convert.ToString(dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][11]);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkCorrecciones_CheckedChanged(object sender, EventArgs e)
        {
            dateCorrecciones.Enabled = cmbConcepto2Calificador1.Enabled = cmbConcepto2Calificador2.Enabled = cmbConcepto2Calificador3.Enabled = cmbConcepto2Calificador4.Enabled = txtRutaConcepto2Calificador1.Enabled = txtRutaConcepto2Calificador2.Enabled = txtRutaConcepto2Calificador3.Enabled = txtRutaConcepto2Calificador4.Enabled = btnSelConcepto2Calificador1.Enabled = btnSelConcepto2Calificador2.Enabled = btnSelConcepto2Calificador3.Enabled = btnSelConcepto2Calificador4.Enabled = btnVerConcepto2Calificador1.Enabled = btnVerConcepto2Calificador2.Enabled = btnVerConcepto2Calificador3.Enabled = btnVerConcepto2Calificador4.Enabled = chkCorrecciones.Checked;
        }

        private void chkSustentacion_CheckedChanged(object sender, EventArgs e)
        {
            dateSustentacion.Enabled = cmbSustentacion.Enabled = txtRutaSustentacion.Enabled = btnSelSustentacion.Enabled = btnVerSustentacion.Enabled = chkSustentacion.Checked;
        }

        private void btnSelConcepto2Calificador4_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos PDF (.pdf)|*.pdf|Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaConcepto2Calificador4.Text = open.FileName;
            }
        }

        private void btnVerConcepto2Calificador4_Click(object sender, EventArgs e)
        {
           padre.VerArchivo(txtRutaConcepto2Calificador4.Text);
        }

        private void btnVerConcepto1Calificador4_Click(object sender, EventArgs e)
        {
            padre.VerArchivo(txtRutaConcepto1Calificador4.Text);
        }
    }
}
