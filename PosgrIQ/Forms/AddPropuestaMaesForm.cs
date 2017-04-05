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
    public partial class AddPropuestaMaesForm : Form
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
        /// Guarda la informacion de la tabla PropuestasMaes antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtPropuestas;

        #endregion

        public AddPropuestaMaesForm()
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
                
                foreach (DataRow row in dtProfesores.Rows)
                {
                    cmbCalificador1.Items.Add(row[1]);
                    cmbCalificador2.Items.Add(row[1]);
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

                string query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
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
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                cmbConcepto2Calificador1.Items.Clear();
                cmbConcepto2Calificador2.Items.Clear();
                cmbSustentacion.Items.Clear();

                foreach (DataRow row in dtConceptos.Rows)
                {
                    cmbConcepto1Calificador1.Items.Add(row[1]);
                    cmbConcepto1Calificador2.Items.Add(row[1]);
                    cmbConcepto2Calificador1.Items.Add(row[1]);
                    cmbConcepto2Calificador2.Items.Add(row[1]);
                    cmbSustentacion.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Conceptos", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddPropuestaMaesForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            label5.BackColor = label26.BackColor = label25.BackColor = Color.DarkRed;
            
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
                query = "SELECT * FROM PropuestaMaes ORDER BY codigo ASC";
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

                        codigo = dtPropuestas.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;
                        txtPropuesta.Text = "";
                        txtRutaPropuesta.Text = "";
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        datePropuesta.Value = DateTime.Today;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        dateCorrecciones.Value = DateTime.Today;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        dateSustentacion.Value = DateTime.Today;
                        cmbSustentacion.SelectedIndex = -1;
                        txtRutaSustentacion.Text = "";

                        this.Text = "AGREGAR PROPUESTA DE MAESTRIA";

                        break;
                    case false: // se modifica una propuesta de doctorado

                        DataRow[] seleccionado = dtPropuestas.Select("codigo=" + codigo.ToString());

                        // estudiante
                        // se selecciona el indice en el cmbEstudiante segun el codigo de estudiante en la propuesta
                        string est = seleccionado[0][1].ToString();
                        for (int i = 0; i < dtEstudiantes.Rows.Count;i++ )
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

                        // entrega
                        datePropuesta.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][6]));

                        // concepto 1 calificador 1
                        cmbConcepto1Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][7]);

                        // concepto 1 calificador 2
                        cmbConcepto1Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][8]);

                        // ruta concepto 1 calificador 1
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][9].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][9]);

                        // ruta concepto 1 calificador 2
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][10].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][10]);

                        // correcciones
                        // se revisa si hay fecha de correcciones
                        string fecha = Convert.ToString(seleccionado[0][11]);
                        if(fecha.Length>0)
                        {
                            // hay una fecha registrada, por tanto hay correcciones
                            chkCorrecciones.Checked = true;

                            dateCorrecciones.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][11]));

                            // concepto 2 calificador 1
                            cmbConcepto2Calificador1.SelectedIndex = Convert.ToInt32(seleccionado[0][12]);

                            // concepto 2 calificador 2
                            cmbConcepto2Calificador2.SelectedIndex = Convert.ToInt32(seleccionado[0][13]);

                            // ruta concepto 2 calificador 1
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][14].ToString())) txtRutaConcepto1Calificador1.Text = Convert.ToString(seleccionado[0][14]);

                            // ruta concepto 2 calificador 2
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][15].ToString())) txtRutaConcepto1Calificador2.Text = Convert.ToString(seleccionado[0][15]);
                        }
                        else
                        {
                            // no hay una fecha registrada, por tanto no hay correcciones
                            chkCorrecciones.Checked = false;
                        }
                        
                        // fecha sustentacion
                        // se revisa si hay fecha de sustentacion
                        fecha = Convert.ToString(seleccionado[0][16]);
                        if(fecha.Length>0)
                        {
                            // hay una fecha registrada, por tanto hay sustentacion
                            chkSustentacion.Checked = true;

                            dateSustentacion.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][16]));

                            // concepto sustentacion
                            cmbSustentacion.SelectedIndex = Convert.ToInt32(seleccionado[0][17]);

                            // ruta concepto sustentacion
                            if (!string.IsNullOrWhiteSpace(seleccionado[0][18].ToString())) txtRutaSustentacion.Text = Convert.ToString(seleccionado[0][18]);
                        }
                        else
                        {
                            // no hay una fecha registrada, por tanto no hay correcciones
                            chkSustentacion.Checked = false;
                        }
                        
                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR PROPUESTA DE MAESTRIA";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnRutaPropuesta_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaPropuesta.Text = open.FileName;
            }
        }

        private void btnVerArchivoPropuesta_Click(object sender, EventArgs e)
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

        private void btnAddProfesor_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = 2;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarProfesores();
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
                MessageBox.Show("No ha seleccionado un estudiante de maestria", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // titulo vacio
            if (string.IsNullOrWhiteSpace(txtPropuesta.Text))
            {
                MessageBox.Show("No se ha ingresado el nombre de la propuesta de maestria", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // titulo demasiado largo
                if (txtPropuesta.Text.Length >= 255)
                {
                    MessageBox.Show("El titulo de la propuesta de maestria es demasiado.\r\n\r\nMaximo 255 caractereres", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // ruta vacia
            if (string.IsNullOrWhiteSpace(txtRutaPropuesta.Text))
            {
                MessageBox.Show("No se ha seleccionado el documento de la propuesta de maestria", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string destino="";
            // se intenta mover el archivo de la propuesta. Si no se puede, se cancela todo
            if (!txtRutaPropuesta.Text.Contains("PropuestasMaestria")) // no contienen la cadena => no es necesario verificar
            {
                try
                {
                    destino = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_Propuesta.pdf";
                    System.IO.File.Copy(txtRutaPropuesta.Text, padre.sourceONE + "\\Soportes\\" + destino, true);
                }
                catch
                {
                    MessageBox.Show("No se tiene acceso al archivo de la propuesta. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no hay calificador1
            if (cmbCalificador1.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 1 de la propuesta de maestria", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // no hay calificador2
            if (cmbCalificador2.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado el calificador 2 de la propuesta de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if ((cmbCalificador1.SelectedIndex >= 0) & (cmbCalificador2.SelectedIndex >= 0) & (cmbCalificador1.SelectedIndex == cmbCalificador2.SelectedIndex))
            {
                MessageBox.Show("Se asigno el mismo profesor a calificador 1 y calificador 2", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string destinoC1C1 = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_PC1C1.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador1.Text.Contains("PropuestasMaestria")) // no contienen la cadena => es necesario verificar
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

            string destinoC1C2 = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_PC1C2.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text))
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto1Calificador2.Text.Contains("PropuestasMaestria")) // no contienen la cadena => es necesario verificar
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

            string destinoC2C1 = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_PC2C1.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador1.Text.Contains("PropuestasMaestria")) // no contienen la cadena => no es necesario verificar
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

            string destinoC2C2 = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_PC2C2.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text) & chkCorrecciones.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaConcepto2Calificador2.Text.Contains("PropuestasMaestria")) // no contienen la cadena => no es necesario verificar
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

            string destinoSust = "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_Sustentacion.pdf";
            if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text) & chkSustentacion.Checked)
            {
                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaSustentacion.Text.Contains("PropuestasMaestria")) // no contienen la cadena => no es necesario verificar
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
                        query = "INSERT INTO PropuestaMaes (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += codigo.ToString();

                        query += ", estudiante";
                        query2 += ", " + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        query += ", titulo";
                        query2 += ", '" + txtPropuesta.Text + "'";

                        query += ", ruta";
                        query2 += ", '" + "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_Propuesta.pdf" + "'";

                        query += ", calificador1";
                        query2 += ", " + (dtProfesores.Rows[cmbCalificador1.SelectedIndex][0]).ToString();

                        query += ", calificador2";
                        query2 += ", " + (dtProfesores.Rows[cmbCalificador2.SelectedIndex][0]).ToString();

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
                            query2 += ", 0";
                        }

                        if (cmbConcepto1Calificador2.SelectedIndex >= 0)
                        {
                            query += ", concepto1calificador2";
                            query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex).ToString();
                        }
                        else
                        {
                            query += ", concepto1calificador2";
                            query2 += ", 0";
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

                        if (chkCorrecciones.Checked)
                        {
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
                                query2 += ", 0";
                            }

                            if (cmbConcepto2Calificador2.SelectedIndex >= 0)
                            {
                                query += ", concepto2calificador2";
                                query2 += ", " + (cmbConcepto1Calificador2.SelectedIndex).ToString();
                            }
                            else
                            {
                                query += ", concepto2calificador2";
                                query2 += ", 0";
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
                        }
                        else
                        {
                            query += ", concepto2calificador1";
                            query2 += ", 0";

                            query += ", concepto2calificador2";
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
                                query2 += ", '" + destinoSust + "'";
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

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        txtPropuesta.Text = "";
                        txtRutaConcepto1Calificador1.Text = "";
                        txtRutaConcepto1Calificador2.Text = "";
                        txtRutaConcepto2Calificador1.Text = "";
                        txtRutaConcepto2Calificador2.Text = "";
                        txtRutaPropuesta.Text = "";
                        txtRutaSustentacion.Text = "";
                        codigo++;
                        cmbCalificador1.SelectedIndex = -1;
                        cmbCalificador2.SelectedIndex = -1;
                        cmbConcepto1Calificador1.SelectedIndex = -1;
                        cmbConcepto1Calificador2.SelectedIndex = -1;
                        cmbConcepto2Calificador1.SelectedIndex = -1;
                        cmbConcepto2Calificador2.SelectedIndex = -1;
                        cmbEstudiante.SelectedIndex = -1;
                        cmbSustentacion.SelectedIndex = -1;

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.propuestaMaesForm.PropuestaMaesForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false:

                    // se modifica la propuesta

                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE PropuestaMaes SET ";
                        query += "codigo=" + codigo.ToString();

                        query += ", estudiante=" + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();
                        
                        query += ", titulo='" + txtPropuesta.Text + "'";
                        query += ", ruta='" + "PropuestasMaestria\\" + cmbEstudiante.Text.Replace(" ", "") + "_Propuesta.pdf" + "'";
                        query += ", calificador1=" + (cmbCalificador1.SelectedIndex + 1).ToString();
                        query += ", calificador2=" + (cmbCalificador2.SelectedIndex + 1).ToString();
                        query += ", entrega1='" + MainForm.Fecha2Texto(datePropuesta.Value) + "'";

                        if (cmbConcepto1Calificador1.SelectedIndex >= 0) query += ", concepto1calificador1=" + (cmbConcepto1Calificador1.SelectedIndex).ToString();
                        else query += ", concepto1calificador1=0";
                        
                        if (cmbConcepto1Calificador1.SelectedIndex >= 0) query += ", concepto1calificador2=" + (cmbConcepto1Calificador2.SelectedIndex).ToString();
                        else query += ", concepto1calificador2=0";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador1.Text)) query += ", rutaconcepto1calificador1='" + destinoC1C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto1Calificador2.Text)) query += ", rutaconcepto1calificador2='" + destinoC1C2 + "'";

                        if (chkCorrecciones.Checked)
                        {
                            // existen correcciones, se guardan en la base de datos
                            query += ", correcciones='" + MainForm.Fecha2Texto(dateCorrecciones.Value) + "'";

                            if (cmbConcepto2Calificador1.SelectedIndex >= 0) query += ", concepto2calificador1=" + (cmbConcepto2Calificador1.SelectedIndex).ToString();
                            else query += ", concepto2calificador1=0";

                            if (cmbConcepto2Calificador2.SelectedIndex >= 0) query += ", concepto2calificador2=" + (cmbConcepto2Calificador2.SelectedIndex).ToString();
                            else query += ", concepto2calificador2=0";
                        }
                        else 
                        {
                            // no existen correcciones, por tanto se guardan ceros
                            query += ", concepto2calificador1=0";
                            query += ", concepto2calificador2=0";
                        }

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador1.Text)) query += ", rutaconcepto2calificador1='" + destinoC2C1 + "'";

                        if (!string.IsNullOrWhiteSpace(txtRutaConcepto2Calificador2.Text)) query += ", rutaconcepto2calificador2='" + destinoC2C2 + "'";

                        if(chkSustentacion.Checked)
                        {
                            // existe sustentacion, se guardan en la base de datos
                            query += ", sustentacion='" + MainForm.Fecha2Texto(dateSustentacion.Value) + "'";

                            if (cmbSustentacion.SelectedIndex >= 0) query += ", conceptofinal=" + (cmbSustentacion.SelectedIndex).ToString();
                            else query += ", conceptofinal=0";
                        }
                        else query += ", conceptofinal=0";

                        if (!string.IsNullOrWhiteSpace(txtRutaSustentacion.Text)) query += ", rutaconceptofinal='" + destinoSust + "'";

                        query += " WHERE codigo=" + codigo.ToString();

                        conection.Open();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Propuestas de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
            }
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
            query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";

            try
            {
                conection.Open();
                command = new OleDbCommand(query, conection);

                command.ExecuteNonQuery();

                da = new OleDbDataAdapter(command);
                dtEstudiantes = new DataTable();
                da.Fill(dtEstudiantes);

                conection.Close();

                txtPropuesta.Text = Convert.ToString(dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][11]);                
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkCorrecciones_CheckedChanged(object sender, EventArgs e)
        {
            dateCorrecciones.Enabled = cmbConcepto2Calificador1.Enabled = cmbConcepto2Calificador2.Enabled = txtRutaConcepto2Calificador1.Enabled = txtRutaConcepto2Calificador2.Enabled = btnSelConcepto2Calificador1.Enabled = btnSelConcepto2Calificador2.Enabled = chkCorrecciones.Checked;
        }

        private void chkSustentacion_CheckedChanged(object sender, EventArgs e)
        {
            dateSustentacion.Enabled = cmbSustentacion.Enabled = txtRutaSustentacion.Enabled = btnSelSustentacion.Enabled = btnVerSustentacion.Enabled = chkSustentacion.Checked;
        }
    }
}
