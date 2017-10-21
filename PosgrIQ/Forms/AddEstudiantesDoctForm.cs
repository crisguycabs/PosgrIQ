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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                conection.Open();

                string query = "SELECT * FROM Condicion ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtCondicion = new DataTable();
                da.Fill(dtCondicion);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

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
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Reglamentos ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtReglamentos = new DataTable();
                da.Fill(dtReglamentos);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

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
            padre.CheckConflicto();

            label5.BackColor = label17.BackColor = label21.BackColor = label18.BackColor = Color.DarkRed;
            
            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
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
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

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
                        txtObservaciones.Text = "";

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

                        // director, es sensible a datos vacios
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][7].ToString()))
                        {
                            for (int i = 0; i < dtProfesores.Rows.Count; i++)
                            {
                                if (Convert.ToInt32(seleccionado[0][7]) == Convert.ToInt32(dtProfesores.Rows[i][0])) cmbDirector.SelectedIndex = i;
                            }
                        }

                        // codirector 1, es sensible a datos vacios
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][8].ToString()))
                        {
                            for (int i = 0; i < dtProfesores.Rows.Count; i++)
                            {
                                if (Convert.ToInt32(seleccionado[0][8]) == Convert.ToInt32(dtProfesores.Rows[i][0])) this.cmbCodirector1.SelectedIndex = i;
                            }
                        }

                        // codirector 2, es sensible a datos vacios
                        if (!string.IsNullOrWhiteSpace(seleccionado[0][9].ToString()))
                        {
                            for (int i = 0; i < dtProfesores.Rows.Count; i++)
                            {
                                if (Convert.ToInt32(seleccionado[0][9]) == Convert.ToInt32(dtProfesores.Rows[i][0])) this.cmbCodirector2.SelectedIndex = i;
                            }
                        }

                        // reglamento
                        cmbReglamentos.SelectedIndex = Convert.ToInt32(seleccionado[0][10]) - 1;

                        // tema
                        txtTema.Text = Convert.ToString(seleccionado[0][11]);
                        if ((!string.IsNullOrEmpty(txtTema.Text)) & (!string.IsNullOrWhiteSpace(txtTema.Text))) chkTema.Checked = true;

                        // fecha
                        dateTema.Value = MainForm.Texto2Fecha(Convert.ToString(seleccionado[0][12]));

                        // concepto, sensible a valores vacios
                        if (string.IsNullOrWhiteSpace(seleccionado[0][13].ToString())) cmbConceptoTema.SelectedIndex = -1;
                        else cmbConceptoTema.SelectedIndex = Convert.ToInt32(seleccionado[0][13]);

                        // ruta
                        txtRutaTema.Text = Convert.ToString(seleccionado[0][14]);

                        // solicito qualify
                        if (Convert.ToString(seleccionado[0][15]) == "Si") cmbSolicitarQualify.SelectedIndex = 1;
                        else cmbSolicitarQualify.SelectedIndex = 0;

                        // aprobo qualify
                        if (Convert.ToString(seleccionado[0][16]) == "Si") this.cmbAprobarQualify.SelectedIndex = 1;
                        else cmbSolicitarQualify.SelectedIndex = 0;

                        // observaciones
                        txtObservaciones.Text = Convert.ToString(seleccionado[0][17]);

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
            // se realizan algunas comprobaciones de seguridad

            // nombre vacio
            if (string.IsNullOrWhiteSpace(txtNombre.Text))
            {
                MessageBox.Show("No se ha introducido el nombre del estudiante", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataRow[] busqueda;
            
            // existe un estudiante con ese codigo. Solo revisar en modo insercion
            if (modo)
            {
                busqueda = dtEstudiantes.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un estudiante con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // existe un estudiante con esa cedula
                busqueda = dtEstudiantes.Select("cedula=" + numCedula.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un estudiante con la cedula " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // correo vacio
            if (string.IsNullOrWhiteSpace(txtCorreo.Text))
            {
                MessageBox.Show("No se ha introducido el correo del estudiante", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }            

            // ciudad vacia
            if (string.IsNullOrWhiteSpace(txtCiudad.Text))
            {
                MessageBox.Show("No se ha introducido la cidudad de la cedula de ciudadania", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no se ha seleccionado una condicion
            if (cmbCondicion.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado una condicion para el estudiante", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // no se ha seleccionado el reglamento
            if (this.cmbReglamentos.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha seleccionado un Reglamento", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // el director debe ser de doctorado o superior
            if (cmbDirector.SelectedIndex >= 0)
            {
                busqueda = dtProfesores.Select("nombre='" + cmbDirector.Items[cmbDirector.SelectedIndex] + "'");
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El director debe ser un Doctor, o superior", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // el codirector1 debe ser de doctorado o superior
            if (cmbCodirector1.SelectedIndex >= 0)
            {
                busqueda = dtProfesores.Select("nombre='" + cmbCodirector1.Items[cmbCodirector1.SelectedIndex] + "'");
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El codirector 1 debe ser un Doctor, o superior", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // el codirector2 debe ser de doctorado o superior
            if (cmbCodirector2.SelectedIndex >= 0)
            {
                busqueda = dtProfesores.Select("nombre='" + cmbCodirector2.Items[cmbCodirector2.SelectedIndex] + "'");
                if (Convert.ToInt32(busqueda[0][2]) < 2)
                {
                    MessageBox.Show("El codirector 2 debe ser un Doctor, o superior", "Falta Información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            // que Director y Codirector1 no sean el mismo
            if ((cmbDirector.SelectedIndex >= 0) & (cmbDirector.SelectedIndex == cmbCodirector1.SelectedIndex))
            {
                MessageBox.Show("Un Profesor no puede ser asignado como Director y Codirector 1 al mismo tiempo", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // que Director y Codirector2 no sean el mismo
            if ((cmbDirector.SelectedIndex >= 0) & (cmbDirector.SelectedIndex == cmbCodirector2.SelectedIndex))
            {
                MessageBox.Show("Un Profesor no puede ser asignado como Director y Codirector 2 al mismo tiempo", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // que Codirector1 y Codirector2 no sean el mismo
            if ((cmbCodirector1.SelectedIndex >= 0) & (cmbCodirector1.SelectedIndex == cmbCodirector2.SelectedIndex))
            {
                MessageBox.Show("Un Profesor no puede ser asignado como Codirector 1 y Codirector 2 al mismo tiempo", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se marco el Tema como NO ENTREGADO pero existe informacion del Tema en la ventana
            if (!chkTema.Checked & !string.IsNullOrWhiteSpace(txtTema.Text))
            {
                if (MessageBox.Show("Ha indicado que no existe aun un Tema de Doctorado pero existe un Titulo de Tema.\r\n\r\nPresione SI para desechar esta información y continuar.\r\n\r\nPresione NO para no continuar.","Error en el Tema",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No) return;
            }

            if (cmbCondicion.SelectedIndex < 0)
            {
                MessageBox.Show("No se ha indicado la condición del nuevo estudiante", "Falta información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se marco el Tema como ENTREGADO
            string destino="";
            if (chkTema.Checked)
            {
                // nombre del tema vacío
                if (string.IsNullOrWhiteSpace(txtTema.Text))
                {
                    MessageBox.Show("No ha indicado un nombre de Tema doctoral", "Falta información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // sin concepto del tema
                if (cmbConceptoTema.SelectedIndex < 0)
                {
                    MessageBox.Show("No ha indicado un concepto para el Tema doctoral", "Falta información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // no hay archivo de soporte
                if(string.IsNullOrWhiteSpace(txtRutaTema.Text))
                {
                    MessageBox.Show("No ha indicado un archivo de soporte del Tema doctoral", "Falta información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // no hay director
                if (cmbDirector.SelectedIndex < 0)
                {
                    MessageBox.Show("No se ha indicado un director para el Tema doctoral", "Falta información", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // se intenta mover el archivo del tema. Si no se puede, se cancela todo
                if (!txtRutaTema.Text.Contains("TemasDoctorado")) // no contienen la cadena => no es necesario verificar
                {
                    try
                    {
                        destino = "TemasDoctorado\\" + txtNombre.Text.Replace(" ", "") + "_Tema.pdf";
                        System.IO.File.Copy(txtRutaTema.Text, padre.sourceONE + "\\Soportes\\" + destino, true);
                    }
                    catch
                    {
                        MessageBox.Show("No se tiene acceso al archivo del tema. Verifique que el archivo no esté abierto o siendo usado", "Fallo acceso a PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    // se agrega el estudiante
                    try
                    {
                        // se prepara la cadena SQL
                        query = "INSERT INTO EstudiantesDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += numCod.Value.ToString();

                        query += ", nombre";
                        query2 += ", '" + txtNombre.Text + "'";

                        query += ", correo";
                        query2 += ", '" + txtCorreo.Text + "'";

                        query += ", cedula";
                        query2 += ", " + numCedula.Value.ToString();

                        query += ", ciudad";
                        query2 += ", '" + txtCiudad.Text + "'";

                        query += ", condicion";
                        query2 += ", " + (cmbCondicion.SelectedIndex + 1).ToString();

                        if (cmbDirector.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbDirector.Items[cmbDirector.SelectedIndex] + "'");
                            query += ", director";
                            query2 += ", '" + busqueda[0][0] + "'";
                        }

                        if (cmbCodirector1.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbCodirector1.Items[cmbCodirector1.SelectedIndex] + "'");
                            query += ", codirector1";
                            query2 += ", '" + busqueda[0][0] + "'";
                        }

                        if (cmbCodirector2.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbCodirector2.Items[cmbCodirector2.SelectedIndex] + "'");
                            query += ", codirector2";
                            query2 += ", '" + busqueda[0][0] + "'";
                        }

                        query += ", reglamento";
                        query2 += ", " + (this.cmbReglamentos.SelectedIndex + 1).ToString();

                        if (chkTema.Checked)
                        {
                            query += ", tema";
                            query2 += ", '" + txtTema.Text + "'";

                            query += ", fecha";
                            query2 += ", '" + MainForm.Fecha2Texto(dateTema.Value) + "'";

                            query += ", concepto";
                            if (cmbConceptoTema.SelectedIndex < 0) query2 += ", 1";
                            else query2 += ", " + (cmbConceptoTema.SelectedIndex).ToString();

                            destino = "TemasDoctorado\\" + txtNombre.Text.Replace(" ", "") + "_Tema.pdf";
                            query += ", ruta";
                            query2 += ", '" + destino + "'";                            
                        }
                        else
                        {
                            query += ", concepto";
                            query2 += ", 1";
                        }

                        query += ", solicitoqualify";
                        // control de valor de indice -1
                        if (cmbSolicitarQualify.SelectedIndex < 0) query2 += ", 'No'";
                        else query2 += ", '" + cmbSolicitarQualify.Items[cmbSolicitarQualify.SelectedIndex] + "'";
                        
                        query += ", aproboqualify";
                        // control de valor de indice -1
                        if (cmbAprobarQualify.SelectedIndex < 0) query2 += ", 'No'";
                        else query2 += ", '" + cmbAprobarQualify.Items[cmbAprobarQualify.SelectedIndex] + "'";

                        query += ", observaciones";
                        query2 += ", '" + txtObservaciones.Text + "'";
                        
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

                        txtCiudad.Text = "";
                        txtCorreo.Text = "";
                        txtNivel.Text = "";
                        txtNombre.Text = "";
                        txtRutaTema.Text = "";
                        txtTema.Text = "";
                        cmbAprobarQualify.SelectedIndex = -1;
                        cmbCodirector1.SelectedIndex = -1;
                        cmbCodirector2.SelectedIndex = -1;
                        cmbConceptoTema.SelectedIndex = -1;
                        cmbCondicion.SelectedIndex = -1;
                        cmbDirector.SelectedIndex = -1;
                        cmbReglamentos.SelectedIndex = -1;
                        cmbSolicitarQualify.SelectedIndex = -1;
                        numCedula.Value = 0;
                        numCod.Value = 0;
                        txtObservaciones.Text = "";

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.estudiantesDoctForm.EstudiantesDoctForm_Load(sender,e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false:

                    // se modifica el estudiante
                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE EstudiantesDoct SET ";
                        query += "codigo=" + numCod.Value.ToString();
                        query += ", nombre='" + txtNombre.Text + "'";
                        query += ", correo='" + txtCorreo.Text + "'";
                        query += ", cedula=" + numCedula.Value.ToString();
                        query += ", ciudad='" + txtCiudad.Text + "'";
                        query += ", condicion=" + (cmbCondicion.SelectedIndex + 1).ToString();
                        query += ", nivel=" + txtNivel.Text;

                        if (cmbDirector.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbDirector.Items[cmbDirector.SelectedIndex] + "'");
                            query += ", director=" + busqueda[0][0];
                        }

                        if (cmbCodirector1.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbCodirector1.Items[cmbCodirector1.SelectedIndex] + "'");
                            query += ", codirector1=" + busqueda[0][0];
                        }

                        if (cmbCodirector2.SelectedIndex >= 0)
                        {
                            busqueda = dtProfesores.Select("nombre='" + cmbCodirector2.Items[cmbCodirector2.SelectedIndex] + "'");
                            query += ", codirector2=" + busqueda[0][0];
                        }

                        query += ", reglamento=" + (cmbReglamentos.SelectedIndex + 1).ToString();
                        query += ", tema='" + txtTema.Text + "'";
                        query += ", fecha='" + MainForm.Fecha2Texto(dateTema.Value) + "'";

                        if (cmbConceptoTema.SelectedIndex >= 0) query += ", concepto=" + (cmbConceptoTema.SelectedIndex).ToString();
                        else query += ", concepto=1";

                        destino = "TemasDoctorado\\" + txtNombre.Text.Replace(" ", "") + "_Tema.pdf";
                        query += ", ruta='" + destino + "'";

                        if (cmbSolicitarQualify.SelectedIndex < 0) cmbSolicitarQualify.SelectedIndex = 0;
                        query += ", solicitoqualify='" + cmbSolicitarQualify.Items[cmbSolicitarQualify.SelectedIndex] + "'";

                        if (cmbAprobarQualify.SelectedIndex < 0) cmbAprobarQualify.SelectedIndex = 0;
                        query += ", aproboqualify='" + cmbAprobarQualify.Items[cmbAprobarQualify.SelectedIndex] + "'";

                        query += ", observaciones='" + txtObservaciones.Text + "'";

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
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
            }

        }

        private void btnAddReglamento_Click(object sender, EventArgs e)
        {
            AddReglamentosForm agregar = new AddReglamentosForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarReglamentos();
        }

        private void btnAddProfesor_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = true;

            if (agregar.ShowDialog() == DialogResult.OK) this.LlenarProfesores();
        }

        private void btnRutaTema_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Archivos DOCX (.docx)|*.docx|Archivos DOC (.doc)|*.doc|Archivos PDF (.pdf)|*.pdf";
            open.FilterIndex = 0;

            if (open.ShowDialog() == DialogResult.OK)
            {
                txtRutaTema.Text = open.FileName;
            }
        }

        private void btnVerArchivoTema_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(txtRutaTema.Text);
            }
            catch
            {
                MessageBox.Show("No se puede abrir el archivo debido a que no existe o esta dañado", "Error al intentar abrir", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
