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
    public partial class AddMatriculaDoctForm : Form
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
        /// Guarda la informacion de la tabla MatriculaDoct antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtMatriculaDoct;

        /// <summary>
        /// Guarda la informacion de la tabla Semestres antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtSemestres;

        #endregion

        public AddMatriculaDoctForm()
        {
            InitializeComponent();
        }

        public void LlenarSemestres()
        {
            // se pide la informacion de las escuelas
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                string query = "SELECT * FROM Semestres ORDER BY codigo ASC";
                OleDbCommand command = new OleDbCommand(query, conection);
                OleDbDataAdapter da = new OleDbDataAdapter(command);

                dtSemestres = new DataTable();
                da.Fill(dtSemestres);

                conection.Close();

                this.cmbSemestre1.Items.Clear();
                this.cmbSemestre2.Items.Clear();
                this.cmbSemestre3.Items.Clear();
                this.cmbSemestre4.Items.Clear();
                this.cmbSemestre5.Items.Clear();
                this.cmbSemestre6.Items.Clear();
                this.cmbSemestre7.Items.Clear();
                this.cmbSemestre8.Items.Clear();
                this.cmbSemestre9.Items.Clear();
                this.cmbSemestre10.Items.Clear();
                this.cmbSemestre11.Items.Clear();
                this.cmbSemestre12.Items.Clear();

                foreach (DataRow row in dtSemestres.Rows)
                {
                    this.cmbSemestre1.Items.Add(row[1]);
                    this.cmbSemestre2.Items.Add(row[1]);
                    this.cmbSemestre3.Items.Add(row[1]);
                    this.cmbSemestre4.Items.Add(row[1]);
                    this.cmbSemestre5.Items.Add(row[1]);
                    this.cmbSemestre6.Items.Add(row[1]);
                    this.cmbSemestre7.Items.Add(row[1]);
                    this.cmbSemestre8.Items.Add(row[1]);
                    this.cmbSemestre9.Items.Add(row[1]);
                    this.cmbSemestre10.Items.Add(row[1]);
                    this.cmbSemestre11.Items.Add(row[1]);
                    this.cmbSemestre12.Items.Add(row[1]);
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Semestres", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void AddMatriculaDoctForm_Load(object sender, EventArgs e)
        {
            label5.BackColor = label28.BackColor = label37.BackColor = Color.DarkRed;
            
            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de la matricula de doctorado
                query = "SELECT * FROM MatriculaDoct ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtMatriculaDoct = new DataTable();
                da.Fill(dtMatriculaDoct);

                conection.Close();

                // se llenan los combobox
                LlenarSemestres();
                LlenarEstudiantes();

                switch (modo)
                {
                    case true: // se agrega una nueva matricula de doctorado

                        numCod.Value = dtMatriculaDoct.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;

                        this.Text = "AGREGAR MATRICULA DE DOCTORADO";

                        break;

                    case false: // se modifica una matricula de doctorado

                        DataRow[] seleccionado = dtMatriculaDoct.Select("codigo=" + codigo.ToString());

                        // codigo
                        numCod.Value = codigo;

                        // estudiante
                        cmbEstudiante.SelectedIndex = Convert.ToInt32(seleccionado[0][1]) - 1;

                        // semestre1
                        cmbSemestre1.SelectedIndex = Convert.ToInt32(seleccionado[0][2]) - 1;
                        numPromedio1.Value = Convert.ToDecimal(seleccionado[0][3]);
                        if (Convert.ToString(seleccionado[0][4]) == "Si") cmbBeca1.SelectedIndex = 0;
                        else cmbBeca1.SelectedIndex = 1;

                        // semestre2
                        cmbSemestre2.SelectedIndex = Convert.ToInt32(seleccionado[0][5]) - 1;
                        numPromedio2.Value = Convert.ToDecimal(seleccionado[0][6]);
                        if (Convert.ToString(seleccionado[0][7]) == "Si") cmbBeca2.SelectedIndex = 0;
                        else cmbBeca2.SelectedIndex = 1;

                        // semestre3
                        cmbSemestre3.SelectedIndex = Convert.ToInt32(seleccionado[0][8]) - 1;
                        numPromedio3.Value = Convert.ToDecimal(seleccionado[0][9]);
                        if (Convert.ToString(seleccionado[0][10]) == "Si") cmbBeca3.SelectedIndex = 0;
                        else cmbBeca3.SelectedIndex = 1;

                        // semestre4
                        cmbSemestre4.SelectedIndex = Convert.ToInt32(seleccionado[0][11]) - 1;
                        numPromedio4.Value = Convert.ToDecimal(seleccionado[0][12]);
                        if (Convert.ToString(seleccionado[0][13]) == "Si") cmbBeca4.SelectedIndex = 0;
                        else cmbBeca4.SelectedIndex = 1;

                        // semestre5
                        cmbSemestre5.SelectedIndex = Convert.ToInt32(seleccionado[0][14]) - 1;
                        numPromedio5.Value = Convert.ToDecimal(seleccionado[0][15]);
                        if (Convert.ToString(seleccionado[0][16]) == "Si") cmbBeca5.SelectedIndex = 0;
                        else cmbBeca5.SelectedIndex = 1;

                        // semestre6
                        cmbSemestre6.SelectedIndex = Convert.ToInt32(seleccionado[0][17]) - 1;
                        numPromedio6.Value = Convert.ToDecimal(seleccionado[0][18]);
                        if (Convert.ToString(seleccionado[0][19]) == "Si") cmbBeca6.SelectedIndex = 0;
                        else cmbBeca6.SelectedIndex = 1;

                        // semestre7
                        cmbSemestre7.SelectedIndex = Convert.ToInt32(seleccionado[0][20]) - 1;
                        numPromedio7.Value = Convert.ToDecimal(seleccionado[0][21]);
                        if (Convert.ToString(seleccionado[0][22]) == "Si") cmbBeca7.SelectedIndex = 0;
                        else cmbBeca7.SelectedIndex = 1;

                        // semestre8
                        cmbSemestre8.SelectedIndex = Convert.ToInt32(seleccionado[0][23]) - 1;
                        numPromedio8.Value = Convert.ToDecimal(seleccionado[0][24]);
                        if (Convert.ToString(seleccionado[0][25]) == "Si") cmbBeca8.SelectedIndex = 0;
                        else cmbBeca8.SelectedIndex = 1;

                        // semestre9
                        cmbSemestre9.SelectedIndex = Convert.ToInt32(seleccionado[0][26]) - 1;
                        numPromedio9.Value = Convert.ToDecimal(seleccionado[0][27]);

                        // semestre10
                        cmbSemestre10.SelectedIndex = Convert.ToInt32(seleccionado[0][28]) - 1;
                        numPromedio10.Value = Convert.ToDecimal(seleccionado[0][29]);

                        // semestre11
                        cmbSemestre11.SelectedIndex = Convert.ToInt32(seleccionado[0][30]) - 1;
                        numPromedio11.Value = Convert.ToDecimal(seleccionado[0][31]);

                        // semestre12
                        cmbSemestre12.SelectedIndex = Convert.ToInt32(seleccionado[0][32]) - 1;
                        numPromedio12.Value = Convert.ToDecimal(seleccionado[0][33]);

                        this.Text = "MODIFICAR MATRICULA DE DOCTORADO";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                busqueda = dtMatriculaDoct.Select("codigo=" + numCod.Value.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una matricula con el codigo " + numCod.Value.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no se ha escogido un estudiante
            if (cmbEstudiante.SelectedIndex < 0)
            {
                MessageBox.Show("No ha seleccionado un estudiante de doctorado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se comprueba que como minimo se seleccione el primer semestre
            if (cmbSemestre1.SelectedIndex < 0)
            {
                MessageBox.Show("Se debe seleccionar al menos el primer semestre matriculado", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // se comprueba que no se dupliquen los semestres seleccionados
            
            #region 1
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre2.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 2 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre3.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 3 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre4.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 4 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre5.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 5 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre6.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 6 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 2 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre1.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre1.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 1 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 2
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre3.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 3 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre4.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 4 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre5.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 5 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre6.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 6 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre2.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre2.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 2 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 3
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre4.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 4 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre5.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 5 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre6.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 6 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre3.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre3.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 3 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 4
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre5.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 5 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre6.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 6 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre4.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre4.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 4 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 5
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre6.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 6 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre5.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre5.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 5 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 6
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre7.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 7 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre6.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre6.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 6 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 7
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 8
            if ((cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex == cmbSemestre9.SelectedIndex))
            {
                MessageBox.Show("El semestre 8 y el semestre 9 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 8 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 8 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 8 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 9
            if ((cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex == cmbSemestre10.SelectedIndex))
            {
                MessageBox.Show("El semestre 9 y el semestre 10 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 9 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre9.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre9.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 9 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 10
            if ((cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex == cmbSemestre11.SelectedIndex))
            {
                MessageBox.Show("El semestre 10 y el semestre 11 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if ((cmbSemestre10.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre10.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 10 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            #region 11
            if ((cmbSemestre11.SelectedIndex >= 0) & (cmbSemestre12.SelectedIndex >= 0) & (cmbSemestre11.SelectedIndex == cmbSemestre12.SelectedIndex))
            {
                MessageBox.Show("El semestre 11 y el semestre 12 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion
            

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true: // se agrega la matricula

                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "INSERT INTO MatriculaDoct (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += numCod.Value.ToString();

                        query += ", estudiante";
                        query2 += ", " + (cmbEstudiante.SelectedIndex + 1).ToString();

                        query += ", semestre1";
                        query2 += ", " + (cmbSemestre1.SelectedIndex + 1).ToString();
                        query += ", promedio1";
                        query2 += ", " + (numPromedio1.Value).ToString();
                        query += ", beca1";
                        if (cmbBeca1.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre2";
                        query2 += ", " + (cmbSemestre2.SelectedIndex + 1).ToString();
                        query += ", promedio2";
                        query2 += ", " + (numPromedio2.Value).ToString();
                        query += ", beca2";
                        if (cmbBeca2.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre3";
                        query2 += ", " + (cmbSemestre3.SelectedIndex + 1).ToString();
                        query += ", promedio3";
                        query2 += ", " + (numPromedio3.Value).ToString();
                        query += ", beca3";
                        if (cmbBeca3.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre4";
                        query2 += ", " + (cmbSemestre4.SelectedIndex + 1).ToString();
                        query += ", promedio4";
                        query2 += ", " + (numPromedio4.Value).ToString();
                        query += ", beca4";
                        if (cmbBeca4.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre5";
                        query2 += ", " + (cmbSemestre5.SelectedIndex + 1).ToString();
                        query += ", promedio5";
                        query2 += ", " + (numPromedio5.Value).ToString();
                        query += ", beca5";
                        if (cmbBeca5.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre6";
                        query2 += ", " + (cmbSemestre6.SelectedIndex + 1).ToString();
                        query += ", promedio6";
                        query2 += ", " + (numPromedio6.Value).ToString();
                        query += ", beca6";
                        if (cmbBeca6.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre7";
                        query2 += ", " + (cmbSemestre7.SelectedIndex + 1).ToString();
                        query += ", promedio7";
                        query2 += ", " + (numPromedio7.Value).ToString();
                        query += ", beca7";
                        if (cmbBeca7.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre8";
                        query2 += ", " + (cmbSemestre8.SelectedIndex + 1).ToString();
                        query += ", promedio8";
                        query2 += ", " + (numPromedio8.Value).ToString();
                        query += ", beca8";
                        if (cmbBeca8.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        query += ", semestre9";
                        query2 += ", " + (cmbSemestre9.SelectedIndex + 1).ToString();
                        query += ", promedio9";
                        query2 += ", " + (numPromedio9.Value).ToString();

                        query += ", semestre10";
                        query2 += ", " + (cmbSemestre10.SelectedIndex + 1).ToString();
                        query += ", promedio10";
                        query2 += ", " + (numPromedio10.Value).ToString();

                        query += ", semestre11";
                        query2 += ", " + (cmbSemestre11.SelectedIndex + 1).ToString();
                        query += ", promedio11";
                        query2 += ", " + (numPromedio11.Value).ToString();

                        query += ", semestre12";
                        query2 += ", " + (cmbSemestre12.SelectedIndex + 1).ToString();
                        query += ", promedio12";
                        query2 += ", " + (numPromedio12.Value).ToString();

                        query += ")";
                        query2 += ")";
                        query += query2;

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.matriculaDoctForm.MatriculaDoctForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false: // se modifica la matricula

                    try
                    {
                        conection.Open();

                        // se prepara la cadena SQL
                        query = "UPDATE MatriculaDoct SET ";
                        query += "codigo=" + numCod.Value.ToString();
                        query += ", estudiante=" + (this.cmbEstudiante.SelectedIndex + 1).ToString();

                        query += ", semestre1=" + (this.cmbSemestre1.SelectedIndex + 1).ToString();
                        query += ", promedio1=" + (this.numPromedio1.Value).ToString();
                        if (cmbBeca1.SelectedIndex == 0) query += ", beca1='Si'";
                        else query += ", 'No'";

                        query += ", semestre2=" + (this.cmbSemestre2.SelectedIndex + 1).ToString();
                        query += ", promedio2=" + (this.numPromedio2.Value).ToString();
                        if (cmbBeca2.SelectedIndex == 0) query += ", beca2='Si'";
                        else query += ", 'No'";

                        query += ", semestre3=" + (this.cmbSemestre3.SelectedIndex + 1).ToString();
                        query += ", promedio3=" + (this.numPromedio3.Value).ToString();
                        if (cmbBeca3.SelectedIndex == 0) query += ", beca3='Si'";
                        else query += ", 'No'";

                        query += ", semestre4=" + (this.cmbSemestre4.SelectedIndex + 1).ToString();
                        query += ", promedio4=" + (this.numPromedio4.Value).ToString();
                        if (cmbBeca4.SelectedIndex == 0) query += ", beca4='Si'";
                        else query += ", 'No'";

                        query += ", semestre5=" + (this.cmbSemestre5.SelectedIndex + 1).ToString();
                        query += ", promedio5=" + (this.numPromedio5.Value).ToString();
                        if (cmbBeca5.SelectedIndex == 0) query += ", beca5='Si'";
                        else query += ", 'No'";

                        query += ", semestre6=" + (this.cmbSemestre6.SelectedIndex + 1).ToString();
                        query += ", promedio6=" + (this.numPromedio6.Value).ToString();
                        if (cmbBeca6.SelectedIndex == 0) query += ", beca6='Si'";
                        else query += ", 'No'";

                        query += ", semestre7=" + (this.cmbSemestre7.SelectedIndex + 1).ToString();
                        query += ", promedio7=" + (this.numPromedio7.Value).ToString();
                        if (cmbBeca7.SelectedIndex == 0) query += ", beca7='Si'";
                        else query += ", 'No'";

                        query += ", semestre8=" + (this.cmbSemestre8.SelectedIndex + 1).ToString();
                        query += ", promedio8=" + (this.numPromedio8.Value).ToString();
                        if (cmbBeca8.SelectedIndex == 0) query += ", beca8='Si'";
                        else query += ", 'No'";

                        query += ", semestre9=" + (this.cmbSemestre9.SelectedIndex + 1).ToString();
                        query += ", promedio9=" + (this.numPromedio9.Value).ToString();

                        query += ", semestre10=" + (this.cmbSemestre10.SelectedIndex + 1).ToString();
                        query += ", promedio10=" + (this.numPromedio10.Value).ToString();

                        query += ", semestre11=" + (this.cmbSemestre11.SelectedIndex + 1).ToString();
                        query += ", promedio11=" + (this.numPromedio11.Value).ToString();

                        query += ", semestre12=" + (this.cmbSemestre12.SelectedIndex + 1).ToString();
                        query += ", promedio12=" + (this.numPromedio12.Value).ToString();

                        query += " WHERE codigo=" + codigo.ToString();

                        command = new OleDbCommand(query, conection);

                        command.ExecuteNonQuery();

                        conection.Close();

                        this.DialogResult = DialogResult.OK;
                    }
                    catch 
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Doctorado", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
            }
        }
    }
}
