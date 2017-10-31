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
    public partial class AddMatriculaMaesForm : Form
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
        public DataTable dtMatriculaMaes;

        /// <summary>
        /// Guarda la informacion de la tabla Semestres antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtSemestres;

        #endregion

        public AddMatriculaMaesForm()
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
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                this.cmbSemestre1.Items.Clear();
                this.cmbSemestre2.Items.Clear();
                this.cmbSemestre3.Items.Clear();
                this.cmbSemestre4.Items.Clear();
                this.cmbSemestre5.Items.Clear();
                this.cmbSemestre6.Items.Clear();
                this.cmbSemestre7.Items.Clear();
                this.cmbSemestre8.Items.Clear();
                

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

                string query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
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
                MessageBox.Show("No se puede acceder a la base de datos, tabla Estudiantes de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddMatriculaMaesForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

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
                query = "SELECT * FROM MatriculaMaes ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
                this.dtMatriculaMaes = new DataTable();
                da.Fill(dtMatriculaMaes);

                conection.Close();
                conection.Dispose();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                {
                    System.Threading.Thread.Sleep(100);
                }

                // se llenan los combobox
                LlenarSemestres();
                LlenarEstudiantes();

                switch (modo)
                {
                    case true: // se agrega una nueva matricula de doctorado

                        codigo = dtMatriculaMaes.Rows.Count + 1;
                        cmbEstudiante.SelectedIndex = -1;

                        this.Text = "AGREGAR MATRICULA DE MAESTRIAS";

                        break;

                    case false: // se modifica una matricula de doctorado

                        DataRow[] seleccionado = dtMatriculaMaes.Select("codigo=" + codigo.ToString());

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

                        // semestre1
                        cmbSemestre1.SelectedIndex = Convert.ToInt32(seleccionado[0][2]) - 1;
                        numPromedio1.Text = Convert.ToString(seleccionado[0][3]);
                        if (Convert.ToString(seleccionado[0][4]) == "Si") cmbBeca1.SelectedIndex = 0;
                        else cmbBeca1.SelectedIndex = 1;

                        // semestre2
                        cmbSemestre2.SelectedIndex = Convert.ToInt32(seleccionado[0][5]) - 1;
                        numPromedio2.Text = Convert.ToString(seleccionado[0][6]);
                        if (Convert.ToString(seleccionado[0][7]) == "Si") cmbBeca2.SelectedIndex = 0;
                        else cmbBeca2.SelectedIndex = 1;

                        // semestre3
                        cmbSemestre3.SelectedIndex = Convert.ToInt32(seleccionado[0][8]) - 1;
                        numPromedio3.Text = Convert.ToString(seleccionado[0][9]);
                        if (Convert.ToString(seleccionado[0][10]) == "Si") cmbBeca3.SelectedIndex = 0;
                        else cmbBeca3.SelectedIndex = 1;

                        // semestre4
                        cmbSemestre4.SelectedIndex = Convert.ToInt32(seleccionado[0][11]) - 1;
                        numPromedio4.Text = Convert.ToString(seleccionado[0][12]);
                        if (Convert.ToString(seleccionado[0][13]) == "Si") cmbBeca4.SelectedIndex = 0;
                        else cmbBeca4.SelectedIndex = 1;

                        // semestre5
                        cmbSemestre5.SelectedIndex = Convert.ToInt32(seleccionado[0][14]) - 1;
                        numPromedio5.Text = Convert.ToString(seleccionado[0][15]);

                        // semestre6
                        cmbSemestre6.SelectedIndex = Convert.ToInt32(seleccionado[0][16]) - 1;
                        numPromedio6.Text = Convert.ToString(seleccionado[0][17]);

                        // semestre7
                        cmbSemestre7.SelectedIndex = Convert.ToInt32(seleccionado[0][18]) - 1;
                        numPromedio7.Text = Convert.ToString(seleccionado[0][19]);

                        // semestre8
                        cmbSemestre8.SelectedIndex = Convert.ToInt32(seleccionado[0][20]) - 1;
                        numPromedio8.Text = Convert.ToString(seleccionado[0][21]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR MATRICULA DE MAESTRIA";

                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                busqueda = dtMatriculaMaes.Select("codigo=" + codigo.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe una matricula con el codigo " + codigo.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // no se ha escogido un estudiante
            if (cmbEstudiante.SelectedIndex < 0)
            {
                MessageBox.Show("No ha seleccionado un estudiante de maestria", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            #endregion
            #region 7
            if ((cmbSemestre7.SelectedIndex >= 0) & (cmbSemestre8.SelectedIndex >= 0) & (cmbSemestre7.SelectedIndex == cmbSemestre8.SelectedIndex))
            {
                MessageBox.Show("El semestre 7 y el semestre 8 estan duplicados.\r\n\r\nPor favor, seleccione semestres diferentes", "Falta informacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion

            // se verifican los contenidos de los promedios para cada semestre
            # region

            if (!MainForm.TestParser(numPromedio1.Text))
            {
                MessageBox.Show("El promedio del semestre 1 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio2.Text))
            {
                MessageBox.Show("El promedio del semestre 2 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio3.Text))
            {
                MessageBox.Show("El promedio del semestre 3 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio4.Text))
            {
                MessageBox.Show("El promedio del semestre 4 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio5.Text))
            {
                MessageBox.Show("El promedio del semestre 5 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio6.Text))
            {
                MessageBox.Show("El promedio del semestre 6 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio7.Text))
            {
                MessageBox.Show("El promedio del semestre 7 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!MainForm.TestParser(numPromedio8.Text))
            {
                MessageBox.Show("El promedio del semestre 8 no es válido. Debe escribirse de la forma #.##", "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            #endregion

            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query, query2;
            OleDbCommand command;

            switch (modo)
            {
                case true: // se agrega la matricula

                    try
                    {
                        // se prepara la cadena SQL
                        query = "INSERT INTO MatriculaMaes (";
                        query2 = " VALUES (";

                        query += "codigo";
                        query2 += codigo.ToString();

                        query += ", estudiante";
                        query2 += ", " + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        // el primer semestre es obligatorio
                        query += ", semestre1";
                        query2 += ", " + (cmbSemestre1.SelectedIndex + 1).ToString();
                        query += ", promedio1";
                        query2 += ", '" + (numPromedio1.Text).ToString() + "'";
                        query += ", beca1";
                        if (cmbBeca1.SelectedIndex == 0) query2 += ", 'Si'";
                        else query2 += ", 'No'";

                        // los demas semestres no son obligatorios
                        if (cmbSemestre2.SelectedIndex >= 0)
                        {
                            query += ", semestre2";
                            query2 += ", " + (cmbSemestre2.SelectedIndex + 1).ToString();
                            query += ", promedio2";
                            query2 += ", '" + (numPromedio2.Text).ToString() + "'";
                            query += ", beca2";
                            if (cmbBeca2.SelectedIndex == 0) query2 += ", 'Si'";
                            else query2 += ", 'No'";
                        }

                        if (cmbSemestre3.SelectedIndex >= 0)
                        {
                            query += ", semestre3";
                            query2 += ", " + (cmbSemestre3.SelectedIndex + 1).ToString();
                            query += ", promedio3";
                            query2 += ", '" + (numPromedio3.Text).ToString() + "'";
                            query += ", beca3";
                            if (cmbBeca3.SelectedIndex == 0) query2 += ", 'Si'";
                            else query2 += ", 'No'";
                        }

                        if (cmbSemestre4.SelectedIndex >= 0)
                        {
                            query += ", semestre4";
                            query2 += ", " + (cmbSemestre4.SelectedIndex + 1).ToString();
                            query += ", promedio4";
                            query2 += ", '" + (numPromedio4.Text).ToString() + "'";
                            query += ", beca4";
                            if (cmbBeca4.SelectedIndex == 0) query2 += ", 'Si'";
                            else query2 += ", 'No'";
                        }

                        if (cmbSemestre5.SelectedIndex >= 0)
                        {
                            query += ", semestre5";
                            query2 += ", " + (cmbSemestre5.SelectedIndex + 1).ToString();
                            query += ", promedio5";
                            query2 += ", '" + (numPromedio5.Text).ToString() + "'";
                        }

                        if (cmbSemestre6.SelectedIndex >= 0)
                        {
                            query += ", semestre6";
                            query2 += ", " + (cmbSemestre6.SelectedIndex + 1).ToString();
                            query += ", promedio6";
                            query2 += ", '" + (numPromedio6.Text).ToString() + "'";
                        }

                        if (cmbSemestre7.SelectedIndex >= 0)
                        {
                            query += ", semestre7";
                            query2 += ", " + (cmbSemestre7.SelectedIndex + 1).ToString();
                            query += ", promedio7";
                            query2 += ", '" + (numPromedio7.Text).ToString() + "'";
                        }

                        if (cmbSemestre8.SelectedIndex >= 0)
                        {
                            query += ", semestre8";
                            query2 += ", " + (cmbSemestre8.SelectedIndex + 1).ToString();
                            query += ", promedio8";
                            query2 += ", '" + (numPromedio8.Text+ "'").ToString() + "'";
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

                        codigo++;
                        this.numPromedio1.Text = "0.00";
                        this.numPromedio2.Text = "0.00";
                        this.numPromedio3.Text = "0.00";
                        this.numPromedio4.Text = "0.00";
                        numPromedio5.Text = "0.00";
                        numPromedio6.Text = "0.00";
                        numPromedio7.Text = "0.00";
                        numPromedio8.Text = "0.00";          
                        cmbBeca1.SelectedIndex = -1;
                        cmbBeca2.SelectedIndex = -1;
                        cmbBeca3.SelectedIndex = -1;
                        cmbBeca4.SelectedIndex = -1;
                        cmbEstudiante.SelectedIndex = -1;
                        cmbSemestre1.SelectedIndex = -1;
                        cmbSemestre2.SelectedIndex = -1;
                        cmbSemestre3.SelectedIndex = -1;
                        cmbSemestre4.SelectedIndex = -1;
                        cmbSemestre5.SelectedIndex = -1;
                        cmbSemestre6.SelectedIndex = -1;
                        cmbSemestre7.SelectedIndex = -1;
                        cmbSemestre8.SelectedIndex = -1;
                        
                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.matriculaMaesForm.MatriculaMaesForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false: // se modifica la matricula

                    try
                    {
                        // se prepara la cadena SQL
                        query = "UPDATE MatriculaMaes SET ";
                        query += "codigo=" + codigo.ToString();
                        query += ", estudiante=" + (dtEstudiantes.Rows[cmbEstudiante.SelectedIndex][0]).ToString();

                        // el primer semestre es obligatorio
                        query += ", semestre1=" + (this.cmbSemestre1.SelectedIndex + 1).ToString();
                        query += ", promedio1='" + (this.numPromedio1.Text).ToString() + "'";
                        if (cmbBeca1.SelectedIndex == 0) query += ", beca1='Si'";
                        else query += ", beca1='No'";

                        // los demas semestres no son obligatorios
                        if (cmbSemestre2.SelectedIndex >= 0)
                        {
                            query += ", semestre2=" + (this.cmbSemestre2.SelectedIndex + 1).ToString();
                            query += ", promedio2='" + (this.numPromedio2.Text).ToString() + "'";
                            if (cmbBeca2.SelectedIndex == 0) query += ", beca2='Si'";
                            else query += ", beca2='No'";
                        }

                        if (cmbSemestre3.SelectedIndex >= 0)
                        {
                            query += ", semestre3=" + (this.cmbSemestre3.SelectedIndex + 1).ToString();
                            query += ", promedio3='" + (this.numPromedio3.Text).ToString() + "'";
                            if (cmbBeca3.SelectedIndex == 0) query += ", beca3='Si'";
                            else query += ", beca3='No'";
                        }

                        if (cmbSemestre4.SelectedIndex >= 0)
                        {
                            query += ", semestre4=" + (this.cmbSemestre4.SelectedIndex + 1).ToString();
                            query += ", promedio4='" + (this.numPromedio4.Text).ToString() + "'";
                            if (cmbBeca4.SelectedIndex == 0) query += ", beca4='Si'";
                            else query += ", beca4='No'";
                        }

                        if (cmbSemestre5.SelectedIndex >= 0)
                        {
                            query += ", semestre5=" + (this.cmbSemestre5.SelectedIndex + 1).ToString();
                            query += ", promedio5='" + (this.numPromedio5.Text).ToString() + "'";
                        }

                        if (cmbSemestre6.SelectedIndex >= 0)
                        {
                            query += ", semestre6=" + (this.cmbSemestre6.SelectedIndex + 1).ToString();
                            query += ", promedio6='" + (this.numPromedio6.Text).ToString() + "'";
                        }
                        
                        if (cmbSemestre7.SelectedIndex >= 0)
                        {
                            query += ", semestre7=" + (this.cmbSemestre7.SelectedIndex + 1).ToString();
                            query += ", promedio7='" + (this.numPromedio7.Text).ToString() + "'";
                        }

                        if (cmbSemestre8.SelectedIndex >= 0)
                        {
                            query += ", semestre8=" + (this.cmbSemestre8.SelectedIndex + 1).ToString();
                            query += ", promedio8='" + (this.numPromedio8.Text).ToString() + "'";
                        }
                        
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
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Matricula de Maestria", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;
            }
        }
    }
}
