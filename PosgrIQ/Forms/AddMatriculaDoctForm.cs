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

                foreach (DataRow row in dtEstudiantes.Rows)
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

        private void AddMatriculaDoctForm_Load(object sender, EventArgs e)
        {
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
                MessageBox.Show("No se puede acceder a la base de datos, tabla MatriculaDoct", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            
        }
    }
}
