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
    public partial class AddSemestresForm : Form
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
        /// Codigo de la escuela a modificar
        /// </summary>
        public int codigo;

        /// <summary>
        /// Guarda la informacion de la tabla Escuelas antes de hacer cualquier modificacion
        /// </summary>
        public DataTable dtSemestres;

        #endregion

        public AddSemestresForm()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void AddSemestresForm_Load(object sender, EventArgs e)
        {
            padre.CheckConflicto();

            // se lee desde la BD la cantidad de Profesores, Colegiatura y Escuelas que existen actualmente
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                conection.Open();

                // algunas variables
                string query;
                OleDbCommand command;
                OleDbDataAdapter da;

                // se pide la informacion de los semestres
                query = "SELECT * FROM Semestres ORDER BY codigo ASC";
                command = new OleDbCommand(query, conection);

                da = new OleDbDataAdapter(command);
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

                switch (modo)
                {
                    case true: // se agrega un nuevo semestre

                        // nuevo codigo
                        codigo = dtSemestres.Rows.Count + 1;

                        // año
                        numAno.Value = DateTime.Now.Year;

                        // periodo
                        cmbPeriodo.SelectedIndex = 0;

                        // tomar qualify
                        dateTomarQualify.Value = DateTime.Now;

                        // tomar qualify
                        this.datePropuestaDoct.Value = DateTime.Now;

                        // tomar qualify
                        this.datePropuestaMaes.Value = DateTime.Now;

                        // tomar qualify
                        this.datePedirQualify.Value = DateTime.Now;

                        // tomar qualify
                        this.dateTema.Value = DateTime.Now;

                        this.Text = "AGREGAR NUEVO SEMESTRE";

                        break;

                    case false: // se modifica una escuela

                        // se escriben en los controles la informacion de la escuela seleccionada

                        DataRow[] seleccionado = dtSemestres.Select("codigo=" + codigo.ToString());

                        numAno.Value = Convert.ToDecimal(seleccionado[0][2]);

                        switch (Convert.ToString(seleccionado[0][3]))
                        {
                            case "1":
                                cmbPeriodo.SelectedIndex = 0;
                                break;
                            case "2":
                                cmbPeriodo.SelectedIndex = 1;
                                break;
                        }

                        dateTomarQualify.Value = MainForm.Texto2Fecha((string)seleccionado[0][4]);

                        datePropuestaDoct.Value = MainForm.Texto2Fecha((string)seleccionado[0][5]);

                        datePropuestaMaes.Value = MainForm.Texto2Fecha((string)seleccionado[0][6]);

                        datePedirQualify.Value = MainForm.Texto2Fecha((string)seleccionado[0][7]);

                        dateTema.Value = MainForm.Texto2Fecha((string)seleccionado[0][8]);

                        btnAdd.Text = "Modificar";
                        this.Text = "MODIFICAR SEMESTRE";

                        break;
                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // se realizan algunas comprobaciones de seguridad

            // Solo revisar en modo insercion
            if (modo)
            {
                // existe un semestre con ese codigo. 
                DataRow[] busqueda = dtSemestres.Select("codigo=" + codigo.ToString());
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un semestre con el codigo " + codigo.ToString(), "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // ya existen dos semestres en el año que se ha indicado
                busqueda = dtSemestres.Select("ano=" + numAno.Value.ToString());
                if (busqueda.Length > 2)
                {
                    MessageBox.Show("Ya existen dos semestres en el año " + numAno.Value.ToString() + ". No se puede agregar un tercero", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // ya existe un año-periodo con ese nombre
                busqueda = dtSemestres.Select("nombre='" + numAno.Value.ToString() + "-" + cmbPeriodo.Items[cmbPeriodo.SelectedIndex].ToString() + "'");
                if (busqueda.Length > 0)
                {
                    MessageBox.Show("Ya existe un semestre en ese año y periodo", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // se duplican fechas
            if (dateTomarQualify.Value == datePropuestaDoct.Value)
            {
                MessageBox.Show("Las fechas de tomar el Qualify y la entrega de la propuesta de doctorado son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value == datePropuestaMaes.Value)
            {
                MessageBox.Show("Las fechas de tomar el Qualify y la entrega de la propuesta de maestria son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value == datePedirQualify.Value)
            {
                MessageBox.Show("Las fechas de tomar el Qualify y solicitar el Qualify son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value == dateTema.Value)
            {
                MessageBox.Show("Las fechas de tomar el Qualify y la entrega del tema son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePropuestaDoct.Value == datePropuestaMaes.Value)
            {
                MessageBox.Show("Las fechas de la entrega de propuesta de doctorado y la entrega de propuesta de maestria son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaDoct.Value == datePedirQualify.Value)
            {
                MessageBox.Show("Las fechas de la entrega de propuesta de doctorado y solicitar el Qualify son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaDoct.Value == dateTema.Value)
            {
                MessageBox.Show("Las fechas de la entrega de propuesta de doctorado y la entrega del tema son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePropuestaMaes.Value == datePedirQualify.Value)
            {
                MessageBox.Show("Las fechas de la entrega de propuesta de maestria y solicitar el Qualify son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaMaes.Value == dateTema.Value)
            {
                MessageBox.Show("Las fechas de la entrega de propuesta de maestria y la entrega del tema son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePedirQualify.Value == dateTema.Value)
            {
                MessageBox.Show("Las fechas de solicitud del Qualify y la entrega del tema son iguales", "Error de duplicado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            // ------------------------------------------------------
            if (dateTomarQualify.Value > datePropuestaDoct.Value)
            {
                MessageBox.Show("La fecha de tomar el Qualify no puede ser posterior a la entrega de la propuesta de doctorado", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value > datePropuestaMaes.Value)
            {
                MessageBox.Show("La fecha de tomar el Qualify no puede ser posterior a la entrega de la propuesta de maestria", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value > datePedirQualify.Value)
            {
                MessageBox.Show("La fecha de tomar el Qualify no puede ser posterior a solicitar el Qualify", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTomarQualify.Value > dateTema.Value)
            {
                MessageBox.Show("La fecha de tomar el Qualify no puede ser posterior a la entrega del tema", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePropuestaDoct.Value > datePropuestaMaes.Value)
            {
                MessageBox.Show("La fecha de la entrega de propuesta de doctorado no puede ser posterior a la entrega de propuesta de maestria", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaDoct.Value > datePedirQualify.Value)
            {
                MessageBox.Show("La fecha de la entrega de propuesta de doctorado no puede ser posterior a solicitar el Qualify", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaDoct.Value > dateTema.Value)
            {
                MessageBox.Show("La fecha de la entrega de propuesta de doctorado no puede ser posterior a la entrega del tema", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePropuestaMaes.Value > datePedirQualify.Value)
            {
                MessageBox.Show("La fecha de la entrega de propuesta de maestria no puede ser posterior a solicitar el Qualify", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (datePropuestaMaes.Value > dateTema.Value)
            {
                MessageBox.Show("La fecha de la entrega de propuesta de maestria no puede ser posterior a la entrega del tema", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            if (datePedirQualify.Value > dateTema.Value)
            {
                MessageBox.Show("La fecha de solicitud del Qualify no puede ser posterior a la entrega del tema", "Error de fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ------------------------------------------------------
            // ------------------------------------------------------


            // se prepara la conexion
            OleDbConnection conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            string query;
            OleDbCommand command;

            switch (modo)
            {
                case true:

                    // se agrega el semestre
                    try
                    {
                        datePropuestaDoct.Format = DateTimePickerFormat.Short;

                        // se prepara la cadena SQL
                        query = "INSERT INTO Semestres VALUES(";
                        query += codigo.ToString() + ",";
                        query += "'" + numAno.Value.ToString() + "-" + cmbPeriodo.Items[cmbPeriodo.SelectedIndex] + "',";
                        query += numAno.Value.ToString() + ",";
                        query += cmbPeriodo.Items[cmbPeriodo.SelectedIndex] + ",";
                        query += "'" + MainForm.Fecha2Texto(dateTomarQualify.Value) + "',";
                        query += "'" + MainForm.Fecha2Texto(datePropuestaDoct.Value) + "',";
                        query += "'" + MainForm.Fecha2Texto(datePropuestaMaes.Value) + "',";
                        query += "'" + MainForm.Fecha2Texto(datePedirQualify.Value) + "',";
                        query += "'" + MainForm.Fecha2Texto(dateTema.Value) + "')";

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

                        //numAno.Value = 0;
                        codigo++;
                        cmbPeriodo.SelectedIndex = -1;

                        //this.DialogResult = DialogResult.OK;
                        try { 
                        padre.semestresForm.SemestresForm_Load(sender, e);
                        }
                        catch { }
                    }
                    catch
                    {
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Semestres", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case false:

                    // se modifica el semestre
                    try
                    {
                        // se prepara la cadena SQL

                        query = "UPDATE Semestres SET ";
                        query += "codigo=" + codigo.ToString() + ", ";
                        query += "nombre='" + numAno.Value.ToString() + "-" + cmbPeriodo.Items[cmbPeriodo.SelectedIndex] + "', ";
                        query += "ano=" + numAno.Value.ToString() + ", ";
                        query += "periodo=" + cmbPeriodo.Items[cmbPeriodo.SelectedIndex] + ", ";
                        query += "tomarqualify='" + MainForm.Fecha2Texto(dateTomarQualify.Value) + "', ";
                        query += "propuestadoct='" + MainForm.Fecha2Texto(datePropuestaDoct.Value) + "', ";
                        query += "propuestamaes='" + MainForm.Fecha2Texto(datePropuestaMaes.Value) + "', ";
                        query += "pedirqualify='" + MainForm.Fecha2Texto(datePedirQualify.Value) + "', ";
                        query += "tema='" + MainForm.Fecha2Texto(dateTema.Value) + "' ";
                        query += "WHERE codigo=" + codigo.ToString();

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
                        MessageBox.Show("No se puede acceder a la base de datos, tabla Semestres", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    break;
            }
        }

        private void numAno_ValueChanged(object sender, EventArgs e)
        {
            DateTime actual;

            actual = datePedirQualify.Value;
            datePedirQualify.Value = new DateTime((int)numAno.Value, actual.Month, actual.Day);

            actual = datePropuestaDoct.Value;
            datePropuestaDoct.Value = new DateTime((int)numAno.Value, actual.Month, actual.Day);

            actual = datePropuestaMaes.Value;
            datePropuestaMaes.Value = new DateTime((int)numAno.Value, actual.Month, actual.Day);

            actual = dateTema.Value;
            dateTema.Value = new DateTime((int)numAno.Value, actual.Month, actual.Day);

            actual = dateTomarQualify.Value;
            dateTomarQualify.Value = new DateTime((int)numAno.Value, actual.Month, actual.Day);
        }
    }
}
