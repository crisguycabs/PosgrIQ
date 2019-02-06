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

namespace PosgrIQ
{
    public partial class ConsultaForm : Form
    {
        public MainForm padre;

        public ConsultaForm()
        {
            InitializeComponent();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarConsultaForm();
        }

        private void ConsultaForm_Load(object sender, EventArgs e)
        {
            comboConsulta.SelectedIndex = 0;
        }

        private void btnConsultar_Click(object sender, EventArgs e)
        {
            // se inicia la consulta en función de la seleccion del combobox

            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);

            // algunas variables
            string query;
            OleDbCommand command;
            DataTable dt, dtEstudiantesMaes, dtEstudiantesDoct, dtPropuestasMaes;
            OleDbDataAdapter da;

            DateTime ini = new DateTime(dateIni.Value.Year, dateIni.Value.Month, dateIni.Value.Day);
            DateTime fin = new DateTime(dateFin.Value.Year, dateFin.Value.Month, dateFin.Value.Day);
            DateTime fecha = new DateTime();
            string[] fechaArray = null;

            switch (comboConsulta.SelectedIndex)
            {
                case 0: // Fecha de entrega del Tema de Maestria
                    #region tema de maestria
                    try
                    {
                        query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();
                        command = new OleDbCommand(query, conection);
                        da = new OleDbDataAdapter(command);
                        dtEstudiantesMaes = new DataTable();
                        da.Fill(dtEstudiantesMaes);

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        // se crea la tabla

                        dt = new DataTable();
                        dt.Columns.Add("Fecha", typeof(DateTime));
                        dt.Columns.Add("Codigo", typeof(int));
                        dt.Columns.Add("Nombre", typeof(string));
                        dt.Columns.Add("Tema", typeof(string));

                        // se recorre cada estudiante y se compara la fecha de inicio y fin con la de entrega del tema
                        for (int i = 0; i < dtEstudiantesMaes.Rows.Count; i++)
                        {
                            fechaArray = dtEstudiantesMaes.Rows[i][12].ToString().Split('/');
                            fecha = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));

                            if ((ini <= fecha) & (fecha <= fin))
                            {
                                // la fecha de la entrega del tema esta dentro de los limites
                                DataRow fila = dt.NewRow();

                                // fecha de entrega
                                fila[0] = fecha;

                                // codigo del estudiante
                                fila[1] = dtEstudiantesMaes.Rows[i][0];

                                // nombre del estudiante
                                fila[2] = dtEstudiantesMaes.Rows[i][1];

                                // tema
                                fila[3] = dtEstudiantesMaes.Rows[i][11];



                                dt.Rows.Add(fila);
                            }
                        }

                        // se enlaza el datatable con el datagrid
                        dataGridConsulta.DataSource = dt;

                        // se reescala el datagridview
                        MainForm.ReescalarDataGridView(ref dataGridConsulta);

                        dataGridConsulta.Sort(dataGridConsulta.Columns[0], ListSortDirection.Ascending);

                    }
                    catch
                    {
                        MessageBox.Show("No fue posible realizar la consulta solicitada", "Error al consultar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    #endregion
                    break;
                case 1:
                    #region tema de doctorado
                    try
                    {
                        query = "SELECT * FROM EstudiantesDoct ORDER BY codigo ASC";
                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();
                        command = new OleDbCommand(query, conection);
                        da = new OleDbDataAdapter(command);
                        dtEstudiantesDoct = new DataTable();
                        da.Fill(dtEstudiantesDoct);

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        // se crea la tabla

                        dt = new DataTable();
                        dt.Columns.Add("Fecha", typeof(DateTime));
                        dt.Columns.Add("Codigo", typeof(int));
                        dt.Columns.Add("Nombre", typeof(string));
                        dt.Columns.Add("Tema", typeof(string));

                        // se recorre cada estudiante y se compara la fecha de inicio y fin con la de entrega del tema
                        for (int i = 0; i < dtEstudiantesDoct.Rows.Count; i++)
                        {
                            try
                            {
                                fechaArray = dtEstudiantesDoct.Rows[i][12].ToString().Split('/');
                                if (fechaArray.Length > 1)
                                {
                                    fecha = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));

                                    if ((ini <= fecha) & (fecha <= fin))
                                    {
                                        // la fecha de la entrega del tema esta dentro de los limites
                                        DataRow fila = dt.NewRow();

                                        // fecha de entrega
                                        fila[0] = fecha;

                                        // codigo del estudiante
                                        fila[1] = dtEstudiantesDoct.Rows[i][0];

                                        // nombre del estudiante
                                        fila[2] = dtEstudiantesDoct.Rows[i][1];

                                        // tema
                                        fila[3] = dtEstudiantesDoct.Rows[i][11];

                                        dt.Rows.Add(fila);
                                    }
                                }
                            }
                            catch
                            {
                                int a = 1;
                                a++;
                            }
                        }

                        // se enlaza el datatable con el datagrid
                        dataGridConsulta.DataSource = dt;

                        // se reescala el datagridview
                        MainForm.ReescalarDataGridView(ref dataGridConsulta);

                        dataGridConsulta.Sort(dataGridConsulta.Columns[0], ListSortDirection.Ascending);

                    }
                    catch
                    {
                        MessageBox.Show("No fue posible realizar la consulta solicitada", "Error al consultar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    #endregion
                    break;
                case 2:
                    #region propuesta de maestria

                    try
                    {
                        // se pide la informacion de las propuestas de doctorado
                        query = "SELECT * FROM PropuestaMaes ORDER BY codigo ASC";
                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();
                        command = new OleDbCommand(query, conection);
                        da = new OleDbDataAdapter(command);
                        dtPropuestasMaes = new DataTable();
                        da.Fill(dtPropuestasMaes);

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        // se pide la informacion de los estudiantes de maestria
                        query = "SELECT * FROM EstudiantesMaes ORDER BY codigo ASC";
                        conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
                        conection.Open();
                        command = new OleDbCommand(query, conection);
                        da = new OleDbDataAdapter(command);
                        dtEstudiantesMaes = new DataTable();
                        da.Fill(dtEstudiantesMaes);

                        conection.Close();
                        conection.Dispose();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        while (System.IO.Directory.GetFiles(padre.sourceONE, "*.ldb").Length > 0)
                        {
                            System.Threading.Thread.Sleep(100);
                        }

                        dt = new DataTable();
                        dt.Columns.Add("Fecha", typeof(DateTime));
                        dt.Columns.Add("Codigo", typeof(int));
                        dt.Columns.Add("Nombre", typeof(string));
                        dt.Columns.Add("Propuesta", typeof(string));

                        for (int i = 0; i < dtPropuestasMaes.Rows.Count; i++)
                        {
                            try
                            {
                                fechaArray = dtPropuestasMaes.Rows[i][6].ToString().Split('/');
                                if (fechaArray.Length > 1)
                                {
                                    fecha = new DateTime(Convert.ToInt16(fechaArray[2]), Convert.ToInt16(fechaArray[1]), Convert.ToInt16(fechaArray[0]));

                                    if ((ini <= fecha) & (fecha <= fin))
                                    {
                                        // la fecha de la entrega del tema esta dentro de los limites
                                        DataRow fila = dt.NewRow();

                                        // fecha de entrega
                                        fila[0] = fecha;

                                        // codigo del estudiante
                                        fila[1] = dtPropuestasMaes.Rows[i][0];

                                        // nombre del estudiante
                                        DataRow[] seleccionado = dtEstudiantesMaes.Select("codigo=" + dtPropuestasMaes.Rows[i][1]);
                                        fila[2] = Convert.ToString(seleccionado[0][1]);

                                        // titulo de la propuesta
                                        fila[3] = dtPropuestasMaes.Rows[i][2];

                                        dt.Rows.Add(fila);
                                    }
                                }
                            }
                            catch
                            { }
                        }

                        // se enlaza el datatable con el datagrid
                        dataGridConsulta.DataSource = dt;

                        // se reescala el datagridview
                        MainForm.ReescalarDataGridView(ref dataGridConsulta);

                        dataGridConsulta.Sort(dataGridConsulta.Columns[0], ListSortDirection.Ascending);

                    }
                    catch
                    {
                        MessageBox.Show("No fue posible realizar la consulta solicitada", "Error al consultar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    #endregion
                    break;
                default:
                    break;

            }
        }
    }
}
