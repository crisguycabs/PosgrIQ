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
    public partial class ProfesoresForm : Form
    {
        #region variables de clase

        /// <summary>
        /// Instancia del MainForm padre
        /// </summary>
        public MainForm padre;

        #endregion

        public ProfesoresForm()
        {
            InitializeComponent();
        }

        public void test()
        {
            // Set Access connection and select strings.
            // The path to BugTypes.MDB must be changed if you build 
            // the sample from the command line:
#if USINGPROJECTSYSTEM
            string strAccessConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\\..\\BugTypes.MDB";
#else
            string strAccessConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BugTypes.MDB";
#endif
            string strAccessSelect = "SELECT * FROM Categories";

            // Create the dataset and add the Categories table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(myDataSet, "Categories");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }

            // A dataset can contain multiple tables, so let's get them
            // all into an array:
            DataTableCollection dta = myDataSet.Tables;
            foreach (DataTable dt in dta)
            {
                Console.WriteLine("Found data table {0}", dt.TableName);
            }

            // The next two lines show two different ways you can get the
            // count of tables in a dataset:
            Console.WriteLine("{0} tables in data set", myDataSet.Tables.Count);
            Console.WriteLine("{0} tables in data set", dta.Count);
            // The next several lines show how to get information on
            // a specific table by name from the dataset:
            Console.WriteLine("{0} rows in Categories table", myDataSet.Tables["Categories"].Rows.Count);
            // The column info is automatically fetched from the database,
            // so we can read it here:
            Console.WriteLine("{0} columns in Categories table", myDataSet.Tables["Categories"].Columns.Count);
            DataColumnCollection drc = myDataSet.Tables["Categories"].Columns;
            int i = 0;
            foreach (DataColumn dc in drc)
            {
                // Print the column subscript, then the column's name
                // and its data type:
                Console.WriteLine("Column name[{0}] is {1}, of type {2}", i++, dc.ColumnName, dc.DataType);
            }
            DataRowCollection dra = myDataSet.Tables["Categories"].Rows;
            foreach (DataRow dr in dra)
            {
                // Print the CategoryID as a subscript, then the CategoryName:
                Console.WriteLine("CategoryName[{0}] is {1}", dr[0], dr[1]);
            }
        }

        public void TestGembox()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            // Create new spreadsheet type file.
            var file = new ExcelFile();
            var sheet = file.Worksheets.Add("Sheet");

            // Create some sample content and format it.
            foreach (var cell in sheet.Cells.GetSubrange("B2", "D2"))
            {
                cell.Value = "Sample";
                //cell.Style.FillPattern.SetSolid(Color.LightBlue);
                //cell.SetBorders(MultipleBorders.Top, Color.DarkBlue, LineStyle.Medium);
            }
            foreach (var cell in sheet.Cells.GetSubrange("A4", "E5"))
            {
                cell.Value = "Sample";
                //cell.Style.FillPattern.SetSolid(Color.LightGreen);
                //cell.SetBorders(MultipleBorders.Outside, Color.DarkGreen, LineStyle.Medium);
            }
            foreach (var cell in sheet.Cells.GetSubrange("B7", "D7"))
            {
                cell.Value = "Sample";
                //cell.Style.FillPattern.SetSolid(Color.LightGray);
                //cell.SetBorders(MultipleBorders.Bottom, Color.DarkGray, LineStyle.Medium);
            }

            // Print spreadsheet content.
            //sheet.PrintOptions.PrintGridlines = true;
            sheet.NamedRanges.SetPrintArea(sheet.Cells.GetSubrange("A1", "E8"));
            file.Print();
            // Or save it as PDF file.
            file.Save("Sample.pdf");
        }

        public void ProfesoresForm_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=" + padre.sourceBD);
            try
            {
                // algunas variables
                string query;
                OleDbCommand command;
                DataTable dt, dtProfesores, dtColegiatura, dtEscuelas;
                OleDbDataAdapter da;

                // se pide la informacion de los profesores
                query = "SELECT * FROM Profesores ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtProfesores = new DataTable();
                da.Fill(dtProfesores);

                // se pide la informacion de las colegiaturas
                query = "SELECT * FROM Colegiatura ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtColegiatura = new DataTable();
                da.Fill(dtColegiatura);

                // se pide la informacion de las escuelas
                query = "SELECT * FROM Escuelas ORDER BY codigo ASC";
                conection.Open();
                command = new OleDbCommand(query, conection);
                conection.Close();

                da = new OleDbDataAdapter(command);
                dtEscuelas = new DataTable();
                da.Fill(dtEscuelas);

                // se crea una tabla que reemplaza el codigo de colegiatura y de escuela por el nombre
                dt = new DataTable();
                dt.Columns.Add("Codigo", typeof(int));
                dt.Columns.Add("Nombre", typeof(string));
                dt.Columns.Add("Colegiatura", typeof(string));
                dt.Columns.Add("Escuela", typeof(string));
                dt.Columns.Add("Correo", typeof(string));
                dt.Columns.Add("Telefono", typeof(string));
                dt.Columns.Add("Universidad", typeof(string));

                // se llena el nuevo datatable
                for (int i = 0; i < dtProfesores.Rows.Count; i++)
                {
                    DataRow fila = dt.NewRow();

                    // codigo del profesor
                    fila[0] = dtProfesores.Rows[i][0];

                    // nombre del profesor
                    fila[1] = dtProfesores.Rows[i][1];

                    // colegiatura
                    fila[2] = dtColegiatura.Select("codigo=" + dtProfesores.Rows[i][2].ToString())[0][1];

                    // escuela
                    fila[3] = dtEscuelas.Select("codigo=" + dtProfesores.Rows[i][3].ToString())[0][1];

                    // correo
                    fila[4] = dtProfesores.Rows[i][4];

                    // telefono
                    fila[5] = dtProfesores.Rows[i][5];

                    // escuela
                    fila[6] = dtProfesores.Rows[i][6];

                    dt.Rows.Add(fila);
                }
                
                // se enlaza el datatable con el datagrid
                dataGridProfesores.DataSource = dt;

                // se reescala el datagridview
                MainForm.ReescalarDataGridView(ref dataGridProfesores);

                dataGridProfesores.Sort(dataGridProfesores.Columns[0], ListSortDirection.Ascending);
            }
            catch
            {
                MessageBox.Show("No se puede acceder a la base de datos, tabla Profesores", "Error de conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            padre.CerrarProfesoresForm();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddProfesoresForm agregar = new AddProfesoresForm();
            agregar.padre = this.padre;
            agregar.modo = 1;

            if (agregar.ShowDialog() == DialogResult.OK) this.ProfesoresForm_Load(sender, e);
        }

        private void btnMod_Click(object sender, EventArgs e)
        {
            if (dataGridProfesores.Rows.Count < 1) return;
            
            AddProfesoresForm modificar = new AddProfesoresForm();
            modificar.padre = this.padre;
            modificar.modo = 0;
            modificar.codigo = Convert.ToInt32(dataGridProfesores.SelectedRows[0].Cells[0].Value);

            if (modificar.ShowDialog() == DialogResult.OK) this.ProfesoresForm_Load(sender, e);
        }

        private void dataGridProfesores_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
