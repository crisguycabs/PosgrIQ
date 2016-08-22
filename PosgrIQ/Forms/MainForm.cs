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
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using GemBox.Spreadsheet;

namespace PosgrIQ
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var conection = new OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;" + "data source=database\\dbposgriq.mdb");
            try
            {
                conection.Open();

                var query = "SELECT * FROM Profesores";
                var command = new OleDbCommand(query, conection);

                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.DataSource = dt;

                conection.Close();
            }
            catch
            {


            }

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

        public void Test1()
        {
            // Create a new PDF document

            PdfDocument document = new PdfDocument();

            document.Info.Title = "Created with PDFsharp";

            // Create an empty page

            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Verdana", 20, XFontStyle.BoldItalic);

            // Draw the text
            gfx.DrawString("Hello, World!", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.TopCenter);

            // Save the document...
            const string filename = "HelloWorld.pdf";
            document.Save(filename);
            // ...and start a viewer.
            System.Diagnostics.Process.Start(filename);
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

        private void btnPdf_Click(object sender, EventArgs e)
        {
            TestGembox();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
