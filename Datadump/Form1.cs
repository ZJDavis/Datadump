using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Datadump
{
    public partial class Form1 : Form
    {
        private SQLiteConnection sqliteConn;

        public Form1()
        {
            InitializeComponent();
            // Initialize SQLite connection.
            sqliteConn = new SQLiteConnection("Data Source=sample.db;Version=3;");
            sqliteConn.Open();
            CreateTable();
        }

        private void CreateTable()
        {
            string createTableQuery = @"
                CREATE TABLE IF NOT EXISTS Data (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Date TEXT,
                    Column1 TEXT,
                    Column2 TEXT
                    -- Add other columns as needed based on your CSV structure.
                )";
            using (var cmd = new SQLiteCommand(createTableQuery, sqliteConn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        private void btnLoadCsv_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "CSV files (*.csv)|*.csv";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadCsvIntoDatabase(ofd.FileName);
                }
            }
        }

        private void LoadCsvIntoDatabase(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            var headers = lines[0].Split(',');

            using (var transaction = sqliteConn.BeginTransaction())
            {
                foreach (var line in lines.Skip(1))
                {
                    var values = line.Split(',');

                    string insertQuery = "INSERT INTO Data (Date, Column1, Column2) VALUES (@Date, @Column1, @Column2)";
                    using (var cmd = new SQLiteCommand(insertQuery, sqliteConn))
                    {
                        cmd.Parameters.AddWithValue("@Date", values[0]);
                        cmd.Parameters.AddWithValue("@Column1", values[1]);
                        cmd.Parameters.AddWithValue("@Column2", values[2]);
                        // Add other columns as needed.
                        cmd.ExecuteNonQuery();
                    }
                }
                transaction.Commit();
            }

            MessageBox.Show("CSV file loaded successfully.");
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            DateTime startDate = dateTimePickerStart.Value;
            DateTime endDate = dateTimePickerEnd.Value;
            ExportDataToExcel(startDate, endDate);
        }

        private void ExportDataToExcel(DateTime startDate, DateTime endDate)
        {
            string selectQuery = "SELECT * FROM Data WHERE Date BETWEEN @startDate AND @endDate";
            using (var cmd = new SQLiteCommand(selectQuery, sqliteConn))
            {
                cmd.Parameters.AddWithValue("@startDate", startDate.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@endDate", endDate.ToString("yyyy-MM-dd"));

                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx" })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            using (var package = new ExcelPackage())
                            {
                                var worksheet = package.Workbook.Worksheets.Add("Data");
                                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                                package.SaveAs(new FileInfo(sfd.FileName));
                                MessageBox.Show("Data exported to Excel successfully.");
                            }
                        }
                    }
                }
            }
        }

        private void btnLoadCsv_Click_1(object sender, EventArgs e)
        {

        }
    }
}
