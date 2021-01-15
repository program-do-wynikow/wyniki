using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace WIZUALIZACJA_CAT_STREAM
{
    class ExcelHandler
    {
        public DataTable data { get; set; }
        public OleDbDataAdapter adapter { get; set; }
        public DataSet ds { get; set; }

        public ExcelHandler()
        {
            data = null;
            adapter = null;
            ds = null;
        }
        public ExcelHandler(string fileName, string spreadsheet)
        {
            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
            adapter = new OleDbDataAdapter($"SELECT * FROM [{spreadsheet}$]", connectionString);
            ds = new DataSet();
            adapter.Fill(ds, "A");
            data = ds.Tables["A"];
        }
        public string Row_toString(int row)
        {
            string names = "", club = data.Rows[row][data.Columns.Count-1].ToString();
            for (int j = 0; j < data.Columns.Count-1; j++) names += data.Rows[row][j].ToString() + " ";
            return Pretty_String.Generate(names, club);
        }
    }
}
