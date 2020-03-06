using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace excel_file_loader_fm
{
    public partial class Form1 : Form
    {
        OleDbConnection OleDbcon;
        string sheetName;
        string comboname = "";
        bool error1 = false;
        bool tableexist = false;
        int startrow = 0;
        int endrow = 0;
        DataTable dat;
        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
        public Form1()
        {
            InitializeComponent();
            if (comboBox1.Items.Count == 0)
            {
                loadButton.Enabled = false;
            }
            else
            {
                loadButton.Enabled = true;

            }
            if (dataGridView2.Rows.Count > 0)
            {
                savesqlButton.Enabled = true;
                buttonError.Enabled = true;
            }
            else
            {
                savesqlButton.Enabled = false;
                buttonError.Enabled = false;
            }


        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'exloaddataDataSet.financial' table. You can move, or remove it, as needed.
            // this.financialTableAdapter.Fill(this.exloaddataDataSet.financial);
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.FileName = "";
            openfiledialog1.Filter = "Excel File|*.xlsx;*.xls";
            openfiledialog1.ShowDialog();
            if (!string.IsNullOrEmpty(openfiledialog1.FileName))
            {
                String conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openfiledialog1.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;';";
                OleDbcon = new OleDbConnection(conStr);
                OleDbcon.Open();
                DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                OleDbcon.Close();
                textBox1.Text = openfiledialog1.FileName;
                comboBox1.Items.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)

                {

                    sheetName = dt.Rows[i]["TABLE_NAME"].ToString();

                    sheetName = sheetName.Substring(0, sheetName.Length - 1);

                    comboBox1.Items.Add(sheetName);

                }
                comboBox1.SelectedIndex = 0;

            }

            if (comboBox1.Items.Count == 0)
            {
                loadButton.Enabled = false;
            }
            else
            {
                loadButton.Enabled = true;

            }
        }


        private void loadTable()
        {
            error1 = false;
            tableexist = false;
            startrow = 0;
            endrow = 0;
            dat = null;
            //OleDbCommand oconn= new OleDbCommand("Select * from [" + comboBox1.Text + "$]", OleDbcon);
            try
            {

                if (comboBox1.InvokeRequired)
                {
                    comboBox1.Invoke(new MethodInvoker(delegate { comboname = comboBox1.Text; }));
                }
                OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from [" + comboname + "$]", OleDbcon);
                DataTable data = new DataTable();
                oledbDa.Fill(data);
                dat = data.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string)))
.CopyToDataTable();

                // dataGridView2.DataSource = dat;
                setDataSource(dat);
            }
            catch (Exception ee)
            {
                if (panel1.InvokeRequired)
                {
                    panel1.Invoke(new MethodInvoker(delegate { panel1.Visible = false; }));
                }
                MessageBox.Show("Can't import excel file", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal delegate void SetDataSourceDelegate(DataTable table);
        private void setDataSource(DataTable table)
        {
            // Invoke method if required:
            if (this.InvokeRequired)
            {
                this.Invoke(new SetDataSourceDelegate(setDataSource), table);
            }
            else
            {
                dataGridView2.DataSource = table;
                int numRows = dataGridView2.Rows.Count;

                for (int i = 0; i < numRows; i++)
                {
                    comboBoxTo.Items.Add(i + 1);
                    comboBoxFrom.Items.Add(i + 1);

                }
                comboBoxFrom.SelectedIndex = 0;
                int last = comboBoxTo.Items.Count - 1;
                comboBoxTo.SelectedIndex = last;
                panel1.Visible = false;
            }

        }
        private void Button2_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            textprogress.Text = "Loading Excel Sheet...";
            progressBar1.Style = ProgressBarStyle.Marquee;
            System.Threading.Thread thread =
              new System.Threading.Thread(new System.Threading.ThreadStart(loadTable));
            thread.Start();

        }
        /*
        private void copytoserver(DataTable dt, SqlBulkCopy sqlBulkCopy)
        {
            sqlBulkCopy.WriteToServer(dt);
            int rowsCopied = Class1.GetRowsCopied(sqlBulkCopy);
            if (panel1.InvokeRequired)
            {
                panel1.Invoke(new MethodInvoker(delegate { panel1.Visible = false; MessageBox.Show("Data saved successfully. Number of rows copied " + rowsCopied + ".", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            }
         
        }
        */



        public bool tablesAreTheSame(DataTable table1, DataTable table2)
        {
            DataTable dt;
            dt = getDifferentRecords(table1, table2);

            if (dt.Rows.Count == 0)
                return true;
            else
                return false;
        }




        private void loadTableToServer()
        {
            dataGridView2.ClearSelection();
            validatedata();
            DataTable dtNew;
            if (dat.Rows.Count > 0 && error1 == false)
            {

                if (comboBoxFrom.InvokeRequired) { comboBoxFrom.Invoke(new MethodInvoker(delegate { startrow = Convert.ToInt32(comboBoxFrom.Text); })); }
                if (comboBoxTo.InvokeRequired) { comboBoxTo.Invoke(new MethodInvoker(delegate { endrow = Convert.ToInt32(comboBoxTo.Text); })); }


                startrow = startrow - 1;
                int takerow = endrow - startrow;
                dtNew = dat.Select().Skip(startrow).Take(takerow + 1).CopyToDataTable();
                builder.DataSource = Properties.Settings.Default["datasource"].ToString();
                builder.UserID = Properties.Settings.Default["username"].ToString();
                builder.Password = Properties.Settings.Default["password"].ToString();
                builder.InitialCatalog = Properties.Settings.Default["database"].ToString();
                //conStr = @"Data Source=CHACHA\SQLEXPRESS;Initial Catalog=exloaddata;User ID=sa;Password=henry5765";
                using (SqlConnection con = new SqlConnection(builder.ConnectionString))
                {
                    con.Open();
                    try
                    {
                        string createString = createTable(dtNew, RemoveSpecialCharacters(comboname));
                        using (SqlCommand command = new SqlCommand(createString, con))
                            command.ExecuteNonQuery();
                        tableexist = false;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                        tableexist = true;
                    }

                    if (!tableexist)
                    {
                        saveToDatabase(dtNew, builder.ConnectionString, RemoveSpecialCharacters(comboname));
                    }
                    else
                    {

                        if (panel1.InvokeRequired)
                        {
                            panel1.Invoke(new MethodInvoker(delegate { panel1.Visible = false; }));
                        }
                        DialogResult rf = MessageBox.Show("Unable to create new table,Table already exists.Give sheet unique name or add data to current table. Add data to existing Table?", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        string rsstring = rf.ToString();
                        if (rsstring == "OK")
                        {
  
                            if (!isTableSame(dtNew, getFirstThreeColumnFromDataBase(comboname)))
                            {
                                saveToDatabase(dtNew, builder.ConnectionString, RemoveSpecialCharacters(comboname));
                            }
                            else
                            {
                                MessageBox.Show("Cannot save record. Duplicate records found", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                    }
                    // copytoserver(dtNew, sqlBulkCopy);                         
                    con.Close();

                }
            }



        }

        private void Button3_Click(object sender, EventArgs e)
        {
            textprogress.Text = "Saving data to database...";
            panel1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(loadTableToServer));
            thread.Start();

        }

        private void DataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2(RemoveSpecialCharacters(comboname));
            f2.ShowDialog();
        }

        private void DataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

            if (dataGridView2.Rows.Count > 0)
            {
                savesqlButton.Enabled = true;
                buttonError.Enabled = true;
            }
            else
            {
                savesqlButton.Enabled = false;
                buttonError.Enabled = false;
            }
            //setRowNumber(dataGridView2);
            this.dataGridView2.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.ClearSelection();


        }

        private void dgGrid_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
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

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click_1(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.ShowDialog();
        }


        private void validatedata()
        {

            foreach (DataGridViewRow rw in this.dataGridView2.Rows)
            {

                for (int i = 0; i < rw.Cells.Count; i++)
                {
                    if (rw.Cells[i].Value == null || rw.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(rw.Cells[i].Value.ToString()))
                    {
                        //dataGridView2.CurrentCell = rw.Cells[i];
                        rw.Selected = true;
                        // rw.Cells[i].Selected = true;
                        error1 = true;

                    }
                }
            }

            if (error1 == true)
            {
                if (panel1.InvokeRequired)
                {
                    panel1.Invoke(new MethodInvoker(delegate { panel1.Visible = false; }));
                }
                //dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Red;
                MessageBox.Show("Some cells are empty", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);

            }
        }
        private void Button3_Click_1(object sender, EventArgs e)
        {
            DataTable dt = new DataTable(); // create a table for storing selected rows
            var dtTemp = dataGridView2.DataSource as DataTable; // get the source table object
            dt = dtTemp.Clone();  // clone the schema of the source table to new table
            dt.Columns.Add("Record_Number", typeof(System.String));
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Selected)
                {
                    var row = dt.NewRow();

                    // create a new row with the schema 
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        row[j] = dataGridView2[j, i].Value;
                    }
                    row["Record_Number"] = dataGridView2.Rows[i].Index + 1;
                    dt.Rows.Add(row);

                }

            }

            Form4 f4 = new Form4(dt);
            f4.ShowDialog();
        }

        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }


        private void saveToDatabase(DataTable myDataTable, string connectionString, string tablename)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    foreach (DataColumn c in myDataTable.Columns)
                        bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);

                    bulkCopy.DestinationTableName = tablename;
                    try
                    {
                        bulkCopy.WriteToServer(myDataTable);
                        int rowsCopied = Class1.GetRowsCopied(bulkCopy);
                        if (panel1.InvokeRequired)
                        {
                            panel1.Invoke(new MethodInvoker(delegate
                            {
                                buttonDataSql.Enabled = true;
                                panel1.Visible = false; MessageBox.Show("Data saved successfully. Number of rows copied " + rowsCopied + ".", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (panel1.InvokeRequired)
                        {
                            panel1.Invoke(new MethodInvoker(delegate { panel1.Visible = false; }));
                        }
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }
        private DataTable getFirstThreeColumnFromDataBase(string tablename)
        {
            try
            {
                builder.DataSource = Properties.Settings.Default["datasource"].ToString();
                builder.UserID = Properties.Settings.Default["username"].ToString();
                builder.Password = Properties.Settings.Default["password"].ToString();
                builder.InitialCatalog = Properties.Settings.Default["database"].ToString();
                string constr = builder.ConnectionString;
                var select = "SELECT * FROM dbo." + tablename;
                var c = new SqlConnection(constr);
                var dataAdapter = new SqlDataAdapter(select, c);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataTable();
                dataAdapter.Fill(ds);
                return ds;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;

            }

        }

        public DataTable checkNewData(DataTable dataTable1, DataTable dataTable2)
        {
            var differences =
                dataTable1.AsEnumerable().Except(dataTable2.AsEnumerable(),
                                                        DataRowComparer.Default);

            return differences.Any() ? differences.CopyToDataTable() : new DataTable();
        }
        public static bool isTableSame(DataTable tbl1, DataTable tbl2)
        {
            /*
            if (tbl1.Rows.Count == tbl2.Rows.Count)
                return true; 
                */
            
            for (int i = 0; i < tbl1.Rows.Count; i++)
            {
                for (int c = 0; c < tbl1.Columns.Count; c++)
                {
                    if (Equals(tbl2.Rows[i][c], tbl1.Rows[i][c]))
                        return true;
                }
            }
            return false;
        }


        private DataTable getDifferentRecords(DataTable FirstDataTable, DataTable SecondDataTable)
        {
            //Create Empty Table     
            DataTable ResultDataTable = new DataTable("ResultDataTable");

            //use a Dataset to make use of a DataRelation object     
            using (DataSet ds = new DataSet())
            {
                //Add tables     
                ds.Tables.AddRange(new DataTable[] { FirstDataTable.Copy(), SecondDataTable.Copy() });

                //Get Columns for DataRelation     
                DataColumn[] firstColumns = new DataColumn[ds.Tables[0].Columns.Count];
                for (int i = 0; i < firstColumns.Length; i++)
                {
                    firstColumns[i] = ds.Tables[0].Columns[i];
                }

                DataColumn[] secondColumns = new DataColumn[ds.Tables[1].Columns.Count];
                for (int i = 0; i < secondColumns.Length; i++)
                {
                    secondColumns[i] = ds.Tables[1].Columns[i];
                }

                //Create DataRelation     
                DataRelation r1 = new DataRelation(string.Empty, firstColumns, secondColumns, false);
                ds.Relations.Add(r1);

                DataRelation r2 = new DataRelation(string.Empty, secondColumns, firstColumns, false);
                ds.Relations.Add(r2);

                //Create columns for return table     
                for (int i = 0; i < FirstDataTable.Columns.Count; i++)
                {
                    ResultDataTable.Columns.Add(FirstDataTable.Columns[i].ColumnName, FirstDataTable.Columns[i].DataType);
                }

                //If FirstDataTable Row not in SecondDataTable, Add to ResultDataTable.     
                ResultDataTable.BeginLoadData();
                foreach (DataRow parentrow in ds.Tables[0].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r1);
                    if (childrows == null || childrows.Length == 0)
                        ResultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }

                //If SecondDataTable Row not in FirstDataTable, Add to ResultDataTable.     
                foreach (DataRow parentrow in ds.Tables[1].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r2);
                    if (childrows == null || childrows.Length == 0)
                        ResultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }
                ResultDataTable.EndLoadData();
            }

            return ResultDataTable;
        }




        private static string createTable(DataTable table, string tableName)
        {
            string sqlsc;
            sqlsc = "CREATE TABLE " + tableName + "(";
            for (int i = 0; i < table.Columns.Count; i++)
            {

                sqlsc += "\n [" + table.Columns[i].ColumnName + "] ";
                string columnType = table.Columns[i].DataType.ToString();
                switch (columnType)
                {
                    case "System.Int32":
                        sqlsc += " int ";
                        break;
                    case "System.Int64":
                        sqlsc += " bigint ";
                        break;
                    case "System.Int16":
                        sqlsc += " smallint";
                        break;
                    case "System.Byte":
                        sqlsc += " tinyint";
                        break;
                    case "System.Decimal":
                        sqlsc += " decimal ";
                        break;
                    case "System.DateTime":
                        sqlsc += " datetime ";
                        break;
                    case "System.String":
                    default:
                        sqlsc += string.Format(" nvarchar({0}) ", table.Columns[i].MaxLength == -1 ? "max" : table.Columns[i].MaxLength.ToString());
                        break;
                }
                if (table.Columns[i].AutoIncrement)
                    sqlsc += " IDENTITY(" + table.Columns[i].AutoIncrementSeed.ToString() + "," + table.Columns[i].AutoIncrementStep.ToString() + ") ";
                //if (!table.Columns[i].AllowDBNull)
                sqlsc += " NOT NULL ";
                sqlsc += ",";
            }
            return sqlsc.Substring(0, sqlsc.Length - 1) + "\n)";

        }

    }
}
