using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
namespace excel_file_loader_fm{  
    public partial class Form2 : Form
    {
        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
        string tablenm = "";
        public Form2(string tablename)
        {
            InitializeComponent();
            tablenm = tablename;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            panel1.Visible = true;
            progressBar2.Style = ProgressBarStyle.Marquee;
            System.Threading.Thread thread =
              new System.Threading.Thread(new System.Threading.ThreadStart(loadDataFromDataBase));
            thread.Start();

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
                dataGridView1.DataSource = table;              
                panel1.Visible = false;
            }
        }


        private void loadDataFromDataBase(){
            try
            {
                builder.DataSource = Properties.Settings.Default["datasource"].ToString();
                builder.UserID = Properties.Settings.Default["username"].ToString();
                builder.Password = Properties.Settings.Default["password"].ToString();
                builder.InitialCatalog = Properties.Settings.Default["database"].ToString();
                string constr = builder.ConnectionString;
                var select = "SELECT * FROM dbo." + tablenm;
                var c = new SqlConnection(constr);
                var dataAdapter = new SqlDataAdapter(select, c);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                setDataSource(ds.Tables[0]);
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           // dataGridView1.DataSource = ds.Tables[0];
        }


    
        private void DataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {

                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
