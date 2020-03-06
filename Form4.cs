using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace excel_file_loader_fm
{
    public partial class Form4 : Form
    {
        public Form4(DataTable dt)
        {
            InitializeComponent();

            this.dataGridViewRowError.DataSource = dt;

        }

        private void Form4_Load(object sender, EventArgs e)
        {


        }

        public DataRow GetHeadersNew(DataTable dt)
        {
            DataRow row = dt.NewRow();
            DataColumnCollection columns = dt.Columns;


            for (int i = columns.Count; i-- > 0;)
            {
                row[i] = columns[i].ColumnName;
            }

            return row;


        }

        private void DataGridViewRowError_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
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
