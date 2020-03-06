using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace excel_file_loader_fm
{
    public partial class Form3 : Form
    {
        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

        public Form3()
        {
            InitializeComponent();
            try
            {
                textBoxDataSource.Text = Properties.Settings.Default["datasource"].ToString();
                textBoxUserName.Text = Properties.Settings.Default["username"].ToString();
                textBoxPassword.Text = Properties.Settings.Default["password"].ToString();
                textBoxDatabaseName.Text = Properties.Settings.Default["database"].ToString();
            }
            catch (Exception e)
            {

            }
            okButton.Enabled = false;
        }

        private void ButtonTestConn_Click(object sender, EventArgs e)
        {

            if (textBoxDataSource.Text.Equals("") || textBoxDatabaseName.Text.Equals("") || textBoxPassword.Text.Equals("") || textBoxUserName.Text.Equals(""))
            {
                MessageBox.Show("Can't test connection. Some values are empty.Please fill all fields.", "Message",
    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {

                builder.DataSource = textBoxDataSource.Text;
                builder.UserID = textBoxUserName.Text;
                builder.Password = textBoxPassword.Text;
                builder.InitialCatalog = textBoxDatabaseName.Text;
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    try
                    {
                        connection.Open();
                        MessageBox.Show("Connected successfully", "Message",
         MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    catch (SqlException)
                    {
                        MessageBox.Show("Cannot connect to database", "Message",
          MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }

            }

        }

        private void OkButton_Click(object sender, EventArgs e)
        {

            builder.DataSource = textBoxDataSource.Text;
            builder.UserID = textBoxUserName.Text;
            builder.Password = textBoxPassword.Text;
            builder.InitialCatalog = textBoxDatabaseName.Text;
            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();
                    Properties.Settings.Default["datasource"] = textBoxDataSource.Text;
                    Properties.Settings.Default["username"] = textBoxUserName.Text;
                    Properties.Settings.Default["password"] = textBoxPassword.Text;
                    Properties.Settings.Default["database"] = textBoxDatabaseName.Text;
                    Properties.Settings.Default.Save();
                    MessageBox.Show("Connected successfully", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch (SqlException)
                {
                    MessageBox.Show("Cannot connect to database", "Message",
      MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }



        }


        private void checkEmptyValue()
        {

            if (textBoxDataSource.Text.Equals("") || textBoxDatabaseName.Text.Equals("") || textBoxPassword.Text.Equals("") || textBoxUserName.Text.Equals(""))
            {
                okButton.Enabled = false;
            }
            else
            {
                okButton.Enabled = true;
            }

        }



        private void CancelButton_Click(object sender, EventArgs e)
        {

        }

        private void TextBoxDataSource_TextChanged(object sender, EventArgs e)
        {
            checkEmptyValue();
        }

        private void TextBoxUserName_TextChanged(object sender, EventArgs e)
        {
            checkEmptyValue();
        }

        private void TextBoxPassword_TextChanged(object sender, EventArgs e)
        {
            checkEmptyValue();
        }

        private void TextBoxDatabaseName_TextChanged(object sender, EventArgs e)
        {
            checkEmptyValue();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
    }
}
