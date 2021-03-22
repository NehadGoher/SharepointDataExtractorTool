using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContentTypeExtractor
{
    public partial class connectToOnlineForm : Form
    {
        string url = string.Empty;
        string username = string.Empty;
        string password = string.Empty;
        Form1 form =null;
        public connectToOnlineForm()
        {
            InitializeComponent();
            this.txtbx_url.Text = "https://mylinkdev.sharepoint.com/sites/DDADMS";
            this.txtbx_username.Text = "nehad.goher@mylinkdev.onmicrosoft.com";
            this.txtbx_password.Text = "Neh@d123";
            this.label_status.Text = "No Connection";
        }

        private void btn_login_Click(object sender, EventArgs e)
        {
            if (ValidateInputs())
            {
                this.label_status.Text = "Connecting .....";
                SharePointOnlineManager manager = new SharePointOnlineManager(url);
                if (manager.ConnectToSharePoint(username,password,out string res))
                {
                    this.Hide();
                    form = new Form1(manager);
                    form.ShowDialog();
                    this.Close();
                }
                else
                {
                    this.label_status.ForeColor = Color.Red;
                    this.label_status.Text = res;
                } 
            }
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(this.txtbx_url.Text))
            {
                MessageBox.Show("Please Enter valid URL");
                return false;
            }
            url = this.txtbx_url.Text.Trim();
            if (string.IsNullOrWhiteSpace(this.txtbx_username.Text))
            {
                MessageBox.Show("Please Enter valid username");
                return false;
            }
            username = this.txtbx_username.Text.Trim();
            if (string.IsNullOrWhiteSpace(this.txtbx_username.Text))
            {
                MessageBox.Show("Please Enter valid password");
                return false;
            }
            password = this.txtbx_password.Text.Trim();
            return true;
        }
    }
}
