using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tara_app.Services;

namespace Tara_app
{
    public partial class AuthForm : Form
    {
        public AuthForm()
        {
            InitializeComponent();
            UserName.Text = "Username";
            Password.Text = "Password";
        }


        private void AuthForm_Load(object sender, EventArgs e)
        {
            AuthError.Visible = false;
            Password.UseSystemPasswordChar = false;
        }

        private void AuthUser_Enter(object sender, EventArgs e)
        {
            if (UserName.Text == "Username")
            {
                UserName.Text = "";
            }
        }

        private void AuthPassword_Enter(object sender, EventArgs e)
        {
            if (Password.Text == "Password")
            {
                Password.Text = "";
            }

            Password.UseSystemPasswordChar = true;
        }

        private void Login_Click(object sender, EventArgs e)
        {
            AuthHandler.Instance.authenticateUser(UserName.Text, Password.Text);
        }

        public void ShowError(string errorText)
        {
            AuthError.Text = errorText;
            AuthError.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
