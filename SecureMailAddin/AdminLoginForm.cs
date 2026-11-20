using SecureMailAddin.Keycloak;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SecureMailAddin
{
    public class adminLoginForm : Form
    {
        private KeycloakService _kcService;

        private TextBox txtUser;
        private TextBox txtPass;
        private Button btnLogin;
        private Label lblStatus;

        // Access token available after successful login
        public string AccessToken { get; private set; }

        public adminLoginForm(KeycloakService kcService)
        {
            _kcService = kcService;
            InitializeForm();
        }

        private void InitializeForm()
        {
            this.Text = "Admin Login";
            this.Size = new Size(400, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lblStatus = new Label
            {
                Location = new Point(20, 20),
                Size = new Size(340, 20),
                ForeColor = Color.Red,
                Text = ""
            };
            this.Controls.Add(lblStatus);

            // Username TextBox with manual placeholder
            txtUser = new TextBox
            {
                Location = new Point(20, 60),
                Size = new Size(340, 25),
                ForeColor = Color.Gray,
                Text = "Username"
            };
            txtUser.Enter += (s, e) =>
            {
                if (txtUser.ForeColor == Color.Gray)
                {
                    txtUser.Text = "";
                    txtUser.ForeColor = Color.Black;
                }
            };
            txtUser.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtUser.Text))
                {
                    txtUser.Text = "Username";
                    txtUser.ForeColor = Color.Gray;
                }
            };
            this.Controls.Add(txtUser);

            // Password TextBox with manual placeholder
            txtPass = new TextBox
            {
                Location = new Point(20, 100),
                Size = new Size(340, 25),
                ForeColor = Color.Gray,
                Text = "Password",
                UseSystemPasswordChar = false
            };
            txtPass.Enter += (s, e) =>
            {
                if (txtPass.ForeColor == Color.Gray)
                {
                    txtPass.Text = "";
                    txtPass.ForeColor = Color.Black;
                    txtPass.UseSystemPasswordChar = true;
                }
            };
            txtPass.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtPass.Text))
                {
                    txtPass.UseSystemPasswordChar = false;
                    txtPass.Text = "Password";
                    txtPass.ForeColor = Color.Gray;
                }
            };
            this.Controls.Add(txtPass);

            // Login button
            btnLogin = new Button
            {
                Text = "Login",
                Location = new Point(20, 150),
                Size = new Size(340, 35),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += BtnLogin_ClickAsync;
            this.Controls.Add(btnLogin);
        }

        private async void BtnLogin_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                btnLogin.Enabled = false;
                lblStatus.ForeColor = Color.Gray;
                lblStatus.Text = "Logging in...";

                string username = txtUser.ForeColor == Color.Gray ? "" : txtUser.Text.Trim();
                string password = txtPass.ForeColor == Color.Gray ? "" : txtPass.Text;

                if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "Enter credentials.";
                    btnLogin.Enabled = true;
                    return;
                }

                var auth = await _kcService.LoginAsync(username, password);
                if (auth == null || string.IsNullOrEmpty(auth.AccessToken))
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "Invalid credentials.";
                    btnLogin.Enabled = true;
                    return;
                }

                // ✅ Store access token
                AccessToken = auth.AccessToken;

                // Close form with OK result
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = "Login failed: " + ex.Message;
            }
            finally
            {
                btnLogin.Enabled = true;
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // adminLoginForm
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "adminLoginForm";
            this.Load += new System.EventHandler(this.adminLoginForm_Load);
            this.ResumeLayout(false);

        }

        private void adminLoginForm_Load(object sender, EventArgs e)
        {

        }
    }
}
