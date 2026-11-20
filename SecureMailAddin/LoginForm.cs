using Newtonsoft.Json;
using SecureMailAddin.Keycloak;
using SecureMailAddin.ZeroTrust;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SecureMailAddin
{
    public partial class LoginForm : Form
    {
        private Panel cardPanel;
        private PictureBox logoBox;
        private Label titleLabel;
        private TextBox txtUser;
        private TextBox txtPass;
        private CheckBox chkRemember;
        private Button btnLogin;
        private Button btnadminLogin;
        private Button btnRegisterUser;
        private Label lblStatus;
        private Button btnShowPass;
        private Button btnClose;
        private Button btnGenerateKeys;
        private LinkLabel lnkForgotPassword;
        private bool _allowProgrammaticClose = false;



        private KeycloakService _kcService;
        private readonly string rememberFile =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SecureMailUser.txt");

        public AuthResult AuthResult { get; private set; }
        public static bool AllowCloseOnShutdown { get; set; } = false;

        public LoginForm()
        {
            InitializeComponent();
            InitializeComponentCustom();
            _kcService = new KeycloakService();
            LoadRememberedUser();
        }

        private void InitializeComponentCustom()
        {
            this.ControlBox = false;
            this.Text = "SecureMail - Login";
            this.Size = new Size(420, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.BackColor = Color.FromArgb(245, 246, 250);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            btnClose = new Button
            {
                Text = "✖",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Size = new Size(30, 30),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.DarkRed,
                BackColor = Color.White,
                Location = new Point(this.Width - 60, 10),
                TabStop = false
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnClose.Click += BtnClose_Click;
            this.Controls.Add(btnClose);

            cardPanel = new Panel
            {
                Size = new Size(340, 520),
                Location = new Point((this.ClientSize.Width - 340) / 2, (this.ClientSize.Height - 520) / 2),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(cardPanel);

            logoBox = new PictureBox
            {
                Size = new Size(70, 70),
                Location = new Point((cardPanel.Width - 70) / 2, 30),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White,
                Image = SystemIcons.Shield.ToBitmap()
            };
            cardPanel.Controls.Add(logoBox);

            titleLabel = new Label
            {
                Text = "SECUREMAIL LOGIN",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, 110),
                Width = cardPanel.Width,
                ForeColor = Color.FromArgb(40, 40, 40)
            };
            cardPanel.Controls.Add(titleLabel);

            txtUser = new TextBox
            {
                Font = new Font("Segoe UI", 10),
                Size = new Size(260, 28),
                Location = new Point(40, 150),
                ForeColor = Color.Gray,
                Text = "Enter your username",
                BorderStyle = BorderStyle.FixedSingle
            };
            txtUser.Enter += RemoveUsernamePlaceholder;
            txtUser.Leave += AddUsernamePlaceholder;
            txtUser.TextChanged += (s, e) => CheckIfKeysExist();
            cardPanel.Controls.Add(txtUser);

            txtPass = new TextBox
            {
                Font = new Font("Segoe UI", 10),
                Size = new Size(230, 28),
                Location = new Point(40, 190),
                ForeColor = Color.Gray,
                Text = "Enter your password",
                UseSystemPasswordChar = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            txtPass.Enter += RemovePasswordPlaceholder;
            txtPass.Leave += AddPasswordPlaceholder;
            cardPanel.Controls.Add(txtPass);

            btnShowPass = new Button
            {
                Text = "👁",
                Font = new Font("Segoe UI Emoji", 10),
                Size = new Size(30, 28),
                Location = new Point(270, 190),
                FlatStyle = FlatStyle.Flat
            };
            btnShowPass.FlatAppearance.BorderSize = 0;
            btnShowPass.Click += (s, e) =>
            {
                if (txtPass.ForeColor != Color.Gray)
                    txtPass.UseSystemPasswordChar = !txtPass.UseSystemPasswordChar;
            };
            cardPanel.Controls.Add(btnShowPass);

            chkRemember = new CheckBox
            {
                Text = "Remember me",
                Location = new Point(40, 235),
                AutoSize = true
            };
            cardPanel.Controls.Add(chkRemember);

            btnLogin = new Button
            {
                Text = "LOGIN",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                Size = new Size(260, 40),
                Location = new Point(40, 270),
                FlatStyle = FlatStyle.Flat
            };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += BtnLogin_ClickAsync;
            cardPanel.Controls.Add(btnLogin);

            btnadminLogin = new Button
            {
                Text = "Admin LOGIN",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(204, 51, 0),
                ForeColor = Color.White,
                Size = new Size(260, 35),
                Location = new Point(40, 320),
                FlatStyle = FlatStyle.Flat
            };
            btnadminLogin.FlatAppearance.BorderSize = 0;
            btnadminLogin.Click += BtnadminLogin_ClickAsync;
            cardPanel.Controls.Add(btnadminLogin);

            btnRegisterUser = new Button
            {
                Text = "Register New User",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 153, 51),
                ForeColor = Color.White,
                Size = new Size(260, 35),
                Location = new Point(40, 365),
                FlatStyle = FlatStyle.Flat
            };
            btnRegisterUser.FlatAppearance.BorderSize = 0;
            btnRegisterUser.Click += BtnRegisterUser_Click;
            cardPanel.Controls.Add(btnRegisterUser);

            lblStatus = new Label
            {
                Text = "",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Red,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, 455),
                Width = cardPanel.Width
            };
            cardPanel.Controls.Add(lblStatus);

            lnkForgotPassword = new LinkLabel
            {
                Text = "Forgot Password?",
                Location = new Point(180, 235),
                AutoSize = true,
                LinkColor = Color.FromArgb(0, 102, 204)
            };
            lnkForgotPassword.Click += LnkForgotPassword_Click;
            cardPanel.Controls.Add(lnkForgotPassword);

            btnGenerateKeys = new Button
            {
                Text = "Generate Keys",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(102, 51, 153),
                ForeColor = Color.White,
                Size = new Size(260, 35),
                Location = new Point(40, 410),
                FlatStyle = FlatStyle.Flat
            };
            btnGenerateKeys.FlatAppearance.BorderSize = 0;
            btnGenerateKeys.Click += BtnGenerateKeys_Click;
            cardPanel.Controls.Add(btnGenerateKeys);

            CheckIfKeysExist();
        }

        // ---------------------------
        // FORM OPEN HELPERS
        // ---------------------------
        private static bool IsFormOpen<T>() where T : Form
        {
            return Application.OpenForms.OfType<T>().Any();
        }

        private void SetButtonsEnabled(bool enabled)
        {
            btnLogin.Enabled = enabled;
            btnadminLogin.Enabled = enabled;
            btnRegisterUser.Enabled = enabled;
            btnGenerateKeys.Enabled = enabled || CryptoHelper.HasKeyPairFor(txtUser.ForeColor == Color.Gray ? "" : txtUser.Text.Trim());
        }

        // ---------------------------
        // USER LOGIN
        // ---------------------------
        private async void BtnLogin_ClickAsync(object sender, EventArgs e)
        {
            await DoUserLoginAsync();
        }

        private async Task DoUserLoginAsync()
        {
            string username = "";

            try
            {
                btnLogin.Enabled = false;
                lblStatus.ForeColor = Color.Gray;
                lblStatus.Text = "Logging in...";

                username = txtUser.ForeColor == Color.Gray ? "" : txtUser.Text.Trim();
                string password = txtPass.ForeColor == Color.Gray ? "" : txtPass.Text;

                Logger.LogUserLoginStarted(username);

                if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "Enter username and password.";
                    return;
                }

                var auth = await _kcService.LoginAsync(username, password);

                if (auth == null || string.IsNullOrEmpty(auth.AccessToken))
                    throw new Exception("Invalid credentials.");

                this.AuthResult = auth;
                SaveRememberedUser(username);

                Logger.LogUserLoginSucceeded(username);

                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = "Login successful.";

                _allowProgrammaticClose = true;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                Logger.LogUserLoginFailed(username, ex.Message);
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = "Login failed: " + ex.Message;
            }
            finally
            {
                btnLogin.Enabled = true;
            }
        }

        // ---------------------------
        // ADMIN LOGIN
        // ---------------------------
        private async void BtnadminLogin_ClickAsync(object sender, EventArgs e)
        {
            if (IsFormOpen<adminLoginForm>() || IsFormOpen<adminDashboardForm>())
                return;

            try
            {
                Logger.LogAdminLoginStarted();

                SetButtonsEnabled(false);

                using (var adminLoginForm = new adminLoginForm(_kcService))
                {
                    var result = adminLoginForm.ShowDialog(this);

                    if (result != DialogResult.OK)
                    {
                        SetButtonsEnabled(true);
                        return;
                    }

                    string token = adminLoginForm.AccessToken;

                    if (string.IsNullOrWhiteSpace(token))
                    {
                        SetButtonsEnabled(true);

                        MessageBox.Show(
                            "Admin login token is missing.",
                            "Login Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }

                    var zeroTrust = new ZeroTrustHandler();
                    var decision = await zeroTrust.AuthorizeAsync(token, ProtectedActionType.OpenAdminDashboard);

                    if (!decision.IsAllowed)
                    {
                        SetButtonsEnabled(true);

                        Logger.LogAdminDashboardDenied("unknown", decision.Reason);

                        MessageBox.Show(
                            "Access denied.\n\n" + decision.Reason,
                            "Zero Trust Access Control",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    var ui = await _kcService.GetUserInfoWithUserTokenAsync(token);

                    Logger.LogAdminLoginSucceeded(ui?.preferred_username);
                    Logger.LogAdminDashboardOpened(ui?.preferred_username);

                    this.AuthResult = new AuthResult
                    {
                        AccessToken = token,
                        TokenType = "Bearer"
                    };

                    var adminForm = new adminDashboardForm(_kcService, token);
                    adminForm.Show();

                    _allowProgrammaticClose = true;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                SetButtonsEnabled(true);

                Logger.LogAdminDashboardDenied("unknown", ex.Message);

                MessageBox.Show(
                    "Failed to open admin dashboard.\n\n" + ex.Message,
                    "SecureMail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ---------------------------
        // REGISTER USER
        // ---------------------------
        private void BtnRegisterUser_Click(object sender, EventArgs e)
        {
            if (IsFormOpen<RegistrationForm>())
                return;

            try
            {
                SetButtonsEnabled(false);
                this.Hide();

                using (var regForm = new RegistrationForm())
                {
                    var result = regForm.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        var newRequest = new PendingUserRequest
                        {
                            Username = regForm.Username,
                            Email = regForm.Email,
                            Password = regForm.Password,
                            FirstName = regForm.FirstName,
                            LastName = regForm.LastName,
                            EmailVerified = regForm.EmailVerified
                        };

                        PendingUsersRequests.PendingUsers.Add(newRequest);

                        MessageBox.Show(
                            "Registration request submitted successfully.\n\nIt is now pending admin approval.",
                            "Pending Approval",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                }
            }
            finally
            {
                if (!this.IsDisposed)
                {
                    this.Show();
                    SetButtonsEnabled(true);
                }
            }
        }

        // ---------------------------
        // PLACEHOLDERS
        // ---------------------------
        private void RemoveUsernamePlaceholder(object sender, EventArgs e)
        {
            if (txtUser.ForeColor == Color.Gray)
            {
                txtUser.Text = "";
                txtUser.ForeColor = Color.Black;
            }
            CheckIfKeysExist();
        }

        private void AddUsernamePlaceholder(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUser.Text))
            {
                txtUser.Text = "Enter your username";
                txtUser.ForeColor = Color.Gray;
            }
        }

        private void RemovePasswordPlaceholder(object sender, EventArgs e)
        {
            if (txtPass.ForeColor == Color.Gray)
            {
                txtPass.Text = "";
                txtPass.ForeColor = Color.Black;
                txtPass.UseSystemPasswordChar = true;
            }
        }

        private void AddPasswordPlaceholder(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPass.Text))
            {
                txtPass.UseSystemPasswordChar = false;
                txtPass.Text = "Enter your password";
                txtPass.ForeColor = Color.Gray;
            }
        }

        // ---------------------------
        // REMEMBER USER
        // ---------------------------
        private void LoadRememberedUser()
        {
            if (File.Exists(rememberFile))
            {
                string savedUser = File.ReadAllText(rememberFile);
                txtUser.Text = savedUser;
                txtUser.ForeColor = Color.Black;
                chkRemember.Checked = true;
                txtUser.BackColor = Color.LightGray;
            }
        }

        private void SaveRememberedUser(string username)
        {
            if (chkRemember.Checked)
                File.WriteAllText(rememberFile, username);
            else if (File.Exists(rememberFile))
                File.Delete(rememberFile);
        }

        // ---------------------------
        // CLOSE
        // ---------------------------
        private void BtnClose_Click(object sender, EventArgs e)
        {
            if (AllowCloseOnShutdown)
            {
                _allowProgrammaticClose = true;
                this.Close();
                return;
            }

            MessageBox.Show(
                "Login is required to use SecureMail add-in.",
                "Login Required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }

        // ---------------------------
        // KEYS
        // ---------------------------
        private void BtnGenerateKeys_Click(object sender, EventArgs e)
        {
            try
            {
                string currentUser = txtUser.ForeColor == Color.Gray ? "" : txtUser.Text.Trim();

                if (string.IsNullOrWhiteSpace(currentUser))
                {
                    MessageBox.Show("Enter your username first.", "Secure Mail");
                    return;
                }

                string folderPath = @"C:\SecureMailKeys";
                string pubPath = Path.Combine(folderPath, $"{currentUser}_public.xml");
                string privPath = Path.Combine(folderPath, $"{currentUser}_private.xml");

                Directory.CreateDirectory(folderPath);

                if (File.Exists(pubPath) || File.Exists(privPath))
                {
                    var res = MessageBox.Show(
                        "Key files already exist. Overwrite?",
                        "Confirm Overwrite",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );
                    if (res == DialogResult.No)
                        return;
                }

                CryptoHelper.GenerateAndSaveRsaKeyPair(currentUser);

                string pubXml = CryptoHelper.GetPublicKeyXml(currentUser);
                string privXml = CryptoHelper.LoadPrivateKeyXml(currentUser);
                File.WriteAllText(pubPath, pubXml);
                File.WriteAllText(privPath, privXml);

                MessageBox.Show(
                    $"RSA Key Pair generated successfully!\n\nPublic Key: {pubPath}\nPrivate Key: {privPath}",
                    "Secure Mail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to generate keys:\n" + ex.Message, "Secure Mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckIfKeysExist()
        {
            string currentUser = txtUser.ForeColor == Color.Gray ? "" : txtUser.Text.Trim();

            if (string.IsNullOrWhiteSpace(currentUser))
            {
                btnGenerateKeys.Enabled = false;
                return;
            }

            if (CryptoHelper.HasKeyPairFor(currentUser))
            {
                btnGenerateKeys.Enabled = false;
                btnGenerateKeys.Text = "Keys Already Generated";
                btnGenerateKeys.BackColor = Color.Gray;
            }
            else
            {
                btnGenerateKeys.Enabled = true;
                btnGenerateKeys.Text = "Generate Keys";
                btnGenerateKeys.BackColor = Color.FromArgb(102, 51, 153);
            }
        }

        // ---------------------------
        // FORGOT PASSWORD
        // ---------------------------
        private void LnkForgotPassword_Click(object sender, EventArgs e)
        {
            try
            {
                string url =
                    $"{KeycloakConfig.BaseUrl}/realms/{KeycloakConfig.Realm}/login-actions/reset-credentials" +
                    $"?client_id={Uri.EscapeDataString(KeycloakConfig.ClientId)}";

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });

                MessageBox.Show(
                    "A browser window has been opened for password reset.",
                    "SecureMail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Failed to open the password reset page.\n\n" + ex.Message,
                    "SecureMail",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // ---------------------------
        // FORM CLOSING
        // ---------------------------
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (AllowCloseOnShutdown || _allowProgrammaticClose)
            {
                base.OnFormClosing(e);
                return;
            }

            // Only block when the USER is trying to close the form manually
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                MessageBox.Show(
                    "You must log in to continue using Outlook.",
                    "Login Required",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            base.OnFormClosing(e);
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
        }
    }
}