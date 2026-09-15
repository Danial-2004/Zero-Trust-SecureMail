using SecureMailAddin.Keycloak;
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecureMailAddin
{
    public class RegistrationForm : Form
    {
        private Panel cardPanel;
        private TextBox txtUsername, txtEmail, txtPassword, txtConfirm, txtFirstName, txtLastName;
        private CheckBox chkEmailVerified;
        private Button btnRegister, btnShowPassword, btnShowConfirm;
        private Label lblStatus;
        private bool isSubmitting = false;

        public RegistrationForm()
        {
            InitializeComponentCustom();
            LoadOutlookEmail();
        }
        // Add these public properties to expose the registration fields
        public string Username => txtUsername.Text.Trim();
        public string Email => txtEmail.Text.Trim();
        public string Password => txtPassword.Text;
        public string FirstName => txtFirstName.Text.Trim();
        public string LastName => txtLastName.Text.Trim();
        public bool EmailVerified => chkEmailVerified.Checked;
        private void InitializeComponentCustom()
        {
            this.Text = "SecureMail - Registration";
            this.Size = new Size(450, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.BackColor = Color.FromArgb(245, 246, 250);

            cardPanel = new Panel
            {
                Size = new Size(380, 520),
                Location = new Point((this.ClientSize.Width - 380) / 2, (this.ClientSize.Height - 520) / 2),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(cardPanel);

            int y = 20;
            txtFirstName = CreateTextBox("First Name", y); cardPanel.Controls.Add(txtFirstName); y += 50;
            txtLastName = CreateTextBox("Last Name", y); cardPanel.Controls.Add(txtLastName); y += 50;
            txtUsername = CreateTextBox("Username", y); cardPanel.Controls.Add(txtUsername); y += 50;
            txtEmail = CreateTextBox("Email", y); cardPanel.Controls.Add(txtEmail); y += 50;

            txtPassword = CreateTextBox("Password", y, true);
            cardPanel.Controls.Add(txtPassword);
            btnShowPassword = new Button
            {
                Text = "👁",
                Font = new Font("Segoe UI Emoji", 10),
                Size = new Size(30, 28),
                Location = new Point(310, y),
                FlatStyle = FlatStyle.Flat
            };
            btnShowPassword.FlatAppearance.BorderSize = 0;
            btnShowPassword.Click += (s, e) => txtPassword.UseSystemPasswordChar = !txtPassword.UseSystemPasswordChar;
            cardPanel.Controls.Add(btnShowPassword);
            y += 50;

            txtConfirm = CreateTextBox("Confirm Password", y, true);
            cardPanel.Controls.Add(txtConfirm);
            btnShowConfirm = new Button
            {
                Text = "👁",
                Font = new Font("Segoe UI Emoji", 10),
                Size = new Size(30, 28),
                Location = new Point(310, y),
                FlatStyle = FlatStyle.Flat
            };
            btnShowConfirm.FlatAppearance.BorderSize = 0;
            btnShowConfirm.Click += (s, e) => txtConfirm.UseSystemPasswordChar = !txtConfirm.UseSystemPasswordChar;
            cardPanel.Controls.Add(btnShowConfirm);
            y += 50;

            chkEmailVerified = new CheckBox
            {
                Text = "Email Verified",
                Location = new Point(20, y),
                AutoSize = true
            };
            cardPanel.Controls.Add(chkEmailVerified);
            y += 40;

            lblStatus = new Label
            {
                Text = "",
                ForeColor = Color.Red,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, y),
                Width = cardPanel.Width
            };
            cardPanel.Controls.Add(lblStatus);
            y += 40;

            btnRegister = new Button
            {
                Text = "Register",
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                Size = new Size(260, 40),
                Location = new Point(60, y),
                FlatStyle = FlatStyle.Flat
            };
            btnRegister.FlatAppearance.BorderSize = 0;
            btnRegister.Click += BtnRegister_ClickAsync;
            cardPanel.Controls.Add(btnRegister);
        }

        private TextBox CreateTextBox(string placeholder, int y, bool isPassword = false)
        {
            var tb = new TextBox
            {
                Font = new Font("Segoe UI", 10),
                Size = new Size(260, 28),
                Location = new Point(60, y),
                ForeColor = Color.Gray,
                Text = placeholder,
                UseSystemPasswordChar = false,
                BorderStyle = BorderStyle.FixedSingle
            };

            tb.Enter += (s, e) =>
            {
                if (tb.ForeColor == Color.Gray)
                {
                    tb.Text = "";
                    tb.ForeColor = Color.Black;
                    if (isPassword) tb.UseSystemPasswordChar = true;
                }
            };
            tb.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tb.Text))
                {
                    tb.Text = placeholder;
                    tb.ForeColor = Color.Gray;
                    if (isPassword) tb.UseSystemPasswordChar = false;
                }
            };

            return tb;
        }

        private void LoadOutlookEmail()
        {
            try
            {
                Outlook.Application outlookApp = null;
                try
                {
                    outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                catch
                {
                    outlookApp = new Outlook.Application();
                }

                Outlook.Accounts accounts = outlookApp.Session.Accounts;

                if (accounts.Count == 0)
                {
                    MessageBox.Show("No Outlook account found. Please sign in to Outlook first.");
                    return;
                }

                txtEmail.Text = accounts[1].SmtpAddress;
                txtEmail.ForeColor = Color.Black;

                txtEmail.ReadOnly = true;
                txtEmail.Enabled = false;
                txtEmail.KeyPress += (s, e) => e.Handled = true;
                txtEmail.ContextMenu = new ContextMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to load Outlook email:\n\n" + ex.Message);
            }
        }

        private bool ValidateFields(out string errorMessage)
        {
            errorMessage = "";

            if (!Regex.IsMatch(txtFirstName.Text.Trim(), @"^[A-Za-z]+(?: [A-Za-z]+)*$"))
            {
                errorMessage = "First Name must contain only letters and spaces.";
                return false;
            }
            if (!Regex.IsMatch(txtLastName.Text.Trim(), @"^[A-Za-z]+(?: [A-Za-z]+)*$"))
            {
                errorMessage = "Last Name must contain only letters and spaces.";
                return false;
            }
            if (!Regex.IsMatch(txtUsername.Text.Trim(), @"^[A-Za-z0-9_\-@.]{3,20}$"))
            {
                errorMessage = "Username can contain letters, numbers, _, -, @ and . (3-20 chars).";
                return false;
            }
            if (!Regex.IsMatch(txtEmail.Text.Trim(), @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                errorMessage = "Invalid email format.";
                return false;
            }
            if (txtPassword.Text.Length < 6 ||
                !Regex.IsMatch(txtPassword.Text, @"[A-Z]") ||
                !Regex.IsMatch(txtPassword.Text, @"[a-z]") ||
                !Regex.IsMatch(txtPassword.Text, @"[0-9]"))
            {
                errorMessage = "Password must be at least 6 chars with uppercase, lowercase, and a number.";
                return false;
            }
            if (txtPassword.Text != txtConfirm.Text)
            {
                errorMessage = "Passwords do not match.";
                return false;
            }
            return true;
        }

        private void BtnRegister_ClickAsync(object sender, EventArgs e)
        {
            if (isSubmitting) return;
            isSubmitting = true;

            lblStatus.ForeColor = Color.Gray;
            lblStatus.Text = "Submitting for admin approval...";
            btnRegister.Enabled = false;

            string username = txtUsername.ForeColor == Color.Gray ? "" : txtUsername.Text.Trim();
            string email = txtEmail.ForeColor == Color.Gray ? "" : txtEmail.Text.Trim();
            string password = txtPassword.ForeColor == Color.Gray ? "" : txtPassword.Text;
            string confirm = txtConfirm.ForeColor == Color.Gray ? "" : txtConfirm.Text;
            string firstName = txtFirstName.ForeColor == Color.Gray ? "" : txtFirstName.Text.Trim();
            string lastName = txtLastName.ForeColor == Color.Gray ? "" : txtLastName.Text.Trim();
            bool emailVerified = chkEmailVerified.Checked;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(email) ||
                string.IsNullOrEmpty(password) || string.IsNullOrEmpty(confirm))
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = "Please fill in all required fields.";
                btnRegister.Enabled = true;
                isSubmitting = false;
                return;
            }

            if (!ValidateFields(out string validationError))
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = validationError;
                btnRegister.Enabled = true;
                isSubmitting = false;
                return;
            }

            var newRequest = new PendingUserRequest
            {
                Username = username,
                Email = email,
                Password = password,
                FirstName = firstName,
                LastName = lastName,
                EmailVerified = emailVerified
            };

            try
            {
                if (!PendingUsersRequests.AddPendingUser(newRequest))
                {
                    lblStatus.ForeColor = Color.Red;
                    lblStatus.Text = "A pending request or account with this username/email already exists.";
                    btnRegister.Enabled = true;
                    isSubmitting = false;
                    return;
                }

                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = "Registration submitted. Waiting for admin approval.";
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = "Submission failed: " + ex.Message;
            }
            finally
            {
                btnRegister.Enabled = true;
                isSubmitting = false;
            }
        }
    }
}
