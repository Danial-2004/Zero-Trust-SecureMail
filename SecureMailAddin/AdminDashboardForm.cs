using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using SecureMailAddin.Keycloak;
using SecureMailAddin.ZeroTrust;

namespace SecureMailAddin
{
    public class adminDashboardForm : Form
    {
        private KeycloakService _kcService;
        private string _userAccessToken;

        private ListBox lstPendingUsers;
        private Button btnApprove, btnReject;
        private Label lblStatus;

        public adminDashboardForm()
        {
            InitializeComponent();
        }

        public adminDashboardForm(KeycloakService kcService, string userToken) : this()
        {
            _kcService = kcService;
            _userAccessToken = userToken;
        }

        private void InitializeComponent()
        {
            this.lstPendingUsers = new System.Windows.Forms.ListBox();
            this.btnApprove = new System.Windows.Forms.Button();
            this.btnReject = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // lstPendingUsers
            this.lstPendingUsers.FormattingEnabled = true;
            this.lstPendingUsers.ItemHeight = 16;
            this.lstPendingUsers.Location = new System.Drawing.Point(16, 15);
            this.lstPendingUsers.Margin = new System.Windows.Forms.Padding(4);
            this.lstPendingUsers.Name = "lstPendingUsers";
            this.lstPendingUsers.Size = new System.Drawing.Size(465, 180);
            this.lstPendingUsers.TabIndex = 0;
            this.lstPendingUsers.SelectedIndexChanged += new System.EventHandler(this.lstPendingUsers_SelectedIndexChanged);

            // btnApprove
            this.btnApprove.Location = new System.Drawing.Point(16, 215);
            this.btnApprove.Margin = new System.Windows.Forms.Padding(4);
            this.btnApprove.Name = "btnApprove";
            this.btnApprove.Size = new System.Drawing.Size(133, 37);
            this.btnApprove.TabIndex = 1;
            this.btnApprove.Text = "Approve";
            this.btnApprove.UseVisualStyleBackColor = true;
            this.btnApprove.Click += new System.EventHandler(this.BtnApprove_ClickAsync);

            // btnReject
            this.btnReject.Location = new System.Drawing.Point(173, 215);
            this.btnReject.Margin = new System.Windows.Forms.Padding(4);
            this.btnReject.Name = "btnReject";
            this.btnReject.Size = new System.Drawing.Size(133, 37);
            this.btnReject.TabIndex = 2;
            this.btnReject.Text = "Reject";
            this.btnReject.UseVisualStyleBackColor = true;
            this.btnReject.Click += new System.EventHandler(this.BtnReject_ClickAsync);

            // lblStatus
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(16, 271);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 16);
            this.lblStatus.TabIndex = 3;

            // adminDashboardForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(507, 320);
            this.Controls.Add(this.lstPendingUsers);
            this.Controls.Add(this.btnApprove);
            this.Controls.Add(this.btnReject);
            this.Controls.Add(this.lblStatus);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "adminDashboardForm";
            this.Text = "Admin Dashboard";
            this.Load += new System.EventHandler(this.adminDashboardForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void adminDashboardForm_Load(object sender, EventArgs e)
        {
            if (this.DesignMode)
                return;

            LoadPendingUsersAsync();
            LogRolesAsync();
        }

        // =========================
        // THREAD-SAFE UI HELPERS
        // =========================
        private void SetStatusSafe(string text, Color color)
        {
            if (this.IsDisposed || lblStatus == null || lblStatus.IsDisposed)
                return;

            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    if (!this.IsDisposed && lblStatus != null && !lblStatus.IsDisposed)
                    {
                        lblStatus.ForeColor = color;
                        lblStatus.Text = text;
                    }
                }));
            }
            else
            {
                lblStatus.ForeColor = color;
                lblStatus.Text = text;
            }
        }

        private void SetButtonsEnabledSafe(bool approveEnabled, bool rejectEnabled)
        {
            if (this.IsDisposed)
                return;

            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    if (!this.IsDisposed)
                    {
                        btnApprove.Enabled = approveEnabled;
                        btnReject.Enabled = rejectEnabled;
                    }
                }));
            }
            else
            {
                btnApprove.Enabled = approveEnabled;
                btnReject.Enabled = rejectEnabled;
            }
        }

        private void ClearPendingListSafe()
        {
            if (this.IsDisposed || lstPendingUsers == null || lstPendingUsers.IsDisposed)
                return;

            if (lstPendingUsers.InvokeRequired)
            {
                lstPendingUsers.Invoke(new Action(() =>
                {
                    if (!this.IsDisposed && lstPendingUsers != null && !lstPendingUsers.IsDisposed)
                    {
                        lstPendingUsers.Items.Clear();
                    }
                }));
            }
            else
            {
                lstPendingUsers.Items.Clear();
            }
        }

        private void AddPendingListItemSafe(ListBoxItem item)
        {
            if (item == null || this.IsDisposed || lstPendingUsers == null || lstPendingUsers.IsDisposed)
                return;

            if (lstPendingUsers.InvokeRequired)
            {
                lstPendingUsers.Invoke(new Action(() =>
                {
                    if (!this.IsDisposed && lstPendingUsers != null && !lstPendingUsers.IsDisposed)
                    {
                        lstPendingUsers.Items.Add(item);
                    }
                }));
            }
            else
            {
                lstPendingUsers.Items.Add(item);
            }
        }

        private void RemoveListItemByValueSafe(string id)
        {
            if (this.IsDisposed || lstPendingUsers == null || lstPendingUsers.IsDisposed)
                return;

            if (lstPendingUsers.InvokeRequired)
            {
                lstPendingUsers.Invoke(new Action(() => RemoveListItemByValue(id)));
            }
            else
            {
                RemoveListItemByValue(id);
            }
        }

        private int GetPendingItemCountSafe()
        {
            if (this.IsDisposed || lstPendingUsers == null || lstPendingUsers.IsDisposed)
                return 0;

            if (lstPendingUsers.InvokeRequired)
            {
                return (int)lstPendingUsers.Invoke(new Func<int>(() => lstPendingUsers.Items.Count));
            }

            return lstPendingUsers.Items.Count;
        }

        private object GetSelectedListItemSafe()
        {
            if (this.IsDisposed || lstPendingUsers == null || lstPendingUsers.IsDisposed)
                return null;

            if (lstPendingUsers.InvokeRequired)
            {
                return lstPendingUsers.Invoke(new Func<object>(() => lstPendingUsers.SelectedItem));
            }

            return lstPendingUsers.SelectedItem;
        }

        // =========================
        // LOAD PENDING USERS
        // =========================
        private async void LoadPendingUsersAsync()
        {
            try
            {
                SetButtonsEnabledSafe(false, false);
                ClearPendingListSafe();
                SetStatusSafe("Loading pending users...", Color.Gray);

                await Task.Delay(50);

                var pendingUsers = PendingUsersRequests.PendingUsers;

                if (pendingUsers == null || pendingUsers.Count == 0)
                {
                    SetStatusSafe("No pending users.", Color.Green);
                    UpdateActionButtons();
                    return;
                }

                foreach (var user in pendingUsers)
                {
                    if (user == null)
                        continue;

                    AddPendingListItemSafe(new ListBoxItem
                    {
                        Display = $"{user.Username} ({user.Email})",
                        Value = user.Id
                    });
                }

                SetStatusSafe("", Color.Gray);
                UpdateActionButtons();
            }
            catch (Exception ex)
            {
                SetStatusSafe("Error loading users: " + ex.Message, Color.Red);
                UpdateActionButtons();
            }
        }

        // =========================
        // BUTTON STATE
        // =========================
        private void UpdateActionButtons()
        {
            if (this.IsDisposed)
                return;

            if (this.InvokeRequired)
            {
                this.Invoke(new Action(UpdateActionButtons));
                return;
            }

            bool hasRequests = lstPendingUsers.Items.Count > 0;
            bool hasSelection = lstPendingUsers.SelectedItem != null;

            btnApprove.Enabled = hasRequests && hasSelection;
            btnReject.Enabled = hasRequests && hasSelection;
        }

        // =========================
        // SELECTED USER HELPER
        // =========================
        private PendingUserRequest GetSelectedPendingUser()
        {
            if (GetPendingItemCountSafe() == 0)
            {
                SetStatusSafe("No pending requests available.", Color.Gray);
                return null;
            }

            var selected = GetSelectedListItemSafe() as ListBoxItem;
            if (selected == null)
            {
                SetStatusSafe("Select a request first.", Color.DarkOrange);
                return null;
            }

            var user = PendingUsersRequests.PendingUsers.Find(u => u.Id == selected.Value);
            if (user == null)
            {
                SetStatusSafe("Selected request no longer exists.", Color.Red);
                RemoveListItemByValueSafe(selected.Value);
                UpdateActionButtons();
                return null;
            }

            return user;
        }

        // =========================
        // APPROVE
        // =========================
        private async void BtnApprove_ClickAsync(object sender, EventArgs e)
        {
            SetButtonsEnabledSafe(false, false);

            try
            {
                if (_kcService == null)
                {
                    SetStatusSafe("Keycloak service is not initialized.", Color.Red);
                    return;
                }

                var zeroTrust = new ZeroTrustHandler();
                var decision = await zeroTrust.AuthorizeAsync(
                    _userAccessToken,
                    ProtectedActionType.ApprovePendingUser);

                if (!decision.IsAllowed)
                {
                    SetStatusSafe("Access denied: " + decision.Reason, Color.Red);
                    return;
                }

                var user = GetSelectedPendingUser();
                if (user == null)
                    return;

                SetStatusSafe("Checking existing Keycloak account...", Color.Gray);

                bool accountExists = await _kcService.UserExistsByEmailAsync(user.Email);

                if (accountExists)
                {
                    SetStatusSafe(
                        "Account already exists for this email. Only reject/discard is allowed.",
                        Color.DarkOrange);

                    MessageBox.Show(
                        $"A Keycloak account already exists for email:\n\n{user.Email}\n\nThis request cannot be approved again. Please reject/discard it.",
                        "Duplicate Account Detected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                SetStatusSafe("Approving user...", Color.Gray);

                await _kcService.RegisterUserAsync(
                    user.Username,
                    user.Email,
                    user.Password,
                    user.FirstName,
                    user.LastName,
                    user.EmailVerified);

                PendingUsersRequests.RemovePendingUser(user.Id);
                RemoveListItemByValueSafe(user.Id);

                Logger.LogPendingUserApproved("admin", user.Username);

                SetStatusSafe($"User {user.Username} approved successfully.", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusSafe("Approval failed: " + ex.Message, Color.Red);
            }
            finally
            {
                UpdateActionButtons();
            }
        }

        // =========================
        // REJECT
        // =========================
        private async void BtnReject_ClickAsync(object sender, EventArgs e)
        {
            SetButtonsEnabledSafe(false, false);

            try
            {
                var zeroTrust = new ZeroTrustHandler();
                var decision = await zeroTrust.AuthorizeAsync(
                    _userAccessToken,
                    ProtectedActionType.RejectPendingUser);

                if (!decision.IsAllowed)
                {
                    SetStatusSafe("Access denied: " + decision.Reason, Color.Red);
                    return;
                }

                var user = GetSelectedPendingUser();
                if (user == null)
                    return;

                var confirm = MessageBox.Show(
                    $"Do you want to reject/discard this request?\n\nUsername: {user.Username}\nEmail: {user.Email}",
                    "Confirm Reject",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirm != DialogResult.Yes)
                {
                    SetStatusSafe("Reject cancelled.", Color.Gray);
                    return;
                }

                PendingUsersRequests.RemovePendingUser(user.Id);
                RemoveListItemByValueSafe(user.Id);

                Logger.LogPendingUserRejected("admin", user.Username);

                SetStatusSafe($"Request for {user.Username} rejected successfully.", Color.Green);
            }
            catch (Exception ex)
            {
                SetStatusSafe("Reject failed: " + ex.Message, Color.Red);
            }
            finally
            {
                UpdateActionButtons();
            }
        }

        // =========================
        // REMOVE ITEM
        // =========================
        private void RemoveListItemByValue(string id)
        {
            ListBoxItem found = null;

            foreach (var item in lstPendingUsers.Items)
            {
                var listItem = item as ListBoxItem;
                if (listItem != null && listItem.Value == id)
                {
                    found = listItem;
                    break;
                }
            }

            if (found != null)
                lstPendingUsers.Items.Remove(found);

            if (lstPendingUsers.Items.Count == 0)
            {
                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = "No pending users.";
            }
        }

        private void lstPendingUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateActionButtons();
        }

        // =========================
        // DEBUG ROLE LOGGING
        // =========================
        private async void LogRolesAsync()
        {
            try
            {
                if (_kcService == null || string.IsNullOrWhiteSpace(_userAccessToken))
                    return;

                var ui = await _kcService.GetUserInfoWithUserTokenAsync(_userAccessToken);

                if (ui.realm_access?.roles != null)
                {
                    Debug.WriteLine("Realm Roles:");
                    foreach (var r in ui.realm_access.roles)
                        Debug.WriteLine(" - " + r);
                }

                if (ui.resource_access != null)
                {
                    Debug.WriteLine("Client Roles:");
                    foreach (var entry in ui.resource_access)
                    {
                        Debug.WriteLine($"Client = {entry.Key}");
                        if (entry.Value.roles != null)
                        {
                            foreach (var r in entry.Value.roles)
                                Debug.WriteLine("   - " + r);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error fetching roles: " + ex.Message);
            }
        }
    }

    public class ListBoxItem
    {
        public string Display { get; set; }
        public string Value { get; set; }

        public override string ToString()
        {
            return Display;
        }
    }
}