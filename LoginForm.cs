using System.Drawing;
using System.Windows.Forms;

namespace EmployeeLeaveApp {
    public class LoginForm : Form {
        readonly TextBox txtUser = new() { Width = 250, Font = new Font("Segoe UI", 14F) };
        readonly TextBox txtPass = new() { Width = 250, PasswordChar = '●', Font = new Font("Segoe UI", 14F) };
        public string AuthenticatedUser { get; private set; } = "";

        public LoginForm() {
            Text = "تسجيل الدخول للنظام"; Width = 500; Height = 320;
            StartPosition = FormStartPosition.CenterScreen; RightToLeft = RightToLeft.Yes; RightToLeftLayout = true;
            Font = new Font("Segoe UI", 14F); FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;

            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding(35), AutoSize = true };
            panel.Controls.Add(new Label { Text = "اسم المستخدم:", AutoSize = true, Font = new Font("Segoe UI", 14F, FontStyle.Bold) });
            panel.Controls.Add(txtUser);
            panel.Controls.Add(new Label { Text = "كلمة المرور:", AutoSize = true, Margin = new Padding(0, 15, 0, 0), Font = new Font("Segoe UI", 14F, FontStyle.Bold) });
            panel.Controls.Add(txtPass);

            var btnLogin = new Button { Text = "دخول", Width = 140, Height = 48, Margin = new Padding(0, 25, 0, 0), BackColor = Color.FromArgb(41, 128, 185), ForeColor = Color.White, Font = new Font("Segoe UI", 14F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += (_, _) => {
                string u = txtUser.Text.Trim();
                string p = txtPass.Text.Trim();
                if ((u == AppConfig.User1Name && p == AppConfig.User1Pass) || (u == AppConfig.User2Name && p == AppConfig.User2Pass)) {
                    AuthenticatedUser = u;
                    DialogResult = DialogResult.OK;
                    Close();
                } else {
                    MessageBox.Show("اسم المستخدم أو كلمة المرور غير صحيحة.", "خطأ في تسجيل الدخول", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            panel.Controls.Add(btnLogin);
            Controls.Add(panel);
        }
    }
}