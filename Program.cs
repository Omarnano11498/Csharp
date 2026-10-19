using System;
using System.Windows.Forms;

namespace EmployeeLeaveApp {
    static class Program {
        [STAThread]
        static void Main() {
            ApplicationConfiguration.Initialize();
            
            while (true) {
                using (var login = new LoginForm()) {
                    if (login.ShowDialog() == DialogResult.OK) {
                        string user = login.AuthenticatedUser;
                        using (var main = new MainForm(user)) {
                            Application.Run(main);
                            if (!main.SwitchUserRequested) {
                                break;
                            }
                        }
                    } else {
                        break;
                    }
                }
            }
        }
    }
}