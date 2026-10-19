using System;
using System.Drawing;
using System.Windows.Forms;

namespace EmployeeLeaveApp {
    public class MainForm : Form {
        readonly IExcelService excel = new ExcelService(AppConfig.ExcelFilePath);
        readonly string loggedUser;
        public bool SwitchUserRequested { get; private set; } = false;
        
        readonly TextBox code = new() { MaxLength = 4, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox name = new() { ReadOnly = true, Width = 280, Font = new Font("Segoe UI", 16F) };
        readonly TextBox dept = new() { ReadOnly = true, Width = 280, Font = new Font("Segoe UI", 16F) };
        readonly ComboBox leaveType = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker leaveStart = new() { Format = DateTimePickerFormat.Short, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker leaveEnd = new() { Format = DateTimePickerFormat.Short, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly NumericUpDown leaveDays = new() { Minimum = 1, Maximum = 365, ReadOnly = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox balance = new() { ReadOnly = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox newBalance = new() { ReadOnly = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox leaveImage = new() { ReadOnly = true, Width = 350, Font = new Font("Segoe UI", 14F) };
        
        readonly DateTimePicker permitDate = new() { Format = DateTimePickerFormat.Short, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker permitFrom = new() { Format = DateTimePickerFormat.Time, ShowUpDown = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker permitTo = new() { Format = DateTimePickerFormat.Time, ShowUpDown = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox permitReason = new() { Width = 350, Font = new Font("Segoe UI", 16F) };
        readonly TextBox permitImage = new() { ReadOnly = true, Width = 350, Font = new Font("Segoe UI", 14F) };
        
        readonly DateTimePicker missionDate = new() { Format = DateTimePickerFormat.Short, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker missionFrom = new() { Format = DateTimePickerFormat.Time, ShowUpDown = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly DateTimePicker missionTo = new() { Format = DateTimePickerFormat.Time, ShowUpDown = true, Width = 240, Font = new Font("Segoe UI", 16F) };
        readonly TextBox missionLocation = new() { Width = 350, Font = new Font("Segoe UI", 16F) };
        readonly TextBox missionReason = new() { Width = 350, Font = new Font("Segoe UI", 16F) };
        readonly TextBox missionImage = new() { ReadOnly = true, Width = 350, Font = new Font("Segoe UI", 14F) };
        
        readonly ComboBox querySelector = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 550, Font = new Font("Segoe UI", 16F) };

        Button btnMax = null!;

        public MainForm(string username) {
            loggedUser = username;
            Text = AppConfig.CompanyName + " - (المستخدم الحالي: " + username + ")";
            Width = 1250; Height = 1000; MinimumSize = new Size(950, 750);
            StartPosition = FormStartPosition.CenterScreen; RightToLeft = RightToLeft.No; RightToLeftLayout = false;
            AutoScaleMode = AutoScaleMode.Inherit; Font = new Font("Segoe UI", 16F);
            FormBorderStyle = FormBorderStyle.None;
            BuildCustomTitleBar();
            Build();
        }

        void BuildCustomTitleBar() {
            var headerContainer = new Panel { Dock = DockStyle.Top, AutoSize = true, BackColor = Color.FromArgb(28, 40, 51) };
            var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(21, 32, 40), Padding = new Padding(10) };
            
            var btnClose = new Button { Text = "✕ إغلاق النظام", Width = 160, Height = 42, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(192, 57, 43), Font = new Font("Segoe UI", 13F, FontStyle.Bold) };
            btnClose.FlatAppearance.BorderSize = 0; btnClose.Click += (_, _) => Close();

            var btnSwitchUser = new Button { Text = "🔄 تبديل المستخدم", Width = 180, Height = 42, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(211, 84, 0), Font = new Font("Segoe UI", 13F, FontStyle.Bold) };
            btnSwitchUser.FlatAppearance.BorderSize = 0;
            btnSwitchUser.Click += (_, _) => { SwitchUserRequested = true; Close(); };

            btnMax = new Button { Text = "🗖 تكبير", Width = 140, Height = 42, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(41, 128, 185), Font = new Font("Segoe UI", 13F, FontStyle.Bold) };
            btnMax.FlatAppearance.BorderSize = 0; btnMax.Click += (_, _) => { WindowState = WindowState == FormWindowState.Maximized ? FormWindowState.Normal : FormWindowState.Maximized; };

            var btnMin = new Button { Text = "🗕 تصغير", Width = 140, Height = 42, FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.FromArgb(39, 174, 96), Font = new Font("Segoe UI", 13F, FontStyle.Bold) };
            btnMin.FlatAppearance.BorderSize = 0; btnMin.Click += (_, _) => { WindowState = FormWindowState.Minimized; };

            btnPanel.Controls.Add(btnClose);
            btnPanel.Controls.Add(btnSwitchUser);
            btnPanel.Controls.Add(btnMax);
            btnPanel.Controls.Add(btnMin);

            var titleBar = new Panel { Dock = DockStyle.Top, Height = 65, BackColor = Color.FromArgb(28, 40, 51), Padding = new Padding(15) };
            var titleLabel = new Label { Text = Text, ForeColor = Color.White, Font = new Font("Segoe UI", 16F, FontStyle.Bold), AutoSize = true, Location = new Point(300, 16) };
            titleBar.Controls.Add(titleLabel);

            headerContainer.Controls.Add(titleBar); headerContainer.Controls.Add(btnPanel);

            bool dragging = false; Point dragCursorPoint = Point.Empty; Point formLocation = Point.Empty;
            void OnMouseDown(object? s, MouseEventArgs e) { if (e.Button == MouseButtons.Left) { dragging = true; dragCursorPoint = Cursor.Position; formLocation = Location; } }
            void OnMouseMove(object? s, MouseEventArgs e) { if (dragging) { Point diff = Point.Subtract(Cursor.Position, new Size(dragCursorPoint)); Location = Point.Add(formLocation, new Size(diff)); } }
            void OnMouseUp(object? s, MouseEventArgs e) { dragging = false; }

            headerContainer.MouseDown += OnMouseDown; headerContainer.MouseMove += OnMouseMove; headerContainer.MouseUp += OnMouseUp;
            titleBar.MouseDown += OnMouseDown; titleBar.MouseMove += OnMouseMove; titleBar.MouseUp += OnMouseUp;
            titleLabel.MouseDown += OnMouseDown; titleLabel.MouseMove += OnMouseMove; titleLabel.MouseUp += OnMouseUp;
            btnPanel.MouseDown += OnMouseDown; btnPanel.MouseMove += OnMouseMove; btnPanel.MouseUp += OnMouseUp;

            Resize += (_, _) => { if (btnMax != null) { btnMax.Text = WindowState == FormWindowState.Maximized ? "🗗 استعادة" : "🗖 تكبير"; } };
            Controls.Add(headerContainer);
        }

        Label L(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(10, 15, 10, 10), Anchor = AnchorStyles.Right, Font = new Font("Segoe UI", 16F, FontStyle.Bold) };
        Button B(string text, EventHandler handler) {
            var b = new Button { Text = text, AutoSize = true, MinimumSize = new Size(180, 50), Margin = new Padding(10), Padding = new Padding(10), Font = new Font("Segoe UI", 15F, FontStyle.Bold), BackColor = Color.FromArgb(41, 128, 185), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            b.FlatAppearance.BorderSize = 0;
            b.Click += handler; return b;
        }
        
        void Build() {
            var root = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(40) };
            var tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 16F, FontStyle.Bold) };
            
            if (loggedUser == AppConfig.User1Name) {
                var top = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(80) };
                top.Controls.Add(L("كود الموظف")); top.Controls.Add(code); top.Controls.Add(L("الاسم")); top.Controls.Add(name); top.Controls.Add(L("الإدارة")); top.Controls.Add(dept);
                code.KeyPress += (_, e) => { if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b') e.Handled = true; };
                code.Leave += (_, _) => LoadEmployee();
                
                tabs.TabPages.Add(LeaveTab()); tabs.TabPages.Add(PermitTab()); tabs.TabPages.Add(MissionTab()); tabs.TabPages.Add(QueryTabForUser1());
                root.Controls.Add(tabs); root.Controls.Add(top);
            } else {
                tabs.TabPages.Add(ApprovalTabForUser2());
                root.Controls.Add(tabs);
            }
            Controls.Add(root);
        }

        TableLayoutPanel Table(int rows) {
            var f = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 4, RowCount = rows, Padding = new Padding(20), RightToLeft = RightToLeft.Yes };
            f.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15)); f.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            f.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15)); f.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            for (int i = 0; i < rows; i++) f.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return f;
        }
        
        void R(TableLayoutPanel f, int r, string a, Control x, string b, Control y) {
            foreach (var c in new[] { x, y }) { c.Dock = DockStyle.Fill; c.Margin = new Padding(8); c.MinimumSize = new Size(120, 45); }
            f.Controls.Add(L(a), 0, r); f.Controls.Add(x, 1, r); f.Controls.Add(L(b), 2, r); f.Controls.Add(y, 3, r);
        }
        
        TabPage Page(string title, Control content) {
            var p = new TabPage(title) { AutoScroll = true, Padding = new Padding(15) }; content.Dock = DockStyle.Top; p.Controls.Add(content); return p;
        }
        
        TabPage LeaveTab() {
            var f = Table(5); leaveType.Items.AddRange(AppConfig.LeaveTypes); leaveType.SelectedIndex = 0;
            leaveStart.ValueChanged += LeaveChanged; leaveEnd.ValueChanged += LeaveChanged; leaveType.SelectedIndexChanged += LeaveChanged;
            R(f, 0, "نوع الإجازة", leaveType, "بداية الإجازة", leaveStart); R(f, 1, "نهاية الإجازة", leaveEnd, "عدد الأيام", leaveDays);
            R(f, 2, "الرصيد قبل الطلب", balance, "الرصيد بعد الطلب", newBalance); R(f, 3, "المرفق", leaveImage, "", new Panel());
            f.Controls.Add(B("اختيار صورة", (_, _) => Pick(leaveImage)), 1, 4); f.Controls.Add(B("تسجيل الإجازة", SubmitLeave), 3, 4);
            return Page("الإجازات", f);
        }
        
        TabPage PermitTab() {
            var f = Table(4); R(f, 0, "التاريخ", permitDate, "من الساعة", permitFrom); R(f, 1, "إلى الساعة", permitTo, "السبب", permitReason);
            R(f, 2, "المرفق", permitImage, "", new Panel()); f.Controls.Add(B("اختيار صورة", (_, _) => Pick(permitImage)), 1, 3);
            f.Controls.Add(B("تسجيل التصريح", SubmitPermit), 3, 3); return Page("تصاريح الخروج", f);
        }
        
        TabPage MissionTab() {
            var f = Table(4); R(f, 0, "التاريخ", missionDate, "من الساعة", missionFrom); R(f, 1, "إلى الساعة", missionTo, "المكان", missionLocation);
            R(f, 2, "السبب", missionReason, "المرفق", missionImage); f.Controls.Add(B("اختيار صورة", (_, _) => Pick(missionImage)), 1, 3);
            f.Controls.Add(B("تسجيل المأمورية", SubmitMission), 3, 3); return Page("المأموريات", f);
        }

        TabPage QueryTabForUser1() {
            var panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(30), AutoSize = true };
            querySelector.Items.AddRange(new[] { "الملخص الشامل", "سجل الإجازات بالتواريخ", "سجل تصاريح الخروج بالتواريخ", "سجل المأموريات بالتواريخ" });
            querySelector.SelectedIndex = 0;
            panel.Controls.Add(L("نوع الاستعلام المطلوب")); panel.Controls.Add(querySelector);
            panel.Controls.Add(B("فتح شاشة الاستعلام", (_, _) => OpenQueryScreen()));
            panel.Controls.Add(B("تسجيل الخروج", (_, _) => { Close(); }));
            return Page("الاستعلامات", panel);
        }

        TabPage ApprovalTabForUser2() {
    // لوحة رئيسية تملأ الصفحة وتسمح بالتمرير عند الحاجة
    var outerPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };

    // جدول لتنظيم العناصر فوق بعضها وتوسيطها في منتصف الشاشة بدقة
    var table = new TableLayoutPanel {
        ColumnCount = 1,
        RowCount = 4,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        RightToLeft = RightToLeft.Yes
    };
    table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

    // العنوان الرئيسي
    var lblHeader = new Label { 
        Text = "لوحة تحكم المستخدم الثاني - مراجعة المرفقات والتقارير", 
        Font = new Font("Segoe UI", 20F, FontStyle.Bold), 
        AutoSize = true, 
        Anchor = AnchorStyles.None,
        Margin = new Padding(0, 10, 0, 35) 
    };

    // زر استعلام المرفقات غير المستلمة
    var btnOpenQueryModal = new Button { 
        Text = "فتح شاشة استعلام مراجعة جميع المرفقات غير المستلمة", 
        Width = 650, Height = 90, 
        BackColor = Color.FromArgb(41, 128, 185), 
        ForeColor = Color.White, 
        Font = new Font("Segoe UI", 15F, FontStyle.Bold), 
        FlatStyle = FlatStyle.Flat,
        Anchor = AnchorStyles.None,
        Margin = new Padding(0, 0, 0, 20)
    };
    btnOpenQueryModal.FlatAppearance.BorderSize = 0;
    btnOpenQueryModal.Click += (_, _) => {
        using var modalForm = new SecondUserQueryForm(excel);
        modalForm.ShowDialog();
    };

    // زر تقارير الموظفين بفترة
    var btnOpenReportModal = new Button { 
        Text = "طباعة وعرض تقارير الموظفين ", 
        Width = 650, Height = 90, 
        BackColor = Color.FromArgb(39, 174, 96), 
        ForeColor = Color.White, 
        Font = new Font("Segoe UI", 15F, FontStyle.Bold), 
        FlatStyle = FlatStyle.Flat,
        Anchor = AnchorStyles.None,
        Margin = new Padding(0, 0, 0, 20)
    };
    btnOpenReportModal.FlatAppearance.BorderSize = 0;
    btnOpenReportModal.Click += (_, _) => {
        using var reportForm = new SecondUserReportForm(excel);
        reportForm.ShowDialog();
    };

    // زر تسجيل الخروج
    var btnLogout = new Button { 
        Text = "تسجيل الخروج", 
        Width = 250, Height = 55, 
        BackColor = Color.FromArgb(192, 57, 43), 
        ForeColor = Color.White, 
        Font = new Font("Segoe UI", 14F, FontStyle.Bold), 
        FlatStyle = FlatStyle.Flat,
        Anchor = AnchorStyles.None,
        Margin = new Padding(0, 0, 0, 10)
    };
    btnLogout.FlatAppearance.BorderSize = 0;
    btnLogout.Click += (_, _) => Close();

    // إضافة العناصر إلى الجدول (تحت بعضها البعض)
    table.Controls.Add(lblHeader, 0, 0);
    table.Controls.Add(btnOpenQueryModal, 0, 1);
    table.Controls.Add(btnOpenReportModal, 0, 2);
    table.Controls.Add(btnLogout, 0, 3);

    // دالة لتوسيط الجدول في منتصف الشاشة أفقياً وعمودياً بشكل دائم
    void CenterTable() {
        table.Location = new Point(
            Math.Max(0, (outerPanel.Width - table.Width) / 2),
            Math.Max(0, (outerPanel.Height - table.Height) / 2)
        );
    }

    outerPanel.Resize += (_, _) => CenterTable();
    outerPanel.ControlAdded += (_, _) => CenterTable();

    outerPanel.Controls.Add(table);
    
    var tabPage = new TabPage("إدارة المرفقات والتقارير للمستخدم الثاني");
    tabPage.Controls.Add(outerPanel);
    
    // ضبط الموضع عند التحميل الأول
    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 10 };
    timer.Tick += (s, e) => {
        timer.Stop();
        CenterTable();
    };
    timer.Start();

    return tabPage;
}

        void OpenQueryScreen() {
            if (!ReadyForQueryOnly()) return;
            string selected = querySelector.SelectedItem?.ToString() ?? "الملخص الشامل";
            using var queryForm = new QueryForm(code.Text, selected, excel);
            queryForm.ShowDialog();
        }

        int CalculateWorkingDays(DateTime start, DateTime end) {
            if (end.Date < start.Date) return 0;
            int count = 0;
            for (var d = start.Date; d <= end.Date; d = d.AddDays(1)) {
                if (d.DayOfWeek != DayOfWeek.Friday && d.DayOfWeek != DayOfWeek.Saturday) count++;
            }
            return Math.Max(1, count);
        }

        void LoadEmployee() {
            if (code.Text.Length != 4) return; 
            var e = excel.GetEmployee(code.Text);
            if (e == null) { name.Clear(); dept.Clear(); balance.Clear(); newBalance.Clear(); return; }
            
            name.Text = e.Name; 
            dept.Text = e.Department; 
            LeaveChanged(null, EventArgs.Empty);

            if (excel.HasUnreceivedAttachments(code.Text)) {
                MessageBox.Show("تنبيه أمني هام: يوجد طلب سابق لهذا الموظف لم يتم استلام مرفقه من المستخدم الثاني.\nمتاح لك الاستعلام فقط ولا يمكن تسجيل طلبات جديدة حتى يتم الاعتماد.", "تنبيه أمني", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        void LeaveChanged(object? sender, EventArgs e) {
            if (leaveEnd.Value.Date < leaveStart.Value.Date) return;
            int days = CalculateWorkingDays(leaveStart.Value.Date, leaveEnd.Value.Date); leaveDays.Value = Math.Clamp(days, 1, 365);
            if (code.Text.Length == 4) { decimal b = excel.GetLeaveBalance(code.Text, leaveType.Text); balance.Text = b.ToString("N0"); newBalance.Text = Math.Max(0, b - days).ToString("N0"); }
        }
        
        bool Ready() {
            if (code.Text.Length != 4) { MessageBox.Show("أدخل كود الموظف المكون من 4 أرقام.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return false; }
            if (excel.HasUnreceivedAttachments(code.Text)) {
                MessageBox.Show("لا يمكن التسجيل! يوجد مرفق سابق غير معتمد لهذا الموظف (متاح الاستعلام فقط).", "منع التسجيل", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (excel.GetEmployee(code.Text) == null) { MessageBox.Show("كود الموظف غير موجود بالنظام.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; } 
            return true;
        }

        bool ReadyForQueryOnly() {
            if (code.Text.Length != 4) { MessageBox.Show("أدخل كود الموظف المكون من 4 أرقام للاستعلام.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return false; }
            if (excel.GetEmployee(code.Text) == null) { MessageBox.Show("كود الموظف غير موجود.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; } 
            return true;
        }
        
        void SubmitLeave(object? s, EventArgs e) {
            if (!Ready()) return; int days = CalculateWorkingDays(leaveStart.Value.Date, leaveEnd.Value.Date);
            if (days < 1) { MessageBox.Show("تاريخ نهاية الإجازة يجب ألا يسبق تاريخ البداية.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            decimal current = excel.GetLeaveBalance(code.Text, leaveType.Text);
            if (days > current) { MessageBox.Show($"رصيد {leaveType.Text} غير كافٍ. الرصيد المتاح: {current:N0} يوم.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var emp = excel.GetEmployee(code.Text)!; decimal after = current - days;
            excel.SaveLeaveRequest(emp.Code, emp.Name, emp.Department, leaveType.Text, leaveStart.Value.Date, leaveEnd.Value.Date, days, current, after, Environment.MachineName, DateTime.Now, leaveImage.Text);
            MessageBox.Show("تم تسجيل الإجازة بنجاح (بحالة قيد الانتظار لموافقة المستخدم الثاني).", "نتيجة التسجيل", MessageBoxButtons.OK, MessageBoxIcon.Information); 
            LoadEmployee();
        }
        
        void SubmitPermit(object? s, EventArgs e) {
            if (!Ready()) return;
            if (permitTo.Value.TimeOfDay <= permitFrom.Value.TimeOfDay) { MessageBox.Show("وقت النهاية يجب أن يكون بعد البداية.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (string.IsNullOrWhiteSpace(permitReason.Text)) { MessageBox.Show("أدخل سبب التصريح.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (excel.GetMonthlyExitPermitCount(code.Text, permitDate.Value) >= AppConfig.MaxExitPermitsPerMonth) { MessageBox.Show("تم استنفاد تصريحَي الشهر المسموح بهما.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var x = excel.GetEmployee(code.Text)!;
            excel.SaveExitPermit(x.Code, x.Name, x.Department, permitDate.Value.Date, permitFrom.Value.TimeOfDay, permitTo.Value.TimeOfDay, permitReason.Text, Environment.MachineName, DateTime.Now, permitImage.Text);
            MessageBox.Show("تم تسجيل التصريح بنجاح.", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        void SubmitMission(object? s, EventArgs e) {
            if (!Ready()) return;
            if (missionTo.Value.TimeOfDay <= missionFrom.Value.TimeOfDay) { MessageBox.Show("وقت النهاية يجب أن يكون بعد البداية.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (string.IsNullOrWhiteSpace(missionLocation.Text) || string.IsNullOrWhiteSpace(missionReason.Text)) { MessageBox.Show("أدخل المكان والسبب.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var x = excel.GetEmployee(code.Text)!;
            excel.SaveMission(x.Code, x.Name, x.Department, missionDate.Value.Date, missionFrom.Value.TimeOfDay, missionTo.Value.TimeOfDay, missionLocation.Text, missionReason.Text, Environment.MachineName, DateTime.Now, missionImage.Text);
            MessageBox.Show("تم تسجيل المأمورية بنجاح.", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        void Pick(TextBox t) { 
            using var d = new OpenFileDialog { Filter = "صور (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp" }; 
            if (d.ShowDialog() == DialogResult.OK) t.Text = d.FileName; 
        }
    }

    // شاشة استعلام المستخدم الثاني (المرفقات المعلقة)
    public class SecondUserQueryForm : Form {
        readonly DataGridView grid = new() { Dock = DockStyle.Fill, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = false, ReadOnly = true };
        readonly IExcelService _excel;
        readonly ComboBox cmbType = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 320, Font = new Font("Segoe UI", 14F) };

        public SecondUserQueryForm(IExcelService excel) {
            _excel = excel;
            Text = "شاشة استعلام ومراجعة المرفقات (المستخدم الثاني)";
            Width = 1350; Height = 780;
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes; RightToLeftLayout = true;
            Font = new Font("Segoe UI", 14F);
            FormBorderStyle = FormBorderStyle.Sizable;

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 90, Padding = new Padding(15), BackColor = Color.FromArgb(240, 243, 244) };
            
            cmbType.Items.AddRange(new[] { "طلبات الإجازات غير المستلمة", "تصاريح الخروج غير المستلمة", "المأموريات غير المستلمة" });
            cmbType.SelectedIndex = 0;
            cmbType.SelectedIndexChanged += (_, _) => LoadData();

            var closeBtn = new Button { Text = "إغلاق الشاشة", Width = 160, Height = 48, BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, Font = new Font("Segoe UI", 14F, FontStyle.Bold) };
            closeBtn.Click += (_, _) => Close();

            var flowTop = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Left, AutoSize = true, BackColor = Color.FromArgb(240, 243, 244) };
            flowTop.Controls.Add(closeBtn);

            var flowRight = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Right, AutoSize = true, BackColor = Color.FromArgb(240, 243, 244) };
            flowRight.Controls.Add(new Label { Text = "نوع السجل:", AutoSize = true, Margin = new Padding(5, 12, 5, 0), Font = new Font("Segoe UI", 14F, FontStyle.Bold) });
            flowRight.Controls.Add(cmbType);

            topPanel.Controls.Add(flowTop);
            topPanel.Controls.Add(flowRight);

            Controls.Add(grid);
            Controls.Add(topPanel);

            grid.RowTemplate.Height = 45;
            grid.ColumnHeadersHeight = 48;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            grid.CellContentClick += Grid_CellContentClick;

            LoadData();
        }

        void LoadData() {
            grid.Columns.Clear();
            if (cmbType.SelectedIndex == 0) grid.DataSource = _excel.GetAllPendingLeaves();
            else if (cmbType.SelectedIndex == 1) grid.DataSource = _excel.GetAllPendingPermits();
            else grid.DataSource = _excel.GetAllPendingMissions();

            var btnCol = new DataGridViewButtonColumn {
                Name = "BtnAction",
                HeaderText = "إجراء الاستلام",
                Text = "استلام المرفق",
                UseColumnTextForButtonValue = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                Width = 160
            };
            grid.Columns.Add(btnCol);
        }

        void Grid_CellContentClick(object? sender, DataGridViewCellEventArgs e) {
            if (e.ColumnIndex == grid.Columns["BtnAction"]?.Index && e.RowIndex >= 0) {
                int rowIndex = e.RowIndex;
                var rowObj = grid.Rows[rowIndex].DataBoundItem;

                if (cmbType.SelectedIndex == 0 && rowObj is LeaveRecordModel l) {
                    _excel.UpdateAttachmentStatus(AppConfig.LeavesSheet, l.RowIndex, "تم الاستلام");
                } else if (cmbType.SelectedIndex == 1 && rowObj is PermitRecordModel p) {
                    _excel.UpdateAttachmentStatus(AppConfig.ExitPermitsSheet, p.RowIndex, "تم الاستلام");
                } else if (cmbType.SelectedIndex == 2 && rowObj is MissionRecordModel m) {
                    _excel.UpdateAttachmentStatus(AppConfig.MissionsSheet, m.RowIndex, "تم الاستلام");
                }

                MessageBox.Show("تم تأكيد الاستلام بنجاح لهذا البند.", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadData();
            }
        }
    }

    // شاشة تقارير الموظفين (من فترة إلى فترة) مع دعم عدم إجبار كود الموظف لعرض الكل
    public class SecondUserReportForm : Form {
        readonly TextBox txtCode = new() { MaxLength = 10, Width = 160, Font = new Font("Segoe UI", 14F) };
        readonly DateTimePicker dtpFrom = new() { Format = DateTimePickerFormat.Short, Width = 160, Font = new Font("Segoe UI", 14F) };
        readonly DateTimePicker dtpTo = new() { Format = DateTimePickerFormat.Short, Width = 160, Font = new Font("Segoe UI", 14F) };
        readonly ComboBox cmbReportType = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240, Font = new Font("Segoe UI", 14F) };
        readonly DataGridView grid = new() { Dock = DockStyle.Fill, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = false, ReadOnly = true };
        readonly IExcelService _excel;

        public SecondUserReportForm(IExcelService excel) {
            _excel = excel;
            Text = "شاشة تقارير الموظفين بفترة (من - إلى) - المستخدم الثاني";
            Width = 1300; Height = 780;
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes; RightToLeftLayout = true;
            Font = new Font("Segoe UI", 14F);
            FormBorderStyle = FormBorderStyle.Sizable;

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 100, Padding = new Padding(15), BackColor = Color.FromArgb(240, 243, 244) };
            
            cmbReportType.Items.AddRange(new[] { "تقرير الإجازات", "تقرير تصاريح الخروج", "تقرير المأموريات" });
            cmbReportType.SelectedIndex = 0;

            var btnSearch = new Button { Text = "بحث واستعلام", Width = 150, Height = 48, BackColor = Color.FromArgb(41, 128, 185), ForeColor = Color.White, Font = new Font("Segoe UI", 14F, FontStyle.Bold) };
            btnSearch.Click += (_, _) => SearchData();

            var btnPrint = new Button { Text = "🖨 طباعة التقرير", Width = 160, Height = 48, BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, Font = new Font("Segoe UI", 14F, FontStyle.Bold) };
            btnPrint.Click += (_, _) => MessageBox.Show("تم إرسال التقرير للطباعة بنجاح.", "طباعة", MessageBoxButtons.OK, MessageBoxIcon.Information);

            var closeBtn = new Button { Text = "إغلاق", Width = 120, Height = 48, Font = new Font("Segoe UI", 14F, FontStyle.Bold) };
            closeBtn.Click += (_, _) => Close();

            var flowPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true, BackColor = Color.FromArgb(240, 243, 244) };
            flowPanel.Controls.Add(new Label { Text = "كود الموظف (فارغ للكل):", AutoSize = true, Margin = new Padding(5, 12, 5, 0), Font = new Font("Segoe UI", 13F, FontStyle.Bold) });
            flowPanel.Controls.Add(txtCode);
            flowPanel.Controls.Add(new Label { Text = "من تاريخ:", AutoSize = true, Margin = new Padding(10, 12, 5, 0), Font = new Font("Segoe UI", 13F, FontStyle.Bold) });
            flowPanel.Controls.Add(dtpFrom);
            flowPanel.Controls.Add(new Label { Text = "إلى تاريخ:", AutoSize = true, Margin = new Padding(10, 12, 5, 0), Font = new Font("Segoe UI", 13F, FontStyle.Bold) });
            flowPanel.Controls.Add(dtpTo);
            flowPanel.Controls.Add(new Label { Text = "نوع التقرير:", AutoSize = true, Margin = new Padding(10, 12, 5, 0), Font = new Font("Segoe UI", 13F, FontStyle.Bold) });
            flowPanel.Controls.Add(cmbReportType);
            flowPanel.Controls.Add(btnSearch);
            flowPanel.Controls.Add(btnPrint);
            flowPanel.Controls.Add(closeBtn);

            topPanel.Controls.Add(flowPanel);
            Controls.Add(grid);
            Controls.Add(topPanel);

            grid.RowTemplate.Height = 40;
            grid.ColumnHeadersHeight = 48;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        }

        void SearchData() {
            string code = txtCode.Text.Trim();
            // تم إلغاء الإجبار: إذا ترك فارغاً أو كتب علامة، يعرض كل السجِلّات بكل سلاسة.
            DateTime from = dtpFrom.Value;
            DateTime to = dtpTo.Value;

            if (cmbReportType.SelectedIndex == 0) {
                grid.DataSource = _excel.GetEmployeeLeavesByDateRange(code, from, to);
            } else if (cmbReportType.SelectedIndex == 1) {
                grid.DataSource = _excel.GetEmployeeExitPermitsByDateRange(code, from, to);
            } else {
                grid.DataSource = _excel.GetEmployeeMissionsByDateRange(code, from, to);
            }
        }
    }

    public class QueryForm : Form {
        readonly DataGridView grid = new() { Dock = DockStyle.Fill, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, AllowUserToAddRows = false, ReadOnly = true };
        public QueryForm(string employeeCode, string queryType, IExcelService excel) {
            Text = "شاشة الاستعلام - " + queryType;
            Width = 1150; Height = 680; StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes; RightToLeftLayout = true; Font = new Font("Segoe UI", 15F);
            
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(15), BackColor = Color.FromArgb(240, 243, 244) };
            var lblTitle = new Label { Text = "نتائج الاستعلام: " + queryType + " (للموظف: " + employeeCode + ")", Font = new Font("Segoe UI", 16F, FontStyle.Bold), AutoSize = true, Location = new Point(20, 20) };
            var closeBtn = new Button { Text = "إغلاق الشاشة", Width = 160, Height = 45, Dock = DockStyle.Left, Font = new Font("Segoe UI", 14F, FontStyle.Bold) };
            closeBtn.Click += (_, _) => Close();
            
            topPanel.Controls.Add(lblTitle); topPanel.Controls.Add(closeBtn);
            Controls.Add(grid); Controls.Add(topPanel);
            
            grid.RowTemplate.Height = 40; grid.ColumnHeadersHeight = 48;
            
            if (queryType == "سجل الإجازات بالتواريخ") grid.DataSource = excel.GetEmployeeLeaves(employeeCode);
            else if (queryType == "سجل تصاريح الخروج بالتواريخ") grid.DataSource = excel.GetEmployeeExitPermits(employeeCode);
            else if (queryType == "سجل المأموريات بالتواريخ") grid.DataSource = excel.GetEmployeeMissions(employeeCode);
            else {
                var q = excel.Query(employeeCode, DateTime.Today);
                grid.DataSource = new[] { new { اسم_الموظف = q.Employee!.Name, الإدارة = q.Employee.Department, رصيد_الاعتيادية = q.RegularLeaveBalance, رصيد_المرضية = q.SickLeaveBalance, تصاريح_الشهر = q.ExitPermitsThisMonth, مأموريات_الشهر = q.MissionsThisMonth } };
            }
        }
    }
}