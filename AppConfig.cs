namespace EmployeeLeaveApp {
    public static class AppConfig {
        public const string ExcelFilePath = @"C:\Users\Omarpc\Desktop\EmployeeServices2.xlsx";
        public const string CompanyName = "شركة مصر للتعمير - إدارة نظم المعلومات";
        public const int MaxExitPermitsPerMonth = 2;
        public const int MaxMissionsPerMonth = 2;
        
        // بيانات المستخدمين
        public const string User1Name = "user1";
        public const string User1Pass = "123";
        public const string User2Name = "user2";
        public const string User2Pass = "456";

        public const string DashboardSheet = "لوحة المؤشرات";
        public const string EmployeesSheet = "الموظفون";
        public const string BalancesSheet = "أرصدة الإجازات";
        public const string LeavesSheet = "طلبات الإجازات";
        public const string ExitPermitsSheet = "تصاريح الخروج";
        public const string MissionsSheet = "المأموريات";
        public static readonly string[] LeaveTypes = { "إجازة اعتيادية", "إجازة مرضية" };
    }
}