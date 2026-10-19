using System;

namespace EmployeeLeaveApp {
    public sealed class EmployeeInfo {
        public string Code { get; set; } = "";
        public string Name { get; set; } = "";
        public string Department { get; set; } = "";
    }

    public sealed class EmployeeQuery {
        public EmployeeInfo? Employee;
        public decimal RegularLeaveBalance;
        public decimal SickLeaveBalance;
        public int ExitPermitsThisMonth;
        public int MissionsThisMonth;
    }

    public sealed class LeaveRecordModel {
        public int RowIndex { get; set; }
        public string كود_الموظف { get; set; } = "";
        public string اسم_الموظف { get; set; } = "";
        public string نوع_الإجازة { get; set; } = "";
        public string تاريخ_البداية { get; set; } = "";
        public string تاريخ_النهاية { get; set; } = "";
        public int عدد_الأيام { get; set; }
        public string تم_استلام_المرفق { get; set; } = "";
    }

    public sealed class PermitRecordModel {
        public int RowIndex { get; set; }
        public string كود_الموظف { get; set; } = "";
        public string اسم_الموظف { get; set; } = "";
        public string التاريخ { get; set; } = "";
        public string من_الساعة { get; set; } = "";
        public string إلى_الساعة { get; set; } = "";
        public string السبب { get; set; } = "";
        public string تم_استلام_المرفق { get; set; } = "";
    }

    public sealed class MissionRecordModel {
        public int RowIndex { get; set; }
        public string كود_الموظف { get; set; } = "";
        public string اسم_الموظف { get; set; } = "";
        public string التاريخ { get; set; } = "";
        public string من_الساعة { get; set; } = "";
        public string إلى_الساعة { get; set; } = "";
        public string المكان { get; set; } = "";
        public string السبب { get; set; } = "";
        public string تم_استلام_المرفق { get; set; } = "";
    }
}