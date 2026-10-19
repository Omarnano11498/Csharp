using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EmployeeLeaveApp {
    public interface IExcelService {
        EmployeeInfo? GetEmployee(string code);
        decimal GetLeaveBalance(string code, string type);
        int GetMonthlyExitPermitCount(string code, DateTime date);
        int GetMonthlyMissionCount(string code, DateTime date);
        EmployeeQuery Query(string code, DateTime date);
        bool HasUnreceivedAttachments(string code);
        void UpdateAttachmentStatus(string sheetName, int row, string status);
        
        List<LeaveRecordModel> GetEmployeeLeaves(string code);
        List<PermitRecordModel> GetEmployeeExitPermits(string code);
        List<MissionRecordModel> GetEmployeeMissions(string code);

        List<LeaveRecordModel> GetEmployeeLeavesByDateRange(string code, DateTime from, DateTime to);
        List<PermitRecordModel> GetEmployeeExitPermitsByDateRange(string code, DateTime from, DateTime to);
        List<MissionRecordModel> GetEmployeeMissionsByDateRange(string code, DateTime from, DateTime to);

        List<LeaveRecordModel> GetAllPendingLeaves();
        List<PermitRecordModel> GetAllPendingPermits();
        List<MissionRecordModel> GetAllPendingMissions();

        void SaveLeaveRequest(string code, string name, string dept, string type, DateTime start, DateTime end, int days, decimal before, decimal after, string host, DateTime now, string? image);
        void SaveExitPermit(string code, string name, string dept, DateTime date, TimeSpan from, TimeSpan to, string reason, string host, DateTime now, string? image);
        void SaveMission(string code, string name, string dept, DateTime date, TimeSpan from, TimeSpan to, string location, string reason, string host, DateTime now, string? image);
    }

    public sealed class ExcelService : IExcelService {
        readonly string path; 
        static readonly object Sync = new();

        public ExcelService(string filePath) => path = filePath;

        static string S(IXLCell c) => c.GetString().Trim();
        
        XLWorkbook Open() {
            if (File.Exists(path)) { var existing = new XLWorkbook(path); EnsureSheets(existing); return existing; }
            var dir = Path.GetDirectoryName(path); if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            var wb = new XLWorkbook(); EnsureSheets(wb); wb.SaveAs(path); return wb;
        }
        
        static void EnsureSheet(XLWorkbook wb, string name, string[] headers) {
            var ws = wb.Worksheets.FirstOrDefault(x => x.Name == name) ?? wb.Worksheets.Add(name);
            if (ws.Cell(1, 1).IsEmpty()) for (int i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];
            ws.Row(1).Style.Font.Bold = true; ws.SheetView.FreezeRows(1);
            ws.RightToLeft = true;
        }

        static void EnsureSheets(XLWorkbook wb) {
            EnsureDashboardSheet(wb);
            EnsureSheet(wb, AppConfig.EmployeesSheet, new[] { "كود الموظف", "اسم الموظف", "الإدارة" });
            var bs = wb.Worksheets.FirstOrDefault(x => x.Name == AppConfig.BalancesSheet) ?? wb.Worksheets.Add(AppConfig.BalancesSheet);
            bs.RightToLeft = true;
            if (bs.Cell(1, 1).IsEmpty()) {
                bs.Cell(1, 1).Value = "كود الموظف"; bs.Cell(1, 2).Value = "اسم الموظف"; bs.Cell(1, 3).Value = "رصيد الإجازة الاعتيادية"; bs.Cell(1, 4).Value = "رصيد الإجازة المرضية";
            }
            bs.Row(1).Style.Font.Bold = true; bs.SheetView.FreezeRows(1);

            EnsureSheet(wb, AppConfig.LeavesSheet, new[] { "م", "كود الموظف", "اسم الموظف", "الإدارة", "نوع الإجازة", "تاريخ البداية", "تاريخ النهاية", "عدد الأيام", "الرصيد قبل الطلب", "الرصيد بعد الطلب", "اسم الجهاز", "وقت التسجيل", "المرفق", "تم استلام المرفق" });
            EnsureSheet(wb, AppConfig.ExitPermitsSheet, new[] { "م", "كود الموظف", "اسم الموظف", "الإدارة", "التاريخ", "من الساعة", "إلى الساعة", "السبب", "اسم الجهاز", "وقت التسجيل", "المرفق", "تم استلام المرفق" });
            EnsureSheet(wb, AppConfig.MissionsSheet, new[] { "م", "كود الموظف", "اسم الموظف", "الإدارة", "التاريخ", "من الساعة", "إلى الساعة", "المكان", "سبب المأمورية", "اسم الجهاز", "وقت التسجيل", "المرفق", "تم استلام المرفق" });
        }

        static void EnsureDashboardSheet(XLWorkbook wb) {
            var ws = wb.Worksheets.FirstOrDefault(x => x.Name == AppConfig.DashboardSheet) ?? wb.Worksheets.Add(AppConfig.DashboardSheet);
            ws.RightToLeft = true;
            ws.Cell("B2").Value = "لوحة المؤشرات العامة (Dashboard) - " + AppConfig.CompanyName;
            ws.Range("B2:D2").Merge();
            ws.Cell("B2").Style.Font.Bold = true; ws.Cell("B2").Style.Font.FontSize = 16;
            ws.Cell("B2").Style.Font.FontColor = XLColor.White; ws.Cell("B2").Style.Fill.BackgroundColor = XLColor.FromArgb(28, 40, 51);
            ws.Cell("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            ws.Cell("B4").Value = "المؤشر الإحصائي"; ws.Cell("C4").Value = "القيمة الإجمالية";
            ws.Range("B4:C4").Style.Font.Bold = true; ws.Range("B4:C4").Style.Fill.BackgroundColor = XLColor.FromArgb(41, 128, 185);
            ws.Range("B4:C4").Style.Font.FontColor = XLColor.White;

            ws.Cell("B5").Value = "إجمالي عدد الموظفين"; ws.Cell("C5").FormulaA1 = "=COUNTA('الموظفون'!A2:A5000)";
            ws.Cell("B6").Value = "إجمالي طلبات الإجازات"; ws.Cell("C6").FormulaA1 = "=COUNTA('طلبات الإجازات'!B2:B5000)";
            ws.Cell("B7").Value = "إجمالي تصاريح الخروج"; ws.Cell("C7").FormulaA1 = "=COUNTA('تصاريح الخروج'!B2:B5000)";
            ws.Cell("B8").Value = "إجمالي المأموريات"; ws.Cell("C8").FormulaA1 = "=COUNTA('المأموريات'!B2:B5000)";
            ws.Columns().AdjustToContents();
        }

        public bool HasUnreceivedAttachments(string code) {
            lock (Sync) using (var wb = Open()) {
                var wsLeaves = wb.Worksheet(AppConfig.LeavesSheet);
                int lastL = wsLeaves.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= lastL; r++) {
                    if (S(wsLeaves.Cell(r, 2)) == code && S(wsLeaves.Cell(r, 14)) != "تم الاستلام") return true;
                }
                var wsPermits = wb.Worksheet(AppConfig.ExitPermitsSheet);
                int lastP = wsPermits.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= lastP; r++) {
                    if (S(wsPermits.Cell(r, 2)) == code && S(wsPermits.Cell(r, 12)) != "تم الاستلام") return true;
                }
                var wsMissions = wb.Worksheet(AppConfig.MissionsSheet);
                int lastM = wsMissions.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= lastM; r++) {
                    if (S(wsMissions.Cell(r, 2)) == code && S(wsMissions.Cell(r, 13)) != "تم الاستلام") return true;
                }
            }
            return false;
        }

        public void UpdateAttachmentStatus(string sheetName, int row, string status) {
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(sheetName);
                int col = sheetName == AppConfig.LeavesSheet ? 14 : (sheetName == AppConfig.ExitPermitsSheet ? 12 : 13);
                ws.Cell(row, col).Value = status;
                wb.Save();
            }
        }

        public EmployeeInfo? GetEmployee(string code) {
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.EmployeesSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) if (S(ws.Cell(r, 1)) == code) return new EmployeeInfo { Code = code, Name = S(ws.Cell(r, 2)), Department = S(ws.Cell(r, 3)) };
                return null;
            }
        }
        
        public decimal GetLeaveBalance(string code, string type) {
            int col = type == "إجازة مرضية" ? 4 : 3;
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.BalancesSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) if (S(ws.Cell(r, 1)) == code) { decimal.TryParse(S(ws.Cell(r, col)), out var b); return b; }
                return 0;
            }
        }
        
        public int GetMonthlyExitPermitCount(string code, DateTime date) => CountMonth(AppConfig.ExitPermitsSheet, code, date, 5);
        public int GetMonthlyMissionCount(string code, DateTime date) => CountMonth(AppConfig.MissionsSheet, code, date, 5);
        
        int CountMonth(string sheet, string code, DateTime date, int dateCol) {
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(sheet); int n = 0; int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) if (S(ws.Cell(r, 2)) == code && DateTime.TryParse(S(ws.Cell(r, dateCol)), out var d) && d.Year == date.Year && d.Month == date.Month) n++;
                return n;
            }
        }

        public EmployeeQuery Query(string code, DateTime date) {
            var e = GetEmployee(code);
            return new EmployeeQuery { 
                Employee = e, 
                RegularLeaveBalance = GetLeaveBalance(code, "إجازة اعتيادية"), 
                SickLeaveBalance = GetLeaveBalance(code, "إجازة مرضية"),
                ExitPermitsThisMonth = GetMonthlyExitPermitCount(code, date), 
                MissionsThisMonth = GetMonthlyMissionCount(code, date) 
            };
        }

        public List<LeaveRecordModel> GetEmployeeLeaves(string code) {
            var list = new List<LeaveRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.LeavesSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    if (S(ws.Cell(r, 2)) == code) {
                        list.Add(new LeaveRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)), نوع_الإجازة = S(ws.Cell(r, 5)),
                            تاريخ_البداية = DateTime.TryParse(S(ws.Cell(r, 6)), out var d1) ? d1.ToString("yyyy-MM-dd") : S(ws.Cell(r, 6)),
                            تاريخ_النهاية = DateTime.TryParse(S(ws.Cell(r, 7)), out var d2) ? d2.ToString("yyyy-MM-dd") : S(ws.Cell(r, 7)),
                            عدد_الأيام = int.TryParse(S(ws.Cell(r, 8)), out var days) ? days : 0,
                            تم_استلام_المرفق = S(ws.Cell(r, 14))
                        });
                    }
                }
            }
            return list;
        }

        public List<PermitRecordModel> GetEmployeeExitPermits(string code) {
            var list = new List<PermitRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.ExitPermitsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    if (S(ws.Cell(r, 2)) == code) {
                        list.Add(new PermitRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)),
                            التاريخ = DateTime.TryParse(S(ws.Cell(r, 5)), out var d) ? d.ToString("yyyy-MM-dd") : S(ws.Cell(r, 5)),
                            من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), السبب = S(ws.Cell(r, 8)),
                            تم_استلام_المرفق = S(ws.Cell(r, 12))
                        });
                    }
                }
            }
            return list;
        }

        public List<MissionRecordModel> GetEmployeeMissions(string code) {
            var list = new List<MissionRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.MissionsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    if (S(ws.Cell(r, 2)) == code) {
                        list.Add(new MissionRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)),
                            التاريخ = DateTime.TryParse(S(ws.Cell(r, 5)), out var d) ? d.ToString("yyyy-MM-dd") : S(ws.Cell(r, 5)),
                            من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), المكان = S(ws.Cell(r, 8)), السبب = S(ws.Cell(r, 9)),
                            تم_استلام_المرفق = S(ws.Cell(r, 13))
                        });
                    }
                }
            }
            return list;
        }

        // استعلام الإجازات بفترة (من - إلى) مع دعم إرخاء شرط الكود (فارغ أو * أو **** يعرض الكل)
        public List<LeaveRecordModel> GetEmployeeLeavesByDateRange(string code, DateTime from, DateTime to) {
            var list = new List<LeaveRecordModel>();
            bool allEmployees = string.IsNullOrWhiteSpace(code) || code == "*" || code == "****";
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.LeavesSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string empCode = S(ws.Cell(r, 2));
                    if ((allEmployees || empCode == code) && DateTime.TryParse(S(ws.Cell(r, 12)), out var d)) {
                        if (d.Date >= from.Date && d.Date <= to.Date) {
                            list.Add(new LeaveRecordModel {
                                RowIndex = r, كود_الموظف = empCode, اسم_الموظف = S(ws.Cell(r, 3)), نوع_الإجازة = S(ws.Cell(r, 5)),
                                تاريخ_البداية = DateTime.TryParse(S(ws.Cell(r, 6)), out var d1) ? d1.ToString("yyyy-MM-dd") : S(ws.Cell(r, 6)),
                                تاريخ_النهاية = DateTime.TryParse(S(ws.Cell(r, 7)), out var d2) ? d2.ToString("yyyy-MM-dd") : S(ws.Cell(r, 7)),
                                عدد_الأيام = int.TryParse(S(ws.Cell(r, 8)), out var days) ? days : 0,
                                تم_استلام_المرفق = S(ws.Cell(r, 14))
                            });
                        }
                    }
                }
            }
            return list;
        }

        // استعلام التصاريح بفترة (من - إلى) مع دعم عرض كل الموظفين
        public List<PermitRecordModel> GetEmployeeExitPermitsByDateRange(string code, DateTime from, DateTime to) {
            var list = new List<PermitRecordModel>();
            bool allEmployees = string.IsNullOrWhiteSpace(code) || code == "*" || code == "****";
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.ExitPermitsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string empCode = S(ws.Cell(r, 2));
                    if ((allEmployees || empCode == code) && DateTime.TryParse(S(ws.Cell(r, 10)), out var d)) {
                        if (d.Date >= from.Date && d.Date <= to.Date) {
                            list.Add(new PermitRecordModel {
                                RowIndex = r, كود_الموظف = empCode, اسم_الموظف = S(ws.Cell(r, 3)),
                                التاريخ = DateTime.TryParse(S(ws.Cell(r, 5)), out var d5) ? d5.ToString("yyyy-MM-dd") : S(ws.Cell(r, 5)),
                                من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), السبب = S(ws.Cell(r, 8)),
                                تم_استلام_المرفق = S(ws.Cell(r, 12))
                            });
                        }
                    }
                }
            }
            return list;
        }

        // استعلام المأموريات بفترة (من - إلى) مع دعم عرض كل الموظفين
        public List<MissionRecordModel> GetEmployeeMissionsByDateRange(string code, DateTime from, DateTime to) {
            var list = new List<MissionRecordModel>();
            bool allEmployees = string.IsNullOrWhiteSpace(code) || code == "*" || code == "****";
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.MissionsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string empCode = S(ws.Cell(r, 2));
                    if ((allEmployees || empCode == code) && DateTime.TryParse(S(ws.Cell(r, 11)), out var d)) {
                        if (d.Date >= from.Date && d.Date <= to.Date) {
                            list.Add(new MissionRecordModel {
                                RowIndex = r, كود_الموظف = empCode, اسم_الموظف = S(ws.Cell(r, 3)),
                                التاريخ = DateTime.TryParse(S(ws.Cell(r, 5)), out var d5) ? d5.ToString("yyyy-MM-dd") : S(ws.Cell(r, 5)),
                                من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), المكان = S(ws.Cell(r, 8)), السبب = S(ws.Cell(r, 9)),
                                تم_استلام_المرفق = S(ws.Cell(r, 13))
                            });
                        }
                    }
                }
            }
            return list;
        }

        public List<LeaveRecordModel> GetAllPendingLeaves() {
            var list = new List<LeaveRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.LeavesSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string status = S(ws.Cell(r, 14));
                    if (status != "تم الاستلام") {
                        list.Add(new LeaveRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)), نوع_الإجازة = S(ws.Cell(r, 5)),
                            تاريخ_البداية = S(ws.Cell(r, 6)), تاريخ_النهاية = S(ws.Cell(r, 7)),
                            عدد_الأيام = int.TryParse(S(ws.Cell(r, 8)), out var d) ? d : 0,
                            تم_استلام_المرفق = string.IsNullOrEmpty(status) ? "قيد الانتظار" : status
                        });
                    }
                }
            }
            return list;
        }

        public List<PermitRecordModel> GetAllPendingPermits() {
            var list = new List<PermitRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.ExitPermitsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string status = S(ws.Cell(r, 12));
                    if (status != "تم الاستلام") {
                        list.Add(new PermitRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)),
                            التاريخ = S(ws.Cell(r, 5)), من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), السبب = S(ws.Cell(r, 8)),
                            تم_استلام_المرفق = string.IsNullOrEmpty(status) ? "قيد الانتظار" : status
                        });
                    }
                }
            }
            return list;
        }

        public List<MissionRecordModel> GetAllPendingMissions() {
            var list = new List<MissionRecordModel>();
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.MissionsSheet); int last = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= last; r++) {
                    string status = S(ws.Cell(r, 13));
                    if (status != "تم الاستلام") {
                        list.Add(new MissionRecordModel {
                            RowIndex = r, كود_الموظف = S(ws.Cell(r, 2)), اسم_الموظف = S(ws.Cell(r, 3)),
                            التاريخ = S(ws.Cell(r, 5)), من_الساعة = S(ws.Cell(r, 6)), إلى_الساعة = S(ws.Cell(r, 7)), المكان = S(ws.Cell(r, 8)), السبب = S(ws.Cell(r, 9)),
                            تم_استلام_المرفق = string.IsNullOrEmpty(status) ? "قيد الانتظار" : status
                        });
                    }
                }
            }
            return list;
        }

        public void SaveLeaveRequest(string code, string name, string dept, string type, DateTime start, DateTime end, int days, decimal before, decimal after, string host, DateTime now, string? image) {
            int balanceCol = type == "إجازة مرضية" ? 4 : 3;
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.LeavesSheet); int row = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;
                ws.Cell(row, 1).Value = row - 1; ws.Cell(row, 2).Value = code; ws.Cell(row, 3).Value = name; ws.Cell(row, 4).Value = dept;
                ws.Cell(row, 5).Value = type; ws.Cell(row, 6).Value = start.Date; ws.Cell(row, 7).Value = end.Date; ws.Cell(row, 8).Value = days;
                ws.Cell(row, 9).Value = (double)before; ws.Cell(row, 10).Value = (double)after; ws.Cell(row, 11).Value = host; ws.Cell(row, 12).Value = now;
                AddImage(ws, row, 13, image); ws.Cell(row, 14).Value = "قيد الانتظار"; ws.Row(row).Height = 60;
                
                var bs = wb.Worksheet(AppConfig.BalancesSheet); int last = bs.LastRowUsed()?.RowNumber() ?? 1; bool found = false;
                for (int r = 2; r <= last; r++) if (S(bs.Cell(r, 1)) == code) { bs.Cell(r, balanceCol).Value = (double)after; found = true; break; }
                if (!found) { int r = last + 1; bs.Cell(r, 1).Value = code; bs.Cell(r, 2).Value = name; bs.Cell(r, balanceCol).Value = (double)after; }
                wb.Save();
            }
        }

        public void SaveExitPermit(string code, string name, string dept, DateTime date, TimeSpan from, TimeSpan to, string reason, string host, DateTime now, string? image) {
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.ExitPermitsSheet); int r = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;
                ws.Cell(r, 1).Value = r - 1; ws.Cell(r, 2).Value = code; ws.Cell(r, 3).Value = name; ws.Cell(r, 4).Value = dept; ws.Cell(r, 5).Value = date.Date;
                ws.Cell(r, 6).Value = from.ToString(@"hh\:mm"); ws.Cell(r, 7).Value = to.ToString(@"hh\:mm"); ws.Cell(r, 8).Value = reason;
                ws.Cell(r, 9).Value = host; ws.Cell(r, 10).Value = now; AddImage(ws, r, 11, image); ws.Cell(r, 12).Value = "قيد الانتظار"; wb.Save();
            }
        }

        public void SaveMission(string code, string name, string dept, DateTime date, TimeSpan from, TimeSpan to, string location, string reason, string host, DateTime now, string? image) {
            lock (Sync) using (var wb = Open()) {
                var ws = wb.Worksheet(AppConfig.MissionsSheet); int r = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;
                ws.Cell(r, 1).Value = r - 1; ws.Cell(r, 2).Value = code; ws.Cell(r, 3).Value = name; ws.Cell(r, 4).Value = dept; ws.Cell(r, 5).Value = date.Date;
                ws.Cell(r, 6).Value = from.ToString(@"hh\:mm"); ws.Cell(r, 7).Value = to.ToString(@"hh\:mm"); ws.Cell(r, 8).Value = location; ws.Cell(r, 9).Value = reason;
                ws.Cell(r, 10).Value = host; ws.Cell(r, 11).Value = now; AddImage(ws, r, 12, image); ws.Cell(r, 13).Value = "قيد الانتظار"; wb.Save();
            }
        }

        static void AddImage(IXLWorksheet ws, int row, int col, string? file) {
            if (string.IsNullOrWhiteSpace(file) || !File.Exists(file)) return;
            using var s = File.OpenRead(file); var p = ws.AddPicture(s).MoveTo(ws.Cell(row, col)); p.Width = 120; p.Height = 80;
        }
    }
}