# Employee Leave & Duty Management System

A Desktop Management System built with **C#**, **.NET 10 (Windows Forms)**, and **ClosedXML**. The application manages employee leave balances, exit permits, official missions, attachment approvals, and analytical reports using an Excel workbook backend.

---

## 🚀 Key Features

### 👤 Role-Based Access Control & Dual Workflows
- **Standard User (`user1`):**
  - **Leave Management:** Submit regular and sick leave requests with automated working day calculations (excluding Fridays and Saturdays) and instant balance validation.
  - **Exit Permits:** Issue exit permits with monthly limit enforcement (maximum 2 permits/month).
  - **Official Missions:** Register business missions detailing time, location, and reasoning.
  - **Attachment Upload:** Support for attaching image files (`.png`, `.jpg`, `.bmp`) to requests.
  - **Query System:** Comprehensive summary and historical tracking per employee.
  - **Security Lockout:** Prevents new submissions for an employee if previous request attachments remain pending review/approval.

- **Auditor / Approver (`user2`):**
  - **Attachment Approval Control:** Review all pending leave, permit, and mission attachments and mark them as received (`تم الاستلام`).
  - **Date-Range Reporting:** Generate and filter reports for leaves, exit permits, and missions by date range and specific or all employee codes.
  - **Printing Support:** Built-in hooks for printing summary reports.

### 📊 Excel-Driven Storage & Synchronization
- Embedded **ClosedXML** integration acts as an embedded database layer over `EmployeeServices2.xlsx`.
- Thread-safe workbook interactions utilizing locking mechanisms (`lock (Sync)`).
- Auto-initialization and auto-repair for required worksheets, tables, and headers.
- Embedded Excel dashboard worksheet populated with standard dynamic formulas (`COUNTA`).

### 🎨 Custom UI / UX
- Frameless Form design with custom-built drag-and-drop title bars, user switching capabilities, and quick window controls.
- RTL (Right-to-Left) UI support optimized for Arabic language interfaces.
- Responsive table layout grids (`TableLayoutPanel`) and responsive modals.

---

## 🛠️ Project Architecture

```
EmployeeLeaveApp/
├── AppConfig.cs             # Application configuration constants and credential setups
├── EmployeeLeaveApp.csproj  # Project file targeting .NET 10.0-windows with ClosedXML
├── ExcelService.cs          # Core data service interface and ClosedXML implementation
├── LoginForm.cs             # Authentication screen
├── MainForm.cs              # Main container, tabbed interfaces, query forms & reporting dialogs
├── Models.cs                # Data Transfer Objects (DTOs) and entity models
└── Program.cs               # Application entry point with dynamic session loops
```

---

## 📋 Prerequisites

- **Framework:** [.NET 10.0 SDK](https://dotnet.microsoft.com/) (Windows Workload)
- **OS:** Windows 10 / 11
- **Storage Path:** Access to file system path specified in `AppConfig.cs` (Default: `C:\Users\Omarpc\Desktop\EmployeeServices2.xlsx`).

---

## 📦 Dependencies

- **[ClosedXML](https://github.com/ClosedXML/ClosedXML) (v0.104.2):** High-performance library for reading, manipulating, and writing Excel files without requiring Microsoft Office installations.

---

## 🔧 Configuration

All application settings are managed in `AppConfig.cs`:

```csharp
public static class AppConfig {
    public const string ExcelFilePath = @"C:\Users\Omarpc\Desktop\EmployeeServices2.xlsx";
    public const string CompanyName = "شركة مصر للتعمير - إدارة نظم المعلومات";
    public const int MaxExitPermitsPerMonth = 2;
    public const int MaxMissionsPerMonth = 2;
    
    // Default Credentials
    public const string User1Name = "user1";
    public const string User1Pass = "123";
    public const string User2Name = "user2";
    public const string User2Pass = "456";
}
```

> **Note:** Ensure the configured directory in `ExcelFilePath` has write permissions. The system will automatically generate the required `.xlsx` structure if the file does not exist.

---

## 💻 Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/your-username/EmployeeLeaveApp.git
cd EmployeeLeaveApp
```

### 2. Restore Dependencies & Build
```bash
dotnet restore
dotnet build --configuration Release
```

### 3. Run Application
```bash
dotnet run
```

---

## 🔐 Default Credentials

| Role | Username | Password | Access Level |
| :--- | :--- | :--- | :--- |
| **Standard User** | `user1` | `123` | Entry forms, personal balance query, request submission |
| **Auditor / Admin** | `user2` | `456` | Pending attachment approvals, company-wide period reports |

---

## 📄 License

This project is proprietary software developed for **شركة مصر للتعمير - إدارة نظم المعلومات**.