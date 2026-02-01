using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;



namespace PrintLogPdf
{
    
    enum LogCategory
    {
        Login,
        Alarm,
        PlcReason,
        Scada,
        Other
    }
    

    class LogRow
    {
        public string D  { get; set; } = "";
        public string T  { get; set; } = "";
        public string U  { get; set; } = "";
        public string Ty { get; set; } = "";
        public string M  { get; set; } = "";

        public string Recovery { get; set; } = "";
        public LogCategory Category { get; set; }
    }


    public partial class Form1 : Form
    {
        Dictionary<string, string> AllowedUsers = new()
        {
            { "lee", "6666" },
            { "kim", "1234" }
        };

        DateTimePicker dtFrom = new();
        DateTimePicker dtTo = new();
        TextBox txtUser = new();
        TextBox txtPw = new();
        Button btnExport = new();
        Label lblFrom = new();
        Label lblTo   = new();
        Button btnExportAndView = new();

        //터치키보드 실행함수
        void ShowTouchKeyboard()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = @"C:\Program Files\Common Files\Microsoft Shared\ink\TabTip.exe",
                    UseShellExecute = true
                });
            }
            catch
            {
                // 터치PC가 아니거나 TabTip 없는 경우 무시
            }
        }



        string SectionTitle(LogCategory c)
        {
            return c switch
            {
                LogCategory.Login     => "1. Login Logs",
                LogCategory.Alarm     => "2. Alarm Logs",
                LogCategory.PlcReason => "3. Mauual Operation Logs",
                LogCategory.Scada     => "4. HMI Program Open/Close Logs",
                LogCategory.Other     => "5. Other Logs",
                _                     => ""
            };
        }


        public Form1()
        {
            InitializeComponent();

            Text = "Airex Log PDF Export";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(792, 600);
            MinimumSize = new System.Drawing.Size(440, 520);


            var layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(20);
            layout.ColumnCount = 1;
            layout.RowCount = 12;

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 0 From label
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); // 1 From picker
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 14)); // 2 gap

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // 3 To label
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); // 4 To picker
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20)); // 5 gap

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); // 6 User
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 14)); // 7 gap

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); // 8 PW

            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // 9 🔑 남는 공간
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52)); // 10 Button
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10)); // 11 bottom gap


            // ===== From =====
            lblFrom.Text = "From";
            lblFrom.AutoSize = true;
            lblFrom.Dock = DockStyle.Top;
            lblFrom.Padding = new Padding(0, 2, 0, 4);
            lblFrom.TextAlign = ContentAlignment.BottomLeft;


            dtFrom.Dock = DockStyle.Fill;

            // ===== To =====
            lblTo.Text = "To";
            lblTo.AutoSize = true;
            lblTo.Dock = DockStyle.Top;
            lblTo.Padding = new Padding(0, 2, 0, 4);
            lblTo.TextAlign = ContentAlignment.BottomLeft;

            dtTo.Dock = DockStyle.Fill;

            // ===== User ID =====
            txtUser.Dock = DockStyle.Fill;
            txtUser.PlaceholderText = "User ID";

            // ===== Password =====
            txtPw.Dock = DockStyle.Fill;
            txtPw.PasswordChar = '*';
            txtPw.PlaceholderText = "Password";

            // ===== Export Button =====
            btnExport.Dock = DockStyle.Fill;
            btnExport.Text = "출력 (PDF)";
            btnExport.Click += ExportPdf;

            // Export and view Button
            btnExportAndView.Dock = DockStyle.Fill;
            btnExportAndView.Text = "출력 + 보기";
            btnExportAndView.Click += ExportPdfAndView;

            var btnRow = new TableLayoutPanel();
            btnRow.Dock = DockStyle.Fill;
            btnRow.ColumnCount = 2;
            btnRow.RowCount = 1;
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            btnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            btnExport.Dock = DockStyle.Fill;
            btnExportAndView.Dock = DockStyle.Fill;

            btnRow.Controls.Add(btnExport, 0, 0);
            btnRow.Controls.Add(btnExportAndView, 1, 0);

            layout.Controls.Add(lblFrom, 0, 0);
            layout.Controls.Add(dtFrom, 0, 1);
            layout.Controls.Add(lblTo, 0, 3);
            layout.Controls.Add(dtTo, 0, 4);
            layout.Controls.Add(txtUser, 0, 6);
            layout.Controls.Add(txtPw, 0, 8);
            layout.Controls.Add(btnRow, 0, 10);

            Controls.Add(layout);
            //키보드창 팝업
            txtUser.Enter += (s, e) => ShowTouchKeyboard();
            txtPw.Enter   += (s, e) => ShowTouchKeyboard();
        }

        LogCategory Classify(string msg)
        {
            if (msg.Contains("Login", StringComparison.OrdinalIgnoreCase) ||
                msg.Contains("Logout", StringComparison.OrdinalIgnoreCase))
                return LogCategory.Login;

            if (msg.Contains("PLC", StringComparison.OrdinalIgnoreCase) ||
                msg.Contains("Reason", StringComparison.OrdinalIgnoreCase))
                return LogCategory.PlcReason;

            if (msg.Contains("SCADA", StringComparison.OrdinalIgnoreCase))
                return LogCategory.Scada;

            return LogCategory.Other;
        }



        private void ExportPdfAndView(object? sender, EventArgs e)
        {
            var pdfPath = GeneratePdf();
            if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
                return;

            new WebViewPdfForm(pdfPath).Show();

        }



        private string? GeneratePdf()
        {
            try
            {
                string from = dtFrom.Value.ToString("yyyyMMdd");
                string to   = dtTo.Value.ToString("yyyyMMdd");

                string userId = txtUser.Text.Trim();
                if (string.IsNullOrWhiteSpace(userId))
                    userId = "UNKNOWN";

                var rows = new List<LogRow>();

                string SystemDbPath = @"C:\SystemLog\SystemLog.db";
                string AlarmDbPath = @"C:\Alarm\GlobalAlarm.db";

                string SystemconnStr = $"Data Source={SystemDbPath};";
                string AlarmconnStr = $"Data Source={AlarmDbPath}";

                string lastLoginUserId = "UNKNOWN";
                string lastLoginDate   = "-";
                string lastLoginTime   = "-";

                using (var conn = new SQLiteConnection(AlarmconnStr))
                {
                    conn.Open();

                    string sqlAlarm = @"
                        SELECT
                        OCCURE_DATE,
                        OCCURE_TIME,
                        RECOVERY_TIME,
                        MSG
                    FROM TB_ALARM1
                    WHERE OCCURE_DATE BETWEEN @from AND @to
                    ORDER BY OCCURE_DATE DESC, OCCURE_TIME DESC;

                    ";

                    using (var cmd = new SQLiteCommand(sqlAlarm, conn))
                    {
                        cmd.Parameters.AddWithValue("@from", from);
                        cmd.Parameters.AddWithValue("@to", to);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string occurDate    = reader["OCCURE_DATE"]?.ToString() ?? "";
                                string occurTime    = reader["OCCURE_TIME"]?.ToString() ?? "";
                                string recoveryTime = reader["RECOVERY_TIME"]?.ToString() ?? "-";
                                string msg          = reader["MSG"]?.ToString() ?? "";

                                rows.Add(new LogRow
                                {
                                    D = occurDate,
                                    T = occurTime,
                                    U = "-",
                                    Ty = "ALARM",
                                    M = msg,
                                    Recovery = recoveryTime,
                                    Category = LogCategory.Alarm
                                });
                            }
                        }
                    }

                }


                using (var conn = new SQLiteConnection(SystemconnStr))
                {
                    conn.Open();
                    string sqlLogList = @"
                    SELECT
                    LOG_DATE,
                    LOG_TIME,
                    USER_ID,
                    LOG_TYPE,
                    LOG_MSG
                    FROM TB_SECULOG
                    WHERE LOG_DATE BETWEEN @from AND @to
                    ORDER BY LOG_DATE DESC, LOG_TIME DESC;
                    ";
                    
                    string sqlLastLogin = @"
                    SELECT USER_ID, USER_NM, LOG_DATE, LOG_TIME
                    FROM TB_SECULOG
                    WHERE LOG_MSG LIKE 'Login - ID:%'
                    AND USER_ID IS NOT NULL
                    AND LOG_DATE BETWEEN @from AND @to
                    ORDER BY LOG_DATE DESC, LOG_TIME DESC
                    LIMIT 1;
                    ";

                    using (var cmdLast = new SQLiteCommand(sqlLastLogin, conn))
                    {
                        
                        cmdLast.Parameters.AddWithValue("@from", from);
                        cmdLast.Parameters.AddWithValue("@to", to);

                        using (var rLast = cmdLast.ExecuteReader())
                        {
                            if (rLast.Read())
                            {
                                string uid  = rLast["USER_ID"].ToString()!;
                                string role = rLast["USER_NM"].ToString()!;

                                lastLoginUserId = $"{uid}({role.ToLower()})";
                                lastLoginDate   = rLast["LOG_DATE"].ToString()!;
                                lastLoginTime   = rLast["LOG_TIME"].ToString()!;
                            }
                            else
                            {
                                // 기간 내 로그인 없을 때
                                lastLoginUserId = "NONE";
                                lastLoginDate = "-";
                                lastLoginTime = "-";
                            }
                        }
                    }
                    using var cmd = new SQLiteCommand(sqlLogList, conn);

                    cmd.Parameters.AddWithValue("@from", from);
                    cmd.Parameters.AddWithValue("@to", to);

                    using var r = cmd.ExecuteReader();
                    while (r.Read())
                    {
                        var log = new LogRow
                        {
                            D  = r["LOG_DATE"].ToString()!,
                            T  = r["LOG_TIME"].ToString()!,
                            U  = r["USER_ID"].ToString()!,
                            Ty = r["LOG_TYPE"].ToString()!,
                            M  = r["LOG_MSG"].ToString()!
                        };

                        log.Category = Classify(log.M); 
                        rows.Add(log);
                    }
                }

                QuestPDF.Settings.License = LicenseType.Community;

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"Airex_{timestamp}.pdf";
                string titleText = "Isolator Batch Process Record";
                

                string pdfPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    fileName
                );

                Document.Create(doc =>
                {
                    doc.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(30);

                        page.Content().Column(col =>
                        {
                            // ===== 헤더 =====
                            col.Item().PaddingBottom(6)
                                .Text(titleText)
                                .FontSize(16)
                                .Bold();

                            col.Item().PaddingBottom(6).LineHorizontal(3).LineColor(Colors.Green.Darken2);
                            col.Item().Text("");
                            col.Item().Text("");

                            col.Item().PaddingTop(15).LineHorizontal(2).LineColor(Colors.LightBlue.Medium);
                        
                            col.Item().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(130);   // 항목명
                                        columns.RelativeColumn();     // 값
                                    });

                                    void Row(string label, string value)
                                    {
                                        table.Cell().PaddingVertical(8)
                                            .Text(label)
                                            .SemiBold();

                                        table.Cell().PaddingVertical(8)
                                            .Text(value);
                                    }

                                    Row("관리자(문서검토자)", userId);
                                    Row("문서출력시간", timestamp);
                                    Row("조회기간", $"{from} ~ {to}");
                                });

                            
                            col.Item().PaddingTop(15).LineHorizontal(2).LineColor(Colors.LightBlue.Medium);

                            col.Item().Text("");
                            col.Item().Text("");

                            var grouped = rows
                                .GroupBy(r => r.Category)
                                .ToDictionary(g => g.Key, g => g.ToList());

                            // ===== Section 1 : Login Info =====
                            col.Item().PaddingTop(30)
                                .Text("1. Login Info")
                                .FontSize(14)
                                .Bold();

                            col.Item().LineHorizontal(2);
                            col.Item().PaddingBottom(10);

                            if (!grouped.TryGetValue(LogCategory.Login, out var loginItems)
                                || loginItems.Count == 0)
                            {
                                // Login 데이터 없음
                                col.Item()
                                    .PaddingTop(12)
                                    .Text("내용 없음")
                                    .Italic()
                                    .FontColor(Colors.Grey.Medium);
                            }
                            else
                            {
                                // Login 데이터 있음
                                col.Item().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(80);   // 항목명
                                        columns.RelativeColumn();     // 값
                                    });

                                    void Row(string label, string value)
                                    {
                                        table.Cell().PaddingVertical(4)
                                            .Text(label)
                                            .SemiBold();

                                        table.Cell().PaddingVertical(4)
                                            .Text(value);
                                    }

                                    Row("작업자", lastLoginUserId);
                                    Row("작업일", lastLoginDate);
                                    Row("작업시간", lastLoginTime);
                                });

                                col.Item().PaddingBottom(20);
                            }


                            // Section 2 : Alarm
                            col.Item().PaddingTop(30)
                                .Text("2. Alarm Logs")
                                .FontSize(14)
                                .Bold();

                            col.Item().LineHorizontal(2);
                            col.Item().PaddingBottom(10);

                            if (!grouped.TryGetValue(LogCategory.Alarm, out var alarmItems)
                                || alarmItems.Count == 0)
                            {
                                col.Item()
                                    .PaddingTop(12)
                                    .Text("내용 없음")
                                    .Italic()
                                    .FontColor(Colors.Grey.Medium);
                            }
                            else
                            {
                                col.Item().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.RelativeColumn(2);
                                        columns.RelativeColumn(2);
                                        columns.RelativeColumn(3);
                                        columns.RelativeColumn(2);
                                    });

                                    // Header
                                    table.Header(header =>
                                    {
                                        header.Cell().Background(Colors.Grey.Darken3)
                                            .Padding(5).Text("Date").FontColor(Colors.White).Bold();

                                        header.Cell().Background(Colors.Grey.Darken3)
                                            .Padding(5).Text("Occur Time").FontColor(Colors.White).Bold();

                                        header.Cell().Background(Colors.Grey.Darken3)
                                            .Padding(5).Text("Alarm Message").FontColor(Colors.White).Bold();

                                        header.Cell().Background(Colors.Grey.Darken3)
                                            .Padding(5).Text("Recovery Time").FontColor(Colors.White).Bold();
                                    });

                                    for (int i = 0; i < alarmItems.Count; i++)
                                    {
                                        var r = alarmItems[i];
                                        var bg = (i % 2 == 0)
                                            ? Colors.Grey.Lighten5
                                            : Colors.Grey.Lighten2;

                                        table.Cell().Background(bg).Padding(6).Text(r.D).FontSize(9);
                                        table.Cell().Background(bg).Padding(6).Text(r.T).FontSize(9);
                                        table.Cell().Background(bg).Padding(6)
                                            .Text(r.M)
                                            .FontSize(9)
                                            .FontColor(string.IsNullOrEmpty(r.Recovery)
                                                ? Colors.Red.Darken2
                                                : Colors.Black);

                                        table.Cell().Background(bg).Padding(6)
                                            .Text(string.IsNullOrEmpty(r.Recovery) ? "-" : r.Recovery)
                                            .FontSize(9);
                                    }
                                });
                            }

                        });
                    });
                    // 📄 PAGE 2~ : 섹션 하나당 한 페이지
                    LogCategory[] rest =
                    {
                        LogCategory.PlcReason,
                        LogCategory.Scada,
                        LogCategory.Other
                    };
                    
                    foreach (var cat in rest)
                    {
                        var catRows = rows.Where(r => r.Category == cat).ToList();

                        doc.Page(page =>
                        {
                            page.Size(PageSizes.A4);
                            page.Margin(30);

                            page.Content().Column(col =>
                            {
                                // 제목은 항상 출력
                                col.Item()
                                    .Text(SectionTitle(cat))
                                    .FontSize(14)
                                    .Bold();

                                col.Item().LineHorizontal(2);
                                col.Item().PaddingBottom(10);

                                if (catRows.Count == 0)
                                {
                                    col.Item().PaddingTop(12)
                                        .Text("내용 없음")
                                        .Italic()
                                        .FontColor(Colors.Grey.Medium);
                                }
                                else
                                {
                                    for (int i = 0; i < catRows.Count; i++)
                                    {
                                        var r = catRows[i];

                                        //zebra pattern start
                                        var bg = (i % 2 == 0)
                                            ? Colors.Grey.Lighten5
                                            : Colors.Grey.Lighten2;

                                        col.Item()
                                            .Background(bg)          
                                            .PaddingVertical(6)
                                            .PaddingHorizontal(8)
                                            .Text($"{r.D} {r.T} | {r.U} | {r.Ty} | {r.M}")
                                            .FontSize(9)
                                            .LineHeight(1.4f);
                                    }
                                }
                            });
                        });
                    }
                }).GeneratePdf(pdfPath);

                return pdfPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR");
                return null;
            }
        }

        private void ExportPdf(object? sender, EventArgs e)
        {
            var pdfPath = GeneratePdf();
            if (pdfPath != null)
                MessageBox.Show($"PDF 생성 완료\n{pdfPath}");
        }

    }
}
