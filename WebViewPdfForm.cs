using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace PrintLogPdf
{
    public class WebViewPdfForm : Form
    {
        private WebView2 webView = new WebView2();

        public WebViewPdfForm(string pdfPath)
        {
            Text = "Airex PDF Viewer";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 900;
            Height = 800;

            webView.Dock = DockStyle.Fill;
            Controls.Add(webView);

            Load += async (_, __) =>
            {
                // ===== Fixed Version Runtime 경로 =====
                string exeDir = AppContext.BaseDirectory;

                string runtimePath = Path.Combine(
                    exeDir,
                    "WebView2Runtime"
                );

                string userDataPath = Path.Combine(
                    exeDir,
                    "WebView2UserData"
                );

                if (!Directory.Exists(runtimePath))
                {
                    MessageBox.Show(
                        "WebView2Runtime 폴더를 찾을 수 없습니다.",
                        "Runtime Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    Close();
                    return;
                }

                var env = await CoreWebView2Environment.CreateAsync(
                    runtimePath,
                    userDataPath
                );

                await webView.EnsureCoreWebView2Async(env);

                // ===== 새 창 / 외부 링크 차단 =====
                webView.CoreWebView2.NewWindowRequested += (s, e) =>
                {
                    e.Handled = true;
                };

                // ===== 로컬 PDF 접근 허용 =====
                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "local",
                    Path.GetDirectoryName(pdfPath)!,
                    CoreWebView2HostResourceAccessKind.Allow
                );

                var fileName = Path.GetFileName(pdfPath);
                webView.Source = new Uri($"https://local/{fileName}");
            };
        }
    }
}
