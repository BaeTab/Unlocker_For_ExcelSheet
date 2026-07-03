using System;
using System.Text;
using System.Windows.Forms;

namespace Unlocker_For_ExcelSheet
{
    internal static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // NPOI(레거시 CFB/암호화 파일 파싱)가 CP949 등 코드페이지 인코딩을
            // 요구할 수 있으므로 등록해 둔다. (.NET Core+ 는 기본 미포함)
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            }
            catch
            {
                // 등록 실패해도 OOXML 처리에는 영향 없음
            }

            // High DPI + Visual Styles (csproj <UseWindowsForms> 로 소스 생성됨)
            ApplicationConfiguration.Initialize();

            // 전역 예외 안전망: 처리되지 않은 예외로 앱이 조용히 죽는 것을 막고 안내한다.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (s, e) =>
                MessageBox.Show("예기치 않은 오류가 발생했습니다:" + Environment.NewLine + Environment.NewLine + e.Exception.Message,
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                MessageBox.Show("치명적 오류: " + (ex?.Message ?? "알 수 없는 오류"),
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            Application.Run(new Form1());
        }
    }
}
