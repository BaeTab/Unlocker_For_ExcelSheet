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
            Application.Run(new Form1());
        }
    }
}
