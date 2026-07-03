using System.Windows.Forms;

namespace Unlocker_For_ExcelSheet
{
    /// <summary>
    /// 암호화된 엑셀 파일을 열기 위한 암호를 입력받는 모달 다이얼로그.
    /// resx/아이콘 리소스 없이 코드로만 구성한다.
    /// </summary>
    public partial class PasswordPromptForm : Form
    {
        public PasswordPromptForm(string fileName, string message)
        {
            InitializeComponent();
            lblInfo.Text = $"{fileName}{System.Environment.NewLine}{message}";
        }

        public string Password => txtPassword.Text;
    }
}
