using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Ionic.Zip;

namespace ExcelPasswordRemover
{
    public partial class Form1 : Form
    {
        private string excelFilePath = string.Empty;
        private string ziplFilePath = string.Empty;
        private string extractPath = string.Empty;

        public Form1()
        {
            InitializeComponent();
            regEvent();
        }

        private void regEvent()
        {
            button1.Click += Button1_Click;
            button2.Click += Button2_Click;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel 파일 (*.xlsx)|*.xlsx";
            openFileDialog1.Title = "Excel 파일 선택";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = openFileDialog1.FileName;
                ziplFilePath = Path.ChangeExtension(excelFilePath, ".zip");
                extractPath = Path.Combine(Path.GetDirectoryName(excelFilePath), "임시폴더");

                textBox1.Text = excelFilePath;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (excelFilePath != string.Empty)
            {
                RemovePasswordFromExcel(excelFilePath, ziplFilePath, extractPath);
                
            }
            else
            {
                MessageBox.Show("먼저 파일을 선택 해 주세요", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RemovePasswordFromExcel(string excelFilePath, string zipFilePath, string extractPath)
        {
            try
            {
                //  엑셀파일 확장자 변경
                File.Move(excelFilePath, zipFilePath);

                // zip 파일 압축을 푼다
                using (ZipFile zip = ZipFile.Read(zipFilePath))
                {
                    zip.ExtractAll(extractPath);
                }
                // xml 파일에서 보호시트 관련 태그 삭제하는 프로세스
                ProcessXmlFiles(extractPath);
                // 재압축
                using (ZipFile zip = new ZipFile())
                {
                    zip.AddDirectory(extractPath);
                    zip.Save(zipFilePath);
                }
                // 압축된 파일을 다시 xlsx 확장자로 바꿔준다
                string newFilePath = Path.ChangeExtension(zipFilePath, ".xlsx");
                File.Move(zipFilePath, newFilePath);
                label3.Text = "보호시트 비밀번호 해제 완료";

                if (DialogResult.Yes == MessageBox.Show("작업이 완료되었습니다 해당파일을 실행 하시겠습니까?", "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
                {
                    Process.Start(newFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류: {ex.Message}");
            }
            finally
            {
                // 임시파일 제거
                Directory.Delete(extractPath, true);
            }
        }

        // xml 파일 프로세스
        private void ProcessXmlFiles(string directoryPath)
        {
            string[] xmlFiles = Directory.GetFiles(directoryPath, "*.xml", SearchOption.AllDirectories);

            try
            {
                if(null != xmlFiles)
                {
                    foreach (string xmlFile in xmlFiles)
                    {
                        string fileContent = File.ReadAllText(xmlFile);
                        // 정규식을 사용해 sheetProtection 으로 시작하고 그 뒤에 어떤 속성이든 있을 수 있는 태그를 찾는다
                        fileContent = Regex.Replace(fileContent, "<sheetProtection[^>]*>", "", RegexOptions.Singleline);
                        File.WriteAllText(xmlFile, fileContent);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류: {ex.Message}");
            }
        }
    }
}
