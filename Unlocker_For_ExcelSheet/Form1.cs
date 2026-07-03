using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Unlocker_For_ExcelSheet.Services;

namespace Unlocker_For_ExcelSheet
{
    public partial class Form1 : Form
    {
        // 큐에 담을 수 있는 확장자 (대소문자 무시)
        private static readonly string[] SupportedExtensions =
        {
            ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"
        };

        private readonly ExcelUnlocker unlocker = new ExcelUnlocker();
        private string lastOutputFolder;

        public Form1()
        {
            InitializeComponent();
        }

        #region 큐 관리 (버튼 / 드래그앤드롭 공용)

        private void BtnAddFiles_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.Title = "엑셀 파일 선택";
                dialog.Filter = "엑셀 파일 (*.xlsx;*.xlsm;*.xltx;*.xltm;*.xls)|*.xlsx;*.xlsm;*.xltx;*.xltm;*.xls|모든 파일 (*.*)|*.*";

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    foreach (string file in dialog.FileNames)
                    {
                        AddFile(file);
                    }
                }
            }
        }

        private void BtnAddFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "엑셀 파일을 검색할 폴더를 선택하세요";

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    AddPath(dialog.SelectedPath);
                }
            }
        }

        private void BtnRemoveSelected_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listFiles.SelectedItems)
            {
                listFiles.Items.Remove(item);
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            listFiles.Items.Clear();
        }

        private async void BtnCheckUpdate_Click(object sender, EventArgs e)
        {
            btnCheckUpdate.Enabled = false;
            try
            {
                var info = await UpdateChecker.CheckAsync();
                if (info == null)
                {
                    MessageBox.Show(this, "업데이트 정보를 가져오지 못했습니다.", "업데이트",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    var r = MessageBox.Show(this, info.Message + Environment.NewLine + Environment.NewLine +
                        "다운로드 페이지를 여시겠습니까?", "업데이트", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (r == DialogResult.Yes)
                    {
                        OpenWithShell(string.IsNullOrEmpty(info.DownloadUrl) ? info.ReleaseUrl : info.DownloadUrl);
                    }
                }
                else
                {
                    MessageBox.Show(this, info.Message, "업데이트",
                        MessageBoxButtons.OK, info.CheckFailed ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // async void 이므로 예외가 새어나가면 앱이 죽는다. 여기서 흡수.
                MessageBox.Show(this, "업데이트 확인 중 오류: " + ex.Message, "업데이트",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                btnCheckUpdate.Enabled = true;
            }
        }

        /// <summary>URL/폴더 경로를 셸로 연다. 실패(핸들러 없음/경로 소멸 등)해도 앱이 죽지 않도록 흡수.</summary>
        private void OpenWithShell(string target)
        {
            if (string.IsNullOrEmpty(target))
            {
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(target) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "열 수 없습니다: " + ex.Message, "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 창이 뜨자마자 조용히 업데이트를 확인한다. 실패/최신 버전이면 아무 것도 하지 않고,
        /// 새 버전이 있을 때만 하단 상태줄에 안내를 띄운다(다이얼로그로 막지 않음).
        /// </summary>
        private async void Form1_Shown(object sender, EventArgs e)
        {
            try
            {
                var info = await UpdateChecker.CheckAsync();
                if (info != null && info.UpdateAvailable)
                {
                    lblStatus.Text = $"● 새 버전 v{info.LatestVersion} 사용 가능 — [업데이트 확인] 클릭";
                }
            }
            catch
            {
                // 조용히 무시 — 시작 시 확인은 실패해도 앱 동작에 영향 없어야 한다.
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string path in paths)
            {
                AddPath(path);
            }
        }

        /// <summary>
        /// 파일이든 폴더든 받아서 큐에 반영한다. 폴더면 하위 폴더까지 재귀 검색한다.
        /// </summary>
        private void AddPath(string path)
        {
            if (Directory.Exists(path))
            {
                IEnumerable<string> files = SupportedExtensions
                    .SelectMany(ext => Directory.GetFiles(path, "*" + ext, SearchOption.AllDirectories));

                foreach (string file in files)
                {
                    AddFile(file);
                }
            }
            else if (File.Exists(path))
            {
                AddFile(path);
            }
        }

        /// <summary>
        /// 지원 확장자인지, 이미 큐에 있는지 확인한 뒤 listFiles 에 한 줄 추가한다.
        /// </summary>
        private void AddFile(string file)
        {
            if (!IsSupported(file))
            {
                return;
            }

            string fullPath = Path.GetFullPath(file);

            bool alreadyQueued = listFiles.Items.Cast<ListViewItem>()
                .Any(item => string.Equals((string)item.Tag, fullPath, StringComparison.OrdinalIgnoreCase));
            if (alreadyQueued)
            {
                return;
            }

            var item = new ListViewItem(Path.GetFileName(fullPath));
            item.SubItems.Add("대기");
            item.SubItems.Add(fullPath);
            item.Tag = fullPath;

            listFiles.Items.Add(item);
        }

        private static bool IsSupported(string file)
        {
            string ext = Path.GetExtension(file);
            return SupportedExtensions.Any(supported => string.Equals(supported, ext, StringComparison.OrdinalIgnoreCase));
        }

        #endregion

        #region 해제 처리

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (listFiles.Items.Count == 0)
            {
                MessageBox.Show(this, "먼저 처리할 파일을 추가해 주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetBusy(true);

            // 새 배치 시작 시 이전 배치의 결과 폴더 상태를 리셋한다.
            btnOpenFolder.Enabled = false;
            lastOutputFolder = null;

            int successCount = 0;
            int noProtectionCount = 0;
            int failCount = 0;

            progressBar.Minimum = 0;
            progressBar.Maximum = listFiles.Items.Count;
            progressBar.Value = 0;

            try
            {
                foreach (ListViewItem item in listFiles.Items)
                {
                    string path = (string)item.Tag;
                    SetItemStatus(item, "처리 중...", SystemColors.WindowText);

                    try
                    {
                        var opts = new UnlockOptions { OverwriteOriginal = chkOverwrite.Checked };
                        UnlockResult result = await unlocker.UnlockAsync(path, opts);

                        if (result.Status == UnlockStatus.PasswordRequired)
                        {
                            result = await PromptForPasswordAndRetryAsync(path, opts, result);
                        }

                        ApplyResult(item, result);

                        switch (result.Status)
                        {
                            case UnlockStatus.Success:
                                successCount++;
                                break;
                            case UnlockStatus.NoProtectionFound:
                                noProtectionCount++;
                                break;
                            default:
                                failCount++;
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        SetItemStatus(item, "실패", Color.Firebrick);
                        AppendLog($"[{Path.GetFileName(path)}] 오류: {ex.Message}");
                        failCount++;
                    }

                    progressBar.Value++;
                    await Task.Yield();
                }
            }
            finally
            {
                SetBusy(false);
            }

            lblStatus.Text = $"완료 — 성공 {successCount}, 보호 없음 {noProtectionCount}, 실패 {failCount}";

            if (!string.IsNullOrEmpty(lastOutputFolder))
            {
                btnOpenFolder.Enabled = true;
            }
        }

        /// <summary>
        /// 열기 암호가 필요한 파일에 대해 최대 3회까지 암호 입력을 받아 재시도한다.
        /// </summary>
        private async Task<UnlockResult> PromptForPasswordAndRetryAsync(string path, UnlockOptions opts, UnlockResult result)
        {
            const int maxAttempts = 3;
            string fileName = Path.GetFileName(path);
            string message = result.Message;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                using (var promptForm = new PasswordPromptForm(fileName, message))
                {
                    if (promptForm.ShowDialog(this) != DialogResult.OK)
                    {
                        return new UnlockResult
                        {
                            SourcePath = path,
                            OutputPath = null,
                            Status = UnlockStatus.Failed,
                            Message = "암호 미입력으로 건너뜀"
                        };
                    }

                    opts.OpenPassword = promptForm.Password;
                    result = await unlocker.UnlockAsync(path, opts);

                    if (result.Status != UnlockStatus.WrongPassword)
                    {
                        return result;
                    }

                    message = "암호가 올바르지 않습니다. 다시 입력하세요.";
                }
            }

            return result;
        }

        private void ApplyResult(ListViewItem item, UnlockResult result)
        {
            switch (result.Status)
            {
                case UnlockStatus.Success:
                    SetItemStatus(item, "완료", Color.Green);
                    if (!string.IsNullOrEmpty(result.OutputPath))
                    {
                        lastOutputFolder = Path.GetDirectoryName(result.OutputPath);
                        AppendLog($"[{result.FileName}] {result.Message} -> {result.OutputPath}");
                    }
                    else
                    {
                        AppendLog($"[{result.FileName}] {result.Message}");
                    }
                    break;

                case UnlockStatus.NoProtectionFound:
                    SetItemStatus(item, "보호 없음", Color.Gray);
                    AppendLog($"[{result.FileName}] {result.Message}");
                    break;

                default:
                    SetItemStatus(item, "실패", Color.Firebrick);
                    AppendLog($"[{result.FileName}] {result.Message}");
                    break;
            }
        }

        private void SetItemStatus(ListViewItem item, string status, Color color)
        {
            item.SubItems[1].Text = status;
            item.ForeColor = color;
        }

        private void AppendLog(string line)
        {
            txtLog.AppendText(line + Environment.NewLine);
        }

        private void SetBusy(bool busy)
        {
            btnAddFiles.Enabled = !busy;
            btnAddFolder.Enabled = !busy;
            btnRemoveSelected.Enabled = !busy;
            btnClear.Enabled = !busy;
            btnStart.Enabled = !busy;
            btnCheckUpdate.Enabled = !busy;
            chkOverwrite.Enabled = !busy;
        }

        #endregion

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(lastOutputFolder) || !Directory.Exists(lastOutputFolder))
            {
                MessageBox.Show(this, "열 수 있는 결과 폴더가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            OpenWithShell(lastOutputFolder);
        }
    }
}
