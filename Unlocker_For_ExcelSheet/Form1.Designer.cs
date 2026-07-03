namespace Unlocker_For_ExcelSheet
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tablePanelMain = new System.Windows.Forms.TableLayoutPanel();
            this.panelTop = new System.Windows.Forms.Panel();
            this.btnAddFiles = new System.Windows.Forms.Button();
            this.btnAddFolder = new System.Windows.Forms.Button();
            this.btnRemoveSelected = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnCheckUpdate = new System.Windows.Forms.Button();
            this.listFiles = new System.Windows.Forms.ListView();
            this.columnFileName = new System.Windows.Forms.ColumnHeader();
            this.columnStatus = new System.Windows.Forms.ColumnHeader();
            this.columnPath = new System.Windows.Forms.ColumnHeader();
            this.chkOverwrite = new System.Windows.Forms.CheckBox();
            this.panelStart = new System.Windows.Forms.Panel();
            this.btnStart = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.tablePanelMain.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.panelStart.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            //
            // tablePanelMain
            //
            // 상단 버튼바 → listFiles(가변) → 옵션 → 시작줄 → txtLog(가변) → 하단 상태줄
            // 순서로 쌓는 6행 레이아웃. listFiles/txtLog 행만 Percent 로 잡아
            // 창을 늘리면 두 영역이 겹치지 않고 함께 늘어난다.
            this.tablePanelMain.ColumnCount = 1;
            this.tablePanelMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tablePanelMain.Controls.Add(this.panelTop, 0, 0);
            this.tablePanelMain.Controls.Add(this.listFiles, 0, 1);
            this.tablePanelMain.Controls.Add(this.chkOverwrite, 0, 2);
            this.tablePanelMain.Controls.Add(this.panelStart, 0, 3);
            this.tablePanelMain.Controls.Add(this.txtLog, 0, 4);
            this.tablePanelMain.Controls.Add(this.panelBottom, 0, 5);
            this.tablePanelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tablePanelMain.Location = new System.Drawing.Point(0, 0);
            this.tablePanelMain.Name = "tablePanelMain";
            this.tablePanelMain.Padding = new System.Windows.Forms.Padding(12);
            this.tablePanelMain.RowCount = 6;
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 60F));
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.tablePanelMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tablePanelMain.Size = new System.Drawing.Size(720, 520);
            this.tablePanelMain.TabIndex = 0;
            //
            // panelTop
            //
            this.panelTop.Controls.Add(this.btnAddFiles);
            this.panelTop.Controls.Add(this.btnAddFolder);
            this.panelTop.Controls.Add(this.btnRemoveSelected);
            this.panelTop.Controls.Add(this.btnClear);
            this.panelTop.Controls.Add(this.btnCheckUpdate);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTop.Location = new System.Drawing.Point(15, 15);
            this.panelTop.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(690, 38);
            this.panelTop.TabIndex = 0;
            //
            // btnAddFiles
            //
            this.btnAddFiles.Location = new System.Drawing.Point(0, 4);
            this.btnAddFiles.Name = "btnAddFiles";
            this.btnAddFiles.Size = new System.Drawing.Size(100, 30);
            this.btnAddFiles.TabIndex = 0;
            this.btnAddFiles.Text = "파일 추가";
            this.btnAddFiles.UseVisualStyleBackColor = true;
            this.btnAddFiles.Click += new System.EventHandler(this.BtnAddFiles_Click);
            //
            // btnAddFolder
            //
            this.btnAddFolder.Location = new System.Drawing.Point(106, 4);
            this.btnAddFolder.Name = "btnAddFolder";
            this.btnAddFolder.Size = new System.Drawing.Size(100, 30);
            this.btnAddFolder.TabIndex = 1;
            this.btnAddFolder.Text = "폴더 추가";
            this.btnAddFolder.UseVisualStyleBackColor = true;
            this.btnAddFolder.Click += new System.EventHandler(this.BtnAddFolder_Click);
            //
            // btnRemoveSelected
            //
            this.btnRemoveSelected.Location = new System.Drawing.Point(212, 4);
            this.btnRemoveSelected.Name = "btnRemoveSelected";
            this.btnRemoveSelected.Size = new System.Drawing.Size(100, 30);
            this.btnRemoveSelected.TabIndex = 2;
            this.btnRemoveSelected.Text = "선택 제거";
            this.btnRemoveSelected.UseVisualStyleBackColor = true;
            this.btnRemoveSelected.Click += new System.EventHandler(this.BtnRemoveSelected_Click);
            //
            // btnClear
            //
            this.btnClear.Location = new System.Drawing.Point(318, 4);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(100, 30);
            this.btnClear.TabIndex = 3;
            this.btnClear.Text = "목록 비우기";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.BtnClear_Click);
            //
            // btnCheckUpdate
            //
            this.btnCheckUpdate.Location = new System.Drawing.Point(424, 4);
            this.btnCheckUpdate.Name = "btnCheckUpdate";
            this.btnCheckUpdate.Size = new System.Drawing.Size(100, 30);
            this.btnCheckUpdate.TabIndex = 4;
            this.btnCheckUpdate.Text = "업데이트 확인";
            this.btnCheckUpdate.UseVisualStyleBackColor = true;
            this.btnCheckUpdate.Click += new System.EventHandler(this.BtnCheckUpdate_Click);
            //
            // listFiles
            //
            this.listFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFileName,
            this.columnStatus,
            this.columnPath});
            this.listFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listFiles.FullRowSelect = true;
            this.listFiles.GridLines = true;
            this.listFiles.HideSelection = false;
            this.listFiles.Location = new System.Drawing.Point(15, 59);
            this.listFiles.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.listFiles.Name = "listFiles";
            this.listFiles.Size = new System.Drawing.Size(690, 200);
            this.listFiles.TabIndex = 1;
            this.listFiles.UseCompatibleStateImageBehavior = false;
            this.listFiles.View = System.Windows.Forms.View.Details;
            //
            // columnFileName
            //
            this.columnFileName.Text = "파일명";
            this.columnFileName.Width = 220;
            //
            // columnStatus
            //
            this.columnStatus.Text = "상태";
            this.columnStatus.Width = 120;
            //
            // columnPath
            //
            this.columnPath.Text = "경로";
            this.columnPath.Width = 340;
            //
            // chkOverwrite
            //
            this.chkOverwrite.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chkOverwrite.Location = new System.Drawing.Point(15, 265);
            this.chkOverwrite.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.chkOverwrite.Name = "chkOverwrite";
            this.chkOverwrite.Size = new System.Drawing.Size(690, 24);
            this.chkOverwrite.TabIndex = 2;
            this.chkOverwrite.Text = "원본 덮어쓰기 (기본: 원본은 보존하고 *_unlocked 로 저장)";
            this.chkOverwrite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.chkOverwrite.UseVisualStyleBackColor = true;
            //
            // panelStart
            //
            this.panelStart.Controls.Add(this.btnStart);
            this.panelStart.Controls.Add(this.progressBar);
            this.panelStart.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelStart.Location = new System.Drawing.Point(15, 295);
            this.panelStart.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.panelStart.Name = "panelStart";
            this.panelStart.Size = new System.Drawing.Size(690, 42);
            this.panelStart.TabIndex = 3;
            //
            // btnStart
            //
            this.btnStart.Font = new System.Drawing.Font("맑은 고딕", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnStart.Location = new System.Drawing.Point(0, 0);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(150, 42);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "해제 시작";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.BtnStart_Click);
            //
            // progressBar
            //
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(158, 11);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(532, 20);
            this.progressBar.TabIndex = 1;
            //
            // txtLog
            //
            this.txtLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLog.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtLog.Location = new System.Drawing.Point(15, 343);
            this.txtLog.Margin = new System.Windows.Forms.Padding(0, 0, 0, 6);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(690, 128);
            this.txtLog.TabIndex = 4;
            //
            // panelBottom
            //
            this.panelBottom.Controls.Add(this.lblStatus);
            this.panelBottom.Controls.Add(this.btnOpenFolder);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBottom.Location = new System.Drawing.Point(15, 477);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(690, 28);
            this.panelBottom.TabIndex = 5;
            //
            // lblStatus
            //
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Location = new System.Drawing.Point(0, 5);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(556, 20);
            this.lblStatus.TabIndex = 0;
            this.lblStatus.Text = "대기 중";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // btnOpenFolder
            //
            this.btnOpenFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFolder.Enabled = false;
            this.btnOpenFolder.Location = new System.Drawing.Point(560, 0);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(130, 28);
            this.btnOpenFolder.TabIndex = 1;
            this.btnOpenFolder.Text = "결과 폴더 열기";
            this.btnOpenFolder.UseVisualStyleBackColor = true;
            this.btnOpenFolder.Click += new System.EventHandler(this.BtnOpenFolder_Click);
            //
            // Form1
            //
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 520);
            this.Controls.Add(this.tablePanelMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(600, 420);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "엑셀 보호 해제기";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.tablePanelMain.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelStart.ResumeLayout(false);
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tablePanelMain;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Button btnAddFiles;
        private System.Windows.Forms.Button btnAddFolder;
        private System.Windows.Forms.Button btnRemoveSelected;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnCheckUpdate;
        private System.Windows.Forms.ListView listFiles;
        private System.Windows.Forms.ColumnHeader columnFileName;
        private System.Windows.Forms.ColumnHeader columnStatus;
        private System.Windows.Forms.ColumnHeader columnPath;
        private System.Windows.Forms.CheckBox chkOverwrite;
        private System.Windows.Forms.Panel panelStart;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnOpenFolder;
    }
}
