using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;

namespace Unlocker_For_ExcelSheet.Services
{
    /// <summary>해제 처리 결과 상태.</summary>
    public enum UnlockStatus
    {
        Success,            // 보호를 제거하고 파일을 저장함
        NoProtectionFound,  // 제거할 보호가 없었음(원본 그대로)
        PasswordRequired,   // 열기 암호가 걸려 있어 암호 입력이 필요함
        WrongPassword,      // 입력한 열기 암호가 틀림
        Failed              // 오류/미지원
    }

    /// <summary>해제 동작 옵션.</summary>
    public sealed class UnlockOptions
    {
        /// <summary>true 면 원본을 덮어쓰고, false 면 "원본명_unlocked.확장자" 로 저장.</summary>
        public bool OverwriteOriginal { get; set; }

        /// <summary>열기 암호(암호화된 파일). 모르면 null/빈 문자열.</summary>
        public string OpenPassword { get; set; }
    }

    /// <summary>파일 한 개의 해제 결과.</summary>
    public sealed class UnlockResult
    {
        public string SourcePath { get; set; }
        public string OutputPath { get; set; }
        public UnlockStatus Status { get; set; }
        public int SheetProtectionsRemoved { get; set; }
        public bool WorkbookProtectionRemoved { get; set; }
        public bool FileSharingRemoved { get; set; }
        public bool WasEncrypted { get; set; }
        public string Message { get; set; }
        public Exception Error { get; set; }

        public string FileName => SourcePath == null ? string.Empty : Path.GetFileName(SourcePath);
    }

    /// <summary>
    /// 엑셀 보호 해제 엔진.
    /// - OOXML(.xlsx/.xlsm/.xltx/.xltm): 시트 보호(sheetProtection), 통합문서 보호
    ///   (workbookProtection), 쓰기 예약(fileSharing) 요소를 XML 파싱으로 정확히 제거.
    /// - 열기 암호로 암호화된 OOXML(CFB 컨테이너): NPOI 로 복호화 후 위 보호 제거.
    /// - 레거시 .xls(BIFF): 감지하여 안내(현재 미지원).
    /// 원본은 항상 보존한다(임시 파일에서 작업 후 성공 시에만 결과를 저장).
    /// </summary>
    public sealed class ExcelUnlocker
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        // NPOI 의 EncryptionInfo.GetBuilder 는 빌더 타입을 "현재 AppDomain 에 로드된
        // 어셈블리 목록"에서만 찾는다. 지연 로딩 탓에 빌더 어셈블리가 아직 로드되지
        // 않았으면 "Not found type NPOI.POIFS.Crypt.*.*EncryptionInfoBuilder" 예외가
        // 나면서 복호화가 실패한다. 정적 생성자에서 빌더 타입들을 미리 참조해
        // 해당 어셈블리를 강제로 로드시켜 이 문제를 예방한다.
        static ExcelUnlocker()
        {
            var _ = new[]
            {
                typeof(NPOI.POIFS.Crypt.Agile.AgileEncryptionInfoBuilder),
                typeof(NPOI.POIFS.Crypt.Standard.StandardEncryptionInfoBuilder),
                typeof(NPOI.POIFS.Crypt.CryptoAPI.CryptoAPIEncryptionInfoBuilder)
            };
        }

        private enum FileKind { Ooxml, EncryptedOoxml, LegacyXls, Unknown }

        public Task<UnlockResult> UnlockAsync(string filePath, UnlockOptions options, CancellationToken ct = default)
        {
            return Task.Run(() => Unlock(filePath, options), ct);
        }

        public UnlockResult Unlock(string filePath, UnlockOptions options)
        {
            options ??= new UnlockOptions();
            var result = new UnlockResult { SourcePath = filePath, Status = UnlockStatus.Failed };

            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    result.Message = "파일을 찾을 수 없습니다.";
                    return result;
                }

                switch (DetectKind(filePath))
                {
                    case FileKind.Ooxml:
                        return UnlockOoxml(filePath, options);

                    case FileKind.EncryptedOoxml:
                        return UnlockEncrypted(filePath, options);

                    case FileKind.LegacyXls:
                        result.Status = UnlockStatus.Failed;
                        result.Message = "레거시 .xls(97-2003) 형식은 아직 지원하지 않습니다. " +
                                         "엑셀에서 .xlsx 로 저장한 뒤 다시 시도하세요.";
                        return result;

                    default:
                        result.Message = "지원하지 않는 파일 형식입니다.";
                        return result;
                }
            }
            catch (Exception ex)
            {
                result.Status = UnlockStatus.Failed;
                result.Error = ex;
                result.Message = "처리 중 오류: " + ex.Message;
                return result;
            }
        }

        // ── 형식 감지 ─────────────────────────────────────────────
        private static FileKind DetectKind(string path)
        {
            byte[] head = new byte[8];
            int n;
            using (var fs = File.OpenRead(path))
            {
                n = fs.Read(head, 0, head.Length);
            }

            // ZIP(=OOXML) : "PK"
            if (n >= 2 && head[0] == 0x50 && head[1] == 0x4B)
                return FileKind.Ooxml;

            // CFB(OLE2) : D0 CF 11 E0 A1 B1 1A E1  → 암호화 OOXML 또는 레거시 .xls
            bool isCfb = n >= 8 &&
                         head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0 &&
                         head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1;
            if (!isCfb)
                return FileKind.Unknown;

            using (var fs = File.OpenRead(path))
            {
                // POIFSFileSystem 은 IDisposable 이 아니므로 using 대상이 아니다.
                // 파일 핸들은 바깥 FileStream(using)이 정리한다.
                var poifs = new POIFSFileSystem(fs);
                if (RootHasEntry(poifs, "EncryptionInfo"))
                    return FileKind.EncryptedOoxml;
                if (RootHasEntry(poifs, "Workbook") || RootHasEntry(poifs, "Book"))
                    return FileKind.LegacyXls;
                return FileKind.Unknown;
            }
        }

        private static bool RootHasEntry(POIFSFileSystem poifs, string name)
        {
            try
            {
                foreach (var entryName in poifs.Root.EntryNames)
                {
                    if (string.Equals(entryName, name, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // 진입 실패 시 없음으로 간주
            }
            return false;
        }

        // ── OOXML(비암호화) ───────────────────────────────────────
        private UnlockResult UnlockOoxml(string filePath, UnlockOptions options)
        {
            var result = new UnlockResult { SourcePath = filePath };

            string temp = NewTempPath(filePath);
            try
            {
                File.Copy(filePath, temp, overwrite: true);
                // File.Copy 는 원본의 ReadOnly 속성까지 복사한다. ReadOnly 인 채로는
                // ZipFile.Open(Update) 가 UnauthorizedAccessException 을 던지므로 해제.
                ClearReadOnly(temp);
                StripProtections(temp, result);

                bool removedSomething = result.SheetProtectionsRemoved > 0 ||
                                        result.WorkbookProtectionRemoved ||
                                        result.FileSharingRemoved;

                if (!removedSomething)
                {
                    result.Status = UnlockStatus.NoProtectionFound;
                    result.Message = "보호된 시트/통합문서가 없습니다.";
                    result.OutputPath = null;
                    return result;
                }

                string finalPath = FinalPath(filePath, options);
                MoveOver(temp, finalPath);
                result.OutputPath = finalPath;
                result.Status = UnlockStatus.Success;
                result.Message = BuildSuccessMessage(result);
                return result;
            }
            finally
            {
                SafeDelete(temp);
            }
        }

        // ── OOXML(열기 암호 암호화) ───────────────────────────────
        private UnlockResult UnlockEncrypted(string filePath, UnlockOptions options)
        {
            var result = new UnlockResult { SourcePath = filePath, WasEncrypted = true };

            if (string.IsNullOrEmpty(options.OpenPassword))
            {
                result.Status = UnlockStatus.PasswordRequired;
                result.Message = "이 파일은 열기 암호로 암호화되어 있습니다. 암호를 입력하세요.";
                return result;
            }

            string temp = NewTempPath(filePath);
            try
            {
                // 1) 복호화 → 임시 파일로 저장
                using (var stream = File.OpenRead(filePath))
                {
                    var poifs = new POIFSFileSystem(stream);
                    var info = new EncryptionInfo(poifs);
                    var decryptor = Decryptor.GetInstance(info);

                    if (!decryptor.VerifyPassword(options.OpenPassword))
                    {
                        result.Status = UnlockStatus.WrongPassword;
                        result.Message = "암호가 올바르지 않습니다.";
                        return result;
                    }

                    using (var dataStream = decryptor.GetDataStream(poifs))
                    using (var outFs = File.Create(temp))
                    {
                        dataStream.CopyTo(outFs);
                    }
                }

                // 2) 복호화된 OOXML 에서 남은 보호(시트/통합문서)도 함께 제거
                StripProtections(temp, result);

                // 3) 저장 (열기 암호가 없는 평문 파일)
                string finalPath = FinalPath(filePath, options);
                MoveOver(temp, finalPath);
                result.OutputPath = finalPath;
                result.Status = UnlockStatus.Success;
                result.Message = BuildSuccessMessage(result);
                return result;
            }
            finally
            {
                SafeDelete(temp);
            }
        }

        // ── 보호 요소 제거(ZIP in-place 편집) ─────────────────────
        private void StripProtections(string xlsxPath, UnlockResult result)
        {
            using (var zip = ZipFile.Open(xlsxPath, ZipArchiveMode.Update))
            {
                // 편집 중 컬렉션이 변하므로 스냅샷을 뜬다
                foreach (var entry in zip.Entries.ToList())
                {
                    string name = entry.FullName;

                    bool isWorksheet =
                        (name.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) ||
                         name.StartsWith("xl/chartsheets/", StringComparison.OrdinalIgnoreCase)) &&
                        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

                    bool isWorkbook = name.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);

                    if (!isWorksheet && !isWorkbook)
                        continue;

                    string content;
                    using (var s = entry.Open())
                    using (var r = new StreamReader(s, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        content = r.ReadToEnd();
                    }

                    XDocument doc;
                    try
                    {
                        doc = XDocument.Parse(content, LoadOptions.PreserveWhitespace);
                    }
                    catch
                    {
                        continue; // XML 이 아니거나 파싱 실패 → 건너뜀
                    }

                    int changed = 0;

                    if (isWorksheet)
                    {
                        var toRemove = doc.Descendants()
                                          .Where(e => e.Name.LocalName == "sheetProtection")
                                          .ToList();
                        foreach (var e in toRemove) { e.Remove(); changed++; }
                        result.SheetProtectionsRemoved += toRemove.Count;
                    }

                    if (isWorkbook)
                    {
                        var wbProt = doc.Descendants()
                                        .Where(e => e.Name.LocalName == "workbookProtection")
                                        .ToList();
                        foreach (var e in wbProt) { e.Remove(); changed++; }
                        if (wbProt.Count > 0) result.WorkbookProtectionRemoved = true;

                        var fileSharing = doc.Descendants()
                                             .Where(e => e.Name.LocalName == "fileSharing")
                                             .ToList();
                        foreach (var e in fileSharing) { e.Remove(); changed++; }
                        if (fileSharing.Count > 0) result.FileSharingRemoved = true;
                    }

                    if (changed == 0)
                        continue;

                    // 항목 크기가 변하므로 삭제 후 동일 이름으로 재생성
                    entry.Delete();
                    var newEntry = zip.CreateEntry(name, CompressionLevel.Optimal);
                    using (var ws = newEntry.Open())
                    using (var sw = new StreamWriter(ws, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                    {
                        doc.Save(sw, SaveOptions.DisableFormatting);
                    }
                }
            }
        }

        // ── 경로/파일 헬퍼 ────────────────────────────────────────
        private static string FinalPath(string source, UnlockOptions options)
        {
            if (options.OverwriteOriginal)
                return source;

            string dir = Path.GetDirectoryName(source) ?? ".";
            string name = Path.GetFileNameWithoutExtension(source);
            string ext = Path.GetExtension(source);
            return Path.Combine(dir, name + "_unlocked" + ext);
        }

        private static string NewTempPath(string source)
        {
            string ext = Path.GetExtension(source);
            if (string.IsNullOrEmpty(ext)) ext = ".xlsx";
            return Path.Combine(Path.GetTempPath(), "unlocker_" + Guid.NewGuid().ToString("N") + ext);
        }

        private static void MoveOver(string temp, string finalPath)
        {
            // 대상이 ReadOnly 면 덮어쓰기가 실패하므로 먼저 속성 해제.
            ClearReadOnly(finalPath);

            // 제자리 덮어쓰기(File.Copy)는 원자적이지 않아, 중단되면 원본이 손상될 수
            // 있다("원본 보존" 계약 위반). 같은 폴더(동일 볼륨)에 스테이징 복사 후
            // File.Move 로 원자적으로 교체한다.
            string dir = Path.GetDirectoryName(finalPath);
            if (string.IsNullOrEmpty(dir)) dir = ".";
            string stage = Path.Combine(dir, ".unlocker_" + Guid.NewGuid().ToString("N") + Path.GetExtension(finalPath));

            File.Copy(temp, stage, overwrite: true);
            try
            {
                File.Move(stage, finalPath, overwrite: true); // 동일 볼륨이면 원자적 교체
            }
            finally
            {
                SafeDelete(stage);
            }
        }

        private static void ClearReadOnly(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    var attr = File.GetAttributes(path);
                    if ((attr & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                        File.SetAttributes(path, attr & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
                // 속성 변경 실패는 무시(후속 작업에서 오류가 드러남)
            }
        }

        private static void SafeDelete(string path)
        {
            try
            {
                if (path != null && File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // 정리 실패는 무시
            }
        }

        private static string BuildSuccessMessage(UnlockResult r)
        {
            var parts = new List<string>();
            if (r.WasEncrypted) parts.Add("열기 암호 해제");
            if (r.SheetProtectionsRemoved > 0) parts.Add($"시트 보호 {r.SheetProtectionsRemoved}개 제거");
            if (r.WorkbookProtectionRemoved) parts.Add("통합문서 보호 제거");
            if (r.FileSharingRemoved) parts.Add("쓰기 예약 암호 제거");
            return parts.Count > 0 ? string.Join(", ", parts) + " 완료" : "완료";
        }
    }
}
