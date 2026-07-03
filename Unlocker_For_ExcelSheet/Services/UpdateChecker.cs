using System;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Unlocker_For_ExcelSheet.Services
{
    /// <summary>업데이트 확인 결과.</summary>
    public sealed class UpdateInfo
    {
        public bool UpdateAvailable { get; set; }
        public bool CheckFailed { get; set; }
        public string CurrentVersion { get; set; }
        public string LatestVersion { get; set; }
        public string ReleaseUrl { get; set; }     // 릴리즈 페이지(html_url)
        public string DownloadUrl { get; set; }    // 대표 자산(exe/zip) 직접 다운로드 URL (없으면 null)
        public string ReleaseNotes { get; set; }    // 릴리즈 본문
        public string Message { get; set; }         // 사람이 읽는 요약
    }

    /// <summary>
    /// GitHub Releases 기반 업데이트 확인기.
    /// 최신 릴리즈의 태그(vX.Y.Z)를 현재 어셈블리 버전과 비교한다.
    /// 자동 교체(self-update)는 하지 않고, 다운로드 URL/릴리즈 페이지를 제공한다.
    /// </summary>
    public static class UpdateChecker
    {
        public const string Owner = "BaeTab";
        public const string Repo = "Unlocker_For_ExcelSheet";

        private static readonly string LatestReleaseApi =
            $"https://api.github.com/repos/{Owner}/{Repo}/releases/latest";

        public static string CurrentVersion
        {
            get
            {
                var v = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0, 0);
                return $"{v.Major}.{v.Minor}.{v.Build}";
            }
        }

        public static async Task<UpdateInfo> CheckAsync(CancellationToken ct = default)
        {
            var info = new UpdateInfo { CurrentVersion = CurrentVersion };

            try
            {
                using var http = new HttpClient();
                http.Timeout = TimeSpan.FromSeconds(15);
                http.DefaultRequestHeaders.UserAgent.ParseAdd("Unlocker_For_ExcelSheet-UpdateChecker");
                http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");

                using var resp = await http.GetAsync(LatestReleaseApi, ct).ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    info.CheckFailed = true;
                    info.Message = resp.StatusCode == System.Net.HttpStatusCode.NotFound
                        ? "아직 게시된 릴리즈가 없습니다."
                        : $"업데이트 확인 실패 (HTTP {(int)resp.StatusCode}).";
                    return info;
                }

                string json = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                string tag = root.TryGetProperty("tag_name", out var t) ? t.GetString() : null;
                info.LatestVersion = NormalizeVersionString(tag);

                // tag_name 이 없거나 파싱 불가면 "확인 불가"이지 "최신"이 아니다.
                // 가드 없이 아래로 내려가면 UpdateAvailable=false + "최신 버전" 으로
                // 오판(무음 실패)하게 되므로 여기서 CheckFailed 로 처리.
                if (string.IsNullOrEmpty(info.LatestVersion))
                {
                    info.CheckFailed = true;
                    info.Message = "업데이트 확인 실패: 릴리즈 정보 형식이 예상과 다릅니다.";
                    return info;
                }

                info.ReleaseUrl = root.TryGetProperty("html_url", out var h) ? h.GetString() : null;
                info.ReleaseNotes = root.TryGetProperty("body", out var b) ? b.GetString() : null;
                info.DownloadUrl = PickAssetUrl(root);

                var latest = ParseVersion(info.LatestVersion);
                var current = ParseVersion(CurrentVersion);

                if (latest != null && current != null && latest > current)
                {
                    info.UpdateAvailable = true;
                    info.Message = $"새 버전 v{info.LatestVersion} 이(가) 있습니다. (현재 v{info.CurrentVersion})";
                }
                else
                {
                    info.UpdateAvailable = false;
                    info.Message = $"최신 버전을 사용 중입니다. (v{info.CurrentVersion})";
                }

                return info;
            }
            catch (OperationCanceledException)
            {
                info.CheckFailed = true;
                info.Message = "업데이트 확인이 취소되었습니다.";
                return info;
            }
            catch (Exception ex)
            {
                info.CheckFailed = true;
                info.Message = "업데이트 확인 중 오류: " + ex.Message;
                return info;
            }
        }

        /// <summary>릴리즈 자산 중 설치/실행 가능한 대표 파일 URL 을 고른다.</summary>
        private static string PickAssetUrl(JsonElement root)
        {
            if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
                return null;

            string best = null;
            foreach (var asset in assets.EnumerateArray())
            {
                string name = asset.TryGetProperty("name", out var n) ? (n.GetString() ?? "") : "";
                string url = asset.TryGetProperty("browser_download_url", out var u) ? u.GetString() : null;
                if (string.IsNullOrEmpty(url)) continue;

                string lower = name.ToLowerInvariant();
                if (lower.EndsWith(".exe")) return url;             // exe 최우선
                if (lower.EndsWith(".zip") && best == null) best = url;
                if (lower.EndsWith(".msi") && best == null) best = url;
            }
            return best;
        }

        private static string NormalizeVersionString(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return null;
            tag = tag.Trim();
            if (tag.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                tag = tag.Substring(1);
            return tag;
        }

        private static Version ParseVersion(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;

            // "2.1", "2.1.0", "2.1.0.3" 등 허용. 접미사(-beta 등)는 잘라낸다.
            int dash = s.IndexOfAny(new[] { '-', '+' });
            if (dash >= 0) s = s.Substring(0, dash);

            var nums = s.Split('.')
                        .Select(p => int.TryParse(p, out int x) ? x : 0)
                        .ToArray();

            int major = nums.Length > 0 ? nums[0] : 0;
            int minor = nums.Length > 1 ? nums[1] : 0;
            int build = nums.Length > 2 ? nums[2] : 0;
            return new Version(major, minor, build);
        }
    }
}
