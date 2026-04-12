// ─────────────────────────────────────────────────────────────────────────────
// File   : MailSender.cs
// Project: TlsMailSender (Class Library, .NET Framework 4.8, AnyCPU/MSIL)
// Purpose: PowerBuilder에서 .NET Assembly Import 또는 COM 방식으로 호출 가능한
//          TLS(STARTTLS) 메일 발송 기능 (첨부파일 지원 포함, 포트 25)
//          인증서 검증: 시스템 기본 검증 + 화이트리스트 기반 예외 허용
// Author :
// Date   : 2025-06-02
// ─────────────────────────────────────────────────────────────────────────────

using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SimpleNetMail
{
    /// <summary>
    /// PowerBuilder에서 .NET Assembly Import 또는 COM 방식으로 호출할 수 있는 메일 발송 클래스입니다.
    /// SMTP 포트 25를 사용하여 STARTTLS(=TLS) 연결을 수행합니다.
    /// 인증서 검증: 시스템 기본 검증을 따르되, AllowedCerts.txt에 등록된 지문은 예외 허용.
    /// </summary>
    [ComVisible(true)]
    [Guid("31ABF072-3366-44CF-8220-467B95BF08B3")]
    [ClassInterface(ClassInterfaceType.None)]
    public class MailSender
    {
        private static HashSet<string> _allowedThumbprints = null;
        private static readonly object _lockObj = new object();
        private static bool _callbackRegistered = false;

        // ── 지문 정규화 ──────────────────────────────────────────────────────────
        // 기존: line.Trim().Replace(" ","").Replace(":","").Replace("-","").ToUpperInvariant()
        //       → 최소 4개의 중간 string 할당 발생
        //
        // 개선: ReadOnlySpan<char>로 받아 단일 char[] 버퍼에 한 번만 순회
        //       → char[] 1개 + 최종 string 1개, 중간 문자열 없음
        //
        // net481에서 new string(ReadOnlySpan<char>)는 사용 불가 (.NET Core 2.1+ 전용)
        // → new string(char[], int, int) 사용
        internal static string NormalizeThumbprint(ReadOnlySpan<char> text)
        {
            if (text.IsEmpty) return string.Empty;
            if (text.Length > 128) return string.Empty;

            char[] buf = new char[text.Length];
            int pos = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == ' ' || c == ':' || c == '-') continue;
                buf[pos++] = char.ToUpperInvariant(c);
            }
            return pos == 0 ? string.Empty : new string(buf, 0, pos);
        }

        // 테스트 접근용 string 오버로드 (Span 호출부의 trim까지 포함)
        internal static string NormalizeThumbprint(string line)
            => NormalizeThumbprint(line.AsSpan().Trim());

        // ── 수신자 파싱 ──────────────────────────────────────────────────────────
        // to / toDisplayName을 ';' 또는 ',' 기준으로 분리해 (주소, 표시이름) 쌍 목록을 반환합니다.
        internal static List<(string Address, string DisplayName)> ParseRecipients(
            string to, string toDisplayName)
        {
            var result = new List<(string Address, string DisplayName)>();
            if (string.IsNullOrWhiteSpace(to))
                return result;

            string[] recipients = to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            string[] displayNames = string.IsNullOrWhiteSpace(toDisplayName)
                ? Array.Empty<string>()
                : toDisplayName.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < recipients.Length; i++)
            {
                string addr = recipients[i].Trim();
                string dn = (i < displayNames.Length && !string.IsNullOrWhiteSpace(displayNames[i]))
                    ? displayNames[i].Trim()
                    : null;
                result.Add((addr, dn));
            }
            return result;
        }

        /// <summary>
        /// 설정 파일에서 허용된 인증서 지문 목록을 로드합니다.
        /// 파일 위치: DLL과 같은 폴더의 AllowedCerts.txt
        /// 형식: 한 줄에 하나의 지문 (SHA-1 또는 SHA-256), #으로 시작하면 주석
        /// </summary>
        private static HashSet<string> LoadAllowedThumbprints()
        {
            var thumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string configPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "AllowedCerts.txt");

                if (File.Exists(configPath))
                {
                    // File.ReadAllLines → string[] 전체 적재를 피해, StreamReader로 한 줄씩 처리합니다.
                    // 각 줄은 string이지만 string[] 배열 할당이 없고, line.AsSpan()으로 추가 중간 문자열 없이
                    // 지문을 정규화합니다.
                    using (var reader = new StreamReader(configPath))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            ReadOnlySpan<char> trimmed = line.AsSpan().Trim();
                            if (trimmed.IsEmpty || trimmed[0] == '#')
                                continue;

                            string normalized = NormalizeThumbprint(trimmed);
                            if (normalized.Length > 0)
                                thumbprints.Add(normalized);
                        }
                    }
                }
                else
                {
                    LogStatic($"[화이트리스트] 설정 파일 없음 - 시스템 검증만 사용");
                }
            }
            catch (Exception ex)
            {
                LogStatic($"[화이트리스트] 로딩 실패: {ex.Message}");
            }
            return thumbprints;
        }

        /// <summary>
        /// 허용된 지문 목록을 가져옵니다 (lazy loading, thread-safe)
        /// </summary>
        private static HashSet<string> AllowedThumbprints
        {
            get
            {
                if (_allowedThumbprints == null)
                {
                    lock (_lockObj)
                    {
                        if (_allowedThumbprints == null)
                            _allowedThumbprints = LoadAllowedThumbprints();
                    }
                }
                return _allowedThumbprints;
            }
        }

        /// <summary>
        /// 허용된 인증서 목록을 다시 로드합니다.
        /// 설정 파일 변경 후 호출하면 새 목록이 적용됩니다.
        /// </summary>
        public void ReloadAllowedCerts()
        {
            lock (_lockObj)
            {
                _allowedThumbprints = LoadAllowedThumbprints();
            }
        }

        /// <summary>
        /// 인증서 검증 콜백을 등록합니다.
        /// </summary>
        private static void EnsureCertificateValidationCallback()
        {
            if (!_callbackRegistered)
            {
                lock (_lockObj)
                {
                    if (!_callbackRegistered)
                    {
                        ServicePointManager.ServerCertificateValidationCallback = ValidateCertificate;
                        _callbackRegistered = true;
                    }
                }
            }
        }

        /// <summary>
        /// 인증서 검증 로직:
        /// 1. 화이트리스트에 있으면 → 무조건 허용 (시스템 검증 무시)
        /// 2. 시스템 검증 통과(sslPolicyErrors == None) → 허용
        /// 3. 그 외 → 거부
        /// </summary>
        private static bool ValidateCertificate(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            string thumbprint = null;
            X509Certificate2 cert2 = null;

            if (certificate != null)
            {
                cert2 = certificate as X509Certificate2 ?? new X509Certificate2(certificate);
                thumbprint = cert2.Thumbprint;
            }

            if (thumbprint != null && AllowedThumbprints.Contains(thumbprint))
            {
                LogStatic($"[인증서 검증] 화이트리스트 허용: {thumbprint}");
                return true;
            }

            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            LogStatic($"=== 인증서 검증 실패 ===");
            LogStatic($"인증서 지문: {thumbprint ?? "(없음)"}");
            LogStatic($"SSL 정책 오류: {sslPolicyErrors}");
            LogStatic($"화이트리스트 등록 여부: {(thumbprint != null ? (AllowedThumbprints.Contains(thumbprint) ? "등록됨" : "미등록") : "확인불가")}");
            LogStatic($"화이트리스트 항목 수: {AllowedThumbprints.Count}");

            if (cert2 != null)
            {
                try
                {
                    LogStatic($"인증서 주체: {cert2.Subject ?? "(없음)"}");
                    LogStatic($"인증서 발급자: {cert2.Issuer ?? "(없음)"}");
                    LogStatic($"인증서 유효기간: {cert2.NotBefore:yyyy-MM-dd} ~ {cert2.NotAfter:yyyy-MM-dd}");
                    LogStatic($"인증서 만료 여부: {(DateTime.Now > cert2.NotAfter ? "만료됨" : "유효함")}");
                }
                catch { }
            }

            if (chain != null && chain.ChainStatus != null && chain.ChainStatus.Length > 0)
            {
                LogStatic($"인증서 체인 상태:");
                foreach (var status in chain.ChainStatus)
                    LogStatic($"  - {status.Status}: {status.StatusInformation}");
            }

            return false;
        }

        /// <summary>
        /// STARTTLS(=TLS) 연결을 지원하는 SMTP 서버로 메일을 발송합니다.
        /// PowerBuilder에서 string[] attachments 배열을 넘겨 첨부파일을 지정할 수 있습니다.
        /// 인증서 검증: 시스템 기본 검증을 수행하며, AllowedCerts.txt에 등록된 지문은 예외적으로 허용합니다.
        /// </summary>
        /// <param name="smtpServer">SMTP 서버 주소 (예: "smtp.example.com")</param>
        /// <param name="smtpPort">SMTP 포트 (포트 25: STARTTLS, 587: SUBMISSION)</param>
        /// <param name="smtpUser">SMTP 인증용 사용자 계정</param>
        /// <param name="smtpPass">SMTP 인증용 비밀번호</param>
        /// <param name="useTLS">true 시 STARTTLS 연결, false 시 평문 전송</param>
        /// <param name="from">보내는 사람 이메일 주소</param>
        /// <param name="to">받는 사람 이메일. 여러 명일 경우 ";" 또는 "," 로 구분</param>
        /// <param name="subject">메일 제목</param>
        /// <param name="body">메일 본문 (HTML 형식)</param>
        /// <param name="attachments">첨부파일 경로 배열. null 또는 길이 0이면 첨부 없음</param>
        /// <returns>성공 시 true, 실패 시 false</returns>
        public bool SendMail(
            string smtpServer,
            int smtpPort,
            string smtpUser,
            string smtpPass,
            bool useTLS,
            string from,
            string to,
            string subject,
            string body,
            string[] attachments)
        {
            return SendMailInternal(
                smtpServer, smtpPort, smtpUser, smtpPass, useTLS,
                from, to, subject, body, attachments,
                null, null);
        }

        /// <summary>
        /// SendMail과 동일하지만 보내는 사람/받는 사람 표시 이름을 지원합니다.
        /// </summary>
        /// <param name="fromDisplayName">보내는 사람 표시 이름 (선택 사항)</param>
        /// <param name="toDisplayName">받는 사람 표시 이름 (선택 사항, 여러 명일 경우 세미콜론으로 구분)</param>
        public bool SendMailWithAlias(
            string smtpServer,
            int smtpPort,
            string smtpUser,
            string smtpPass,
            bool useTLS,
            string from,
            string to,
            string subject,
            string body,
            string[] attachments,
            string fromDisplayName = null,
            string toDisplayName = null)
        {
            return SendMailInternal(
                smtpServer, smtpPort, smtpUser, smtpPass, useTLS,
                from, to, subject, body, attachments,
                fromDisplayName, toDisplayName);
        }

        private bool SendMailInternal(
            string smtpServer,
            int smtpPort,
            string smtpUser,
            string smtpPass,
            bool useTLS,
            string from,
            string to,
            string subject,
            string body,
            string[] attachments,
            string fromDisplayName,
            string toDisplayName)
        {
            try
            {
                if (useTLS)
                    EnsureCertificateValidationCallback();

                // ── 첨부파일 버퍼 추적 ──────────────────────────────────────────────
                // attachmentStreams : MemoryStream이 rented 배열을 감싸고 있으므로
                //                    client.Send 완료 전까지 유지해야 합니다.
                // rentedBuffers     : Send + Dispose 완료 후 풀에 반환합니다.
                // 반환 순서 위반(풀 반환 → Dispose 나중)은 풀 오염을 일으킵니다.
                var attachmentStreams = new List<MemoryStream>();
                var rentedBuffers = new List<byte[]>();
                try
                {
                    using (var message = new MailMessage())
                    {
                        message.From = string.IsNullOrWhiteSpace(fromDisplayName)
                            ? new MailAddress(from)
                            : new MailAddress(from, fromDisplayName);

                        // Message-ID 주입
                        try
                        {
                            string domain = message.From.Host;
                            if (string.IsNullOrEmpty(domain))
                                throw new Exception("발신자 주소에서 도메인을 추출할 수 없습니다.");

                            // Guid.NewGuid().ToString("N") == ToString().Replace("-","")
                            string msgId = $"<{Guid.NewGuid():N}@{domain}>";
                            message.Headers.Set("Message-ID", msgId);
                            Log($"Message-ID 주입: {msgId}");
                        }
                        catch (Exception ex)
                        {
                            Log($"[Message-ID 생성 실패] {ex.Message}");
                        }

                        // 수신자(To) 설정
                        if (!string.IsNullOrWhiteSpace(to))
                        {
                            foreach (var pair in ParseRecipients(to, toDisplayName))
                            {
                                if (string.IsNullOrWhiteSpace(pair.DisplayName))
                                    message.To.Add(new MailAddress(pair.Address));
                                else
                                    message.To.Add(new MailAddress(pair.Address, pair.DisplayName));
                            }
                        }

                        message.Subject = subject ?? string.Empty;
                        message.Body = body ?? string.Empty;
                        message.IsBodyHtml = true;

                        // ── 첨부파일: ArrayPool 기반 버퍼 읽기 ──────────────────────────
                        // File.ReadAllBytes → 파일 크기만큼의 byte[]를 새로 heap에 할당.
                        //                     85,000 bytes 이상이면 LOH(Large Object Heap)에 올라가
                        //                     GC Gen2 수집 전까지 메모리를 점유합니다.
                        //
                        // ArrayPool.Shared.Rent → 풀에서 재사용 가능한 버퍼를 빌려 LOH 적재와
                        //                         heap 단편화를 줄입니다.
                        //                         Rent(n)은 n 이상의 배열을 반환할 수 있으므로
                        //                         MemoryStream 생성 시 실제 파일 크기(bytesRead)만 지정합니다.
                        if (attachments != null && attachments.Length > 0)
                        {
                            foreach (string filePath in attachments)
                            {
                                if (string.IsNullOrWhiteSpace(filePath)) continue;

                                var info = new FileInfo(filePath);
                                if (!info.Exists)
                                {
                                    Log($"첨부파일을 찾을 수 없습니다: {filePath}");
                                    continue;
                                }

                                int fileLen = (int)info.Length;
                                byte[] rented = ArrayPool<byte>.Shared.Rent(fileLen);
                                rentedBuffers.Add(rented);

                                int bytesRead;
                                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096))
                                    bytesRead = ReadFully(fs, rented, fileLen);

                                // writable=false: Attachment는 스트림을 읽기만 합니다.
                                var ms = new MemoryStream(rented, 0, bytesRead, false);
                                attachmentStreams.Add(ms);
                                message.Attachments.Add(new Attachment(ms, Path.GetFileName(filePath)));
                            }
                        }

                        using (var client = new SmtpClient(smtpServer, smtpPort))
                        {
                            client.EnableSsl = useTLS;
                            client.Credentials = new NetworkCredential(smtpUser, smtpPass);
                            client.Timeout = 300_000;
                            client.Send(message);
                        }
                    }
                }
                finally
                {
                    // 1단계: 스트림을 먼저 Dispose — Attachment/SmtpClient가 스트림을 사용 중인 동안
                    //        풀 반환은 금지입니다. using(SmtpClient) 블록이 끝난 뒤 도달합니다.
                    foreach (var s in attachmentStreams)
                        s?.Dispose();

                    // 2단계: 스트림 해제가 완전히 끝난 뒤 풀에 반환합니다.
                    foreach (var b in rentedBuffers)
                        ArrayPool<byte>.Shared.Return(b);
                }

                return true;
            }
            catch (Exception ex)
            {
                Log($"=== 메일 발송 실패 ===");
                Log($"예외 타입: {ex.GetType().Name}");
                Log($"예외 메시지: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Log($"내부 예외 타입: {ex.InnerException.GetType().Name}");
                    Log($"내부 예외 메시지: {ex.InnerException.Message}");
                }
                Log($"SMTP 서버: {smtpServer}:{smtpPort}");
                Log($"TLS 사용: {useTLS}");
                Log($"발신자: {from}");
                Log($"수신자: {to}");
                Log($"제목: {subject ?? "(없음)"}");
                Log($"예외 스택:");
                Log(ex.StackTrace ?? "(스택 트레이스 없음)");
                return false;
            }
        }

        // ── 파일 전체 읽기 ──────────────────────────────────────────────────────
        // FileStream.Read는 요청 크기보다 적게 읽을 수 있으므로 count 바이트를 보장합니다.
        private static int ReadFully(FileStream fs, byte[] buffer, int count)
        {
            int total = 0;
            while (total < count)
            {
                int n = fs.Read(buffer, total, count - total);
                if (n == 0) break;
                total += n;
            }
            return total;
        }

        private void Log(string message) => LogStatic(message);

        private static void LogStatic(string message)
        {
            try
            {
                string logPath = "C:\\temp\\TlsMailSender.log";
                string logDir = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\n");
            }
            catch { }
        }
    }
}
