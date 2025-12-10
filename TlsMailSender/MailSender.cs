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
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // 허용된 인증서 지문 목록 (대문자, 공백/콜론 제거된 형태)
        private static HashSet<string> _allowedThumbprints = null;
        private static readonly object _lockObj = new object();
        private static bool _callbackRegistered = false;

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
                // DLL이 위치한 폴더에서 설정 파일 찾기
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string configPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "AllowedCerts.txt");

                LogStatic($"[화이트리스트] DLL 경로: {assemblyPath}");
                LogStatic($"[화이트리스트] 설정 파일 경로: {configPath}");

                if (File.Exists(configPath))
                {
                    LogStatic($"[화이트리스트] 설정 파일 발견, 로딩 중...");
                    
                    foreach (string line in File.ReadAllLines(configPath))
                    {
                        string trimmed = line.Trim();

                        // 빈 줄이나 주석은 건너뜀
                        if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#"))
                            continue;

                        // 지문 정규화: 공백, 콜론, 하이픈 제거 후 대문자로
                        string normalized = trimmed
                            .Replace(" ", "")
                            .Replace(":", "")
                            .Replace("-", "")
                            .ToUpperInvariant();

                        if (normalized.Length > 0)
                        {
                            thumbprints.Add(normalized);
                            LogStatic($"[화이트리스트] 지문 등록: {normalized}");
                        }
                    }
                    
                    LogStatic($"[화이트리스트] 총 {thumbprints.Count}개 지문 로딩 완료");
                }
                else
                {
                    LogStatic($"[화이트리스트] 설정 파일 없음 - 시스템 검증만 사용");
                }
            }
            catch (Exception ex)
            {
                LogStatic($"[화이트리스트] 로딩 실패: {ex.Message}");
                // 설정 파일 로드 실패 시 빈 목록 사용 (시스템 검증만 수행)
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
                        {
                            _allowedThumbprints = LoadAllowedThumbprints();
                        }
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

            // 인증서 지문 추출
            if (certificate != null)
            {
                cert2 = certificate as X509Certificate2 ?? new X509Certificate2(certificate);
                thumbprint = cert2.Thumbprint; // 이미 대문자, 공백 없음
            }

            // 1. 화이트리스트 먼저 체크 - 등록된 지문이면 무조건 허용
            if (thumbprint != null && AllowedThumbprints.Contains(thumbprint))
            {
                LogStatic($"[인증서 검증] 화이트리스트 허용: {thumbprint}");
                return true;
            }

            // 2. 시스템 검증 통과 시 허용
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            // 3. 둘 다 실패 - 상세 정보 로깅 후 거부
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
                catch
                {
                    // 인증서 정보 읽기 실패 시 무시
                }
            }

            if (chain != null && chain.ChainStatus != null && chain.ChainStatus.Length > 0)
            {
                LogStatic($"인증서 체인 상태:");
                foreach (var status in chain.ChainStatus)
                {
                    LogStatic($"  - {status.Status}: {status.StatusInformation}");
                }
            }

            return false;
        }

        /// <summary>
        /// STARTTLS(=TLS) 연결을 지원하는 SMTP 서버로 메일을 발송합니다.
        /// PowerBuilder에서 string[] attachments 배열을 넘겨 첨부파일을 지정할 수 있습니다.
        /// 인증서 검증: 시스템 기본 검증을 수행하며, AllowedCerts.txt에 등록된 지문은 예외적으로 허용합니다.
        /// </summary>
        /// <param name="smtpServer">
        /// SMTP 서버 주소 (예: "smtp.example.com"). 
        /// 포트 25를 사용할 때는 해당 서버가 STARTTLS를 지원해야 합니다.
        /// </param>
        /// <param name="smtpPort">
        /// SMTP 포트. 
        /// 포트 25를 사용할 경우 STARTTLS 핸드셰이크가 수행됩니다 (EnableSsl=true).
        /// 일반적으로 포트 587이나 465 대신, ISP/사내 환경에서 25번 포트를 사용할 때 이 파라미터를 25로 지정합니다.
        /// </param>
        /// <param name="smtpUser">SMTP 인증용 사용자 계정 (예: "user@example.com")</param>
        /// <param name="smtpPass">SMTP 인증용 비밀번호 (앱 전용 비밀번호 권장)</param>
        /// <param name="useTLS">
        /// true로 설정 시 STARTTLS(=TLS) 연결을 시도합니다. 
        /// false로 설정 시 암호화 없이 평문으로 전송되므로 보안상 권장되지 않습니다.
        /// </param>
        /// <param name="from">보내는 사람 이메일 주소 (예: "user@example.com")</param>
        /// <param name="to">
        /// 받는 사람 이메일. 여러 명일 경우 ";" 또는 "," 로 구분 (예: "a@domain.com;b@domain.com").
        /// </param>
        /// <param name="subject">메일 제목 (예: "테스트 메일")</param>
        /// <param name="body">메일 본문 (HTML 형식). IsBodyHtml이 true로 설정되어 있습니다.</param>
        /// <param name="attachments">
        /// 첨부파일 경로 배열 (PowerBuilder에서 string[] 형태로 만들어서 넘김). 
        /// 배열이 null이거나 길이가 0이면 첨부 없음.
        /// </param>
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
            string[] attachments
        )
        {
            try
            {


                // ── TLS 사용 시 인증서 검증: 시스템 검증 + 화이트리스트 기반 예외 허용 ───────────────
                // - 정상 인증서: 시스템 검증 통과 시 허용
                // - 사설 인증서: AllowedCerts.txt에 지문 등록 시 허용
                if (useTLS) {
                    EnsureCertificateValidationCallback();
                }

                // 2) MailMessage 객체 생성 및 기본 설정
                List<MemoryStream> attachmentStreams = new List<MemoryStream>();
                try
                {
                    using (MailMessage message = new MailMessage())
                    {
                        message.From = new MailAddress(from);

                        // 3) 받는 사람(To) 설정
                        if (!string.IsNullOrWhiteSpace(to))
                        {
                            string[] recipients = to
                                .Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string addr in recipients)
                            {
                                message.To.Add(addr.Trim());
                            }
                        }

                        // 4) 제목/본문 설정
                        message.Subject = subject ?? string.Empty;
                        message.Body = body ?? string.Empty;
                        message.IsBodyHtml = true; // HTML 메일로 전송됩니다

                        // 5) 첨부파일 처리 - 파일을 메모리 스트림에 명시적으로 로드
                        if (attachments != null && attachments.Length > 0)
                        {
                            foreach (string path in attachments)
                            {
                                if (string.IsNullOrWhiteSpace(path))
                                    continue;

                                if (File.Exists(path))
                                {
                                    // 파일을 메모리 스트림에 명시적으로 로드
                                    byte[] fileBytes = File.ReadAllBytes(path);
                                    MemoryStream fileStream = new MemoryStream(fileBytes);
                                    attachmentStreams.Add(fileStream); // 메일 발송 완료까지 유지
                                    
                                    Attachment attach = new Attachment(fileStream, Path.GetFileName(path));
                                    message.Attachments.Add(attach);
                                }
                                else
                                {
                                    // 경로 오류나 파일이 없는 경우 무시
                                    // (PB에서 사전에 경로 유효성 검사하는 것을 권장)
                                    Log($"첨부파일을 찾을 수 없습니다: {path}");
                                }
                            }
                        }

                        // 6) SmtpClient 설정 및 메일 전송
                        using (SmtpClient client = new SmtpClient(smtpServer, smtpPort))
                        {
                            // STARTTLS(=TLS) 사용 여부를 지정.
                            // 포트 25, useTLS=true일 때, 서버가 EHLO 후 STARTTLS를 지원하면
                            // 암호화 연결로 전환합니다.
                            client.EnableSsl = useTLS;

                            client.Credentials = new NetworkCredential(smtpUser, smtpPass);
                            client.Timeout = 300_000; // 타임아웃 5분

                            client.Send(message);
                        }
                    }
                }
                finally
                {
                    // 첨부파일 스트림 정리
                    foreach (var stream in attachmentStreams)
                    {
                        stream?.Dispose();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                // 예외 발생 시 상세 정보 로깅
                Log($"=== 메일 발송 실패 ===");
                Log($"예외 타입: {ex.GetType().Name}");
                Log($"예외 메시지: {ex.Message}");
                
                // Inner Exception이 있으면 함께 로깅
                if (ex.InnerException != null)
                {
                    Log($"내부 예외 타입: {ex.InnerException.GetType().Name}");
                    Log($"내부 예외 메시지: {ex.InnerException.Message}");
                }
                
                // SMTP 연결 정보 로깅 (민감 정보 제외)
                Log($"SMTP 서버: {smtpServer}:{smtpPort}");
                Log($"TLS 사용: {useTLS}");
                Log($"발신자: {from}");
                Log($"수신자: {to}");
                Log($"제목: {subject ?? "(없음)"}");
                
                // 스택 트레이스 로깅
                Log($"예외 스택:");
                Log(ex.StackTrace ?? "(스택 트레이스 없음)");
                
                return false;
            }
        }

        /// <summary>
        /// STARTTLS(=TLS) 연결을 지원하는 SMTP 서버로 메일을 발송합니다. 별칭(표시 이름)을 지원합니다.
        /// PowerBuilder에서 string[] attachments 배열을 넘겨 첨부파일을 지정할 수 있습니다.
        /// 인증서 검증: 시스템 기본 검증을 수행하며, AllowedCerts.txt에 등록된 지문은 예외적으로 허용합니다.
        /// </summary>
        /// <param name="smtpServer">
        /// SMTP 서버 주소 (예: "smtp.example.com"). 
        /// 포트 25를 사용할 때는 해당 서버가 STARTTLS를 지원해야 합니다.
        /// </param>
        /// <param name="smtpPort">
        /// SMTP 포트. 
        /// 포트 25를 사용할 경우 STARTTLS 핸드셰이크가 수행됩니다 (EnableSsl=true).
        /// 일반적으로 포트 587이나 465 대신, ISP/사내 환경에서 25번 포트를 사용할 때 이 파라미터를 25로 지정합니다.
        /// </param>
        /// <param name="smtpUser">SMTP 인증용 사용자 계정 (예: "user@example.com")</param>
        /// <param name="smtpPass">SMTP 인증용 비밀번호 (앱 전용 비밀번호 권장)</param>
        /// <param name="useTLS">
        /// true로 설정 시 STARTTLS(=TLS) 연결을 시도합니다. 
        /// false로 설정 시 암호화 없이 평문으로 전송되므로 보안상 권장되지 않습니다.
        /// </param>
        /// <param name="from">보내는 사람 이메일 주소 (예: "user@example.com")</param>
        /// <param name="to">
        /// 받는 사람 이메일. 여러 명일 경우 ";" 또는 "," 로 구분 (예: "a@domain.com;b@domain.com").
        /// </param>
        /// <param name="subject">메일 제목 (예: "테스트 메일")</param>
        /// <param name="body">메일 본문 (HTML 형식). IsBodyHtml이 true로 설정되어 있습니다.</param>
        /// <param name="attachments">
        /// 첨부파일 경로 배열 (PowerBuilder에서 string[] 형태로 만들어서 넘김). 
        /// 배열이 null이거나 길이가 0이면 첨부 없음.
        /// </param>
        /// <param name="fromDisplayName">보내는 사람 표시 이름 (선택 사항)</param>
        /// <param name="toDisplayName">받는 사람 표시 이름 (선택 사항, 여러 명일 경우 세미콜론으로 구분)</param>
        /// <returns>성공 시 true, 실패 시 false</returns>
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
            string toDisplayName = null
        )
        {
            try
            {

                // ── TLS 사용 시 인증서 검증: 시스템 검증 + 화이트리스트 기반 예외 허용 ───────────────
                // - 정상 인증서: 시스템 검증 통과 시 허용
                // - 사설 인증서: AllowedCerts.txt에 지문 등록 시 허용
                if (useTLS) {
                    EnsureCertificateValidationCallback();
                }

                // 2) MailMessage 객체 생성 및 기본 설정
                List<MemoryStream> attachmentStreams = new List<MemoryStream>();
                try
                {
                    using (MailMessage message = new MailMessage())
                    {
                        message.From = string.IsNullOrWhiteSpace(fromDisplayName) ? new MailAddress(from) : new MailAddress(from, fromDisplayName);

                        // 3) 받는 사람(To) 설정
                        if (!string.IsNullOrWhiteSpace(to))
                        {
                            string[] recipients = to
                                .Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            string[] displayNames = string.IsNullOrWhiteSpace(toDisplayName) ? new string[0] : toDisplayName
                                .Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

                            for (int i = 0; i < recipients.Length; i++)
                            {
                                string addr = recipients[i].Trim();
                                string displayName = (i < displayNames.Length && !string.IsNullOrWhiteSpace(displayNames[i])) ? displayNames[i].Trim() : null;

                                if (string.IsNullOrWhiteSpace(displayName))
                                {
                                    message.To.Add(new MailAddress(addr));
                                }
                                else
                                {
                                    message.To.Add(new MailAddress(addr, displayName));
                                }
                            }
                        }

                        // 4) 제목/본문 설정
                        message.Subject = subject ?? string.Empty;
                        message.Body = body ?? string.Empty;
                        message.IsBodyHtml = true; // HTML 메일로 전송됩니다

                        // 5) 첨부파일 처리 - 파일을 메모리 스트림에 명시적으로 로드
                        if (attachments != null && attachments.Length > 0)
                        {
                            foreach (string path in attachments)
                            {
                                if (string.IsNullOrWhiteSpace(path))
                                    continue;

                                if (File.Exists(path))
                                {
                                    // 파일을 메모리 스트림에 명시적으로 로드
                                    byte[] fileBytes = File.ReadAllBytes(path);
                                    MemoryStream fileStream = new MemoryStream(fileBytes);
                                    attachmentStreams.Add(fileStream); // 메일 발송 완료까지 유지
                                    
                                    Attachment attach = new Attachment(fileStream, Path.GetFileName(path));
                                    message.Attachments.Add(attach);
                                }
                                else
                                {
                                    // 경로 오류나 파일이 없는 경우 무시
                                    // (PB에서 사전에 경로 유효성 검사하는 것을 권장)
                                    Log($"첨부파일을 찾을 수 없습니다: {path}");
                                }
                            }
                        }

                        // 6) SmtpClient 설정 및 메일 전송
                        using (SmtpClient client = new SmtpClient(smtpServer, smtpPort))
                        {
                            // STARTTLS(=TLS) 사용 여부를 지정.
                            // 포트 25, useTLS=true일 때, 서버가 EHLO 후 STARTTLS를 지원하면
                            // 암호화 연결로 전환합니다.
                            client.EnableSsl = useTLS;

                            client.Credentials = new NetworkCredential(smtpUser, smtpPass);
                            client.Timeout = 300_000; // 타임아웃 5분

                            client.Send(message);
                        }
                    }
                }
                finally
                {
                    // 첨부파일 스트림 정리
                    foreach (var stream in attachmentStreams)
                    {
                        stream?.Dispose();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                // 예외 발생 시 상세 정보 로깅
                Log($"=== 메일 발송 실패 ===");
                Log($"예외 타입: {ex.GetType().Name}");
                Log($"예외 메시지: {ex.Message}");
                
                // Inner Exception이 있으면 함께 로깅
                if (ex.InnerException != null)
                {
                    Log($"내부 예외 타입: {ex.InnerException.GetType().Name}");
                    Log($"내부 예외 메시지: {ex.InnerException.Message}");
                }
                
                // SMTP 연결 정보 로깅 (민감 정보 제외)
                Log($"SMTP 서버: {smtpServer}:{smtpPort}");
                Log($"TLS 사용: {useTLS}");
                Log($"발신자: {from}");
                Log($"수신자: {to}");
                Log($"제목: {subject ?? "(없음)"}");
                
                // 스택 트레이스 로깅
                Log($"예외 스택:");
                Log(ex.StackTrace ?? "(스택 트레이스 없음)");
                
                return false;
            }
        }

        private void Log(string message)
        {
            LogStatic(message);
        }

        private static void LogStatic(string message)
        {
            try
            {
                string logPath = "C:\\temp\\TlsMailSender.log";
                string logDir = Path.GetDirectoryName(logPath);
                
                // 로그 디렉터리가 없으면 생성
                if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\n");
            }
            catch
            {
                // 로그 파일 쓰기 실패 시 무시 (예외 전파 방지)
            }
        }
    }
}
