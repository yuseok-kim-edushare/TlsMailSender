#if NET
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace SimpleNetMail
{
    /// <summary>
    /// Native AOT(C-Export) 진입점. net10.0-windows(NativeLib=Shared) 게시 시
    /// UTF-16(wchar_t*) 포인터로 MSVC++에서 __stdcall 호출할 수 있는 함수를 노출합니다.
    /// 반환값은 1(성공) 또는 0(실패)입니다. 실패 상세는 C:\temp\TlsMailSender.log 에 남습니다.
    /// </summary>
    public static unsafe class NativeExports
    {
        /// <summary>
        /// STARTTLS SMTP로 메일을 발송합니다.
        /// </summary>
        /// <param name="attachments">
        /// UTF-16 첨부 경로 목록. null 또는 빈 문자열이면 첨부가 없습니다.
        /// 여러 파일은 '|' 로 구분합니다.
        /// </param>
        /// <returns>성공 시 1, 실패 시 0.</returns>
        [UnmanagedCallersOnly(EntryPoint = "SendMail", CallConvs = new[] { typeof(CallConvStdcall) })]
        public static int SendMail(
            char* smtpServer,
            int smtpPort,
            char* smtpUser,
            char* smtpPass,
            int useTls,
            char* from,
            char* to,
            char* subject,
            char* body,
            char* attachments)
        {
            try
            {
                var sender = new MailSender();
                bool ok = sender.SendMail(
                    FromUtf16(smtpServer),
                    smtpPort,
                    FromUtf16(smtpUser),
                    FromUtf16(smtpPass),
                    useTls != 0,
                    FromUtf16(from),
                    FromUtf16(to),
                    FromUtf16(subject),
                    FromUtf16(body),
                    SplitAttachments(attachments));
                return ok ? 1 : 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 표시 이름과 Reply-To를 포함해 메일을 발송합니다.
        /// 선택 인자(fromDisplayName, toDisplayName, replyTo)는 null 포인터를 허용합니다.
        /// SendMailWithReplyTo / SendMailWithAlias 는 이 함수에 null 을 넘겨 호출합니다.
        /// </summary>
        /// <returns>성공 시 1, 실패 시 0.</returns>
        [UnmanagedCallersOnly(EntryPoint = "SendMailWithAliasReplyTo", CallConvs = new[] { typeof(CallConvStdcall) })]
        public static int SendMailWithAliasReplyTo(
            char* smtpServer,
            int smtpPort,
            char* smtpUser,
            char* smtpPass,
            int useTls,
            char* from,
            char* to,
            char* subject,
            char* body,
            char* attachments,
            char* fromDisplayName,
            char* toDisplayName,
            char* replyTo)
        {
            try
            {
                var sender = new MailSender();
                bool ok = sender.SendMailWithAliasReplyTo(
                    FromUtf16(smtpServer),
                    smtpPort,
                    FromUtf16(smtpUser),
                    FromUtf16(smtpPass),
                    useTls != 0,
                    FromUtf16(from),
                    FromUtf16(to),
                    FromUtf16(subject),
                    FromUtf16(body),
                    SplitAttachments(attachments),
                    FromUtf16(fromDisplayName),
                    FromUtf16(toDisplayName),
                    FromUtf16(replyTo));
                return ok ? 1 : 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// AllowedCerts.txt 를 다시 읽습니다.
        /// </summary>
        [UnmanagedCallersOnly(EntryPoint = "ReloadAllowedCerts", CallConvs = new[] { typeof(CallConvStdcall) })]
        public static void ReloadAllowedCerts()
        {
            try
            {
                new MailSender().ReloadAllowedCerts();
            }
            catch
            {
                // LoadAllowedThumbprints 가 실패 원인을 로그에 남깁니다.
            }
        }

        private static string FromUtf16(char* value)
        {
            if (value == null)
                return null;
            return new string(value);
        }

        private static string[] SplitAttachments(char* attachments)
        {
            string value = FromUtf16(attachments);
            if (string.IsNullOrWhiteSpace(value))
                return null;

            string[] parts = value.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
                return null;

            for (int i = 0; i < parts.Length; i++)
                parts[i] = parts[i].Trim();

            return parts;
        }
    }
}
#endif
