# SimpleNetMail: PowerBuilder 2019 R3+용 TLS(STARTTLS) 메일 발송 라이브러리 (.NET DLL Import 활용)

## 프로젝트 소개

이 프로젝트는 PowerBuilder 2019 R3 이상 버전에서 .NET Assembly Import 기능을 사용하여 간편하게 TLS (STARTTLS) 암호화 연결로 이메일을 발송할 수 있도록 돕는 .NET Framework Class Library 입니다. 
첨부파일 지원 및 발신자/수신자 표시 이름(Alias) 기능을 제공하며, **인증서 화이트리스트(`AllowedCerts.txt`)** 기능을 통해 사설 인증서 환경에서도 안전하게 메일을 발송할 수 있도록 설계되었습니다.

## 주요 특징

*   **PowerBuilder 2019 R3+ 지원**: PowerBuilder의 내장된 .NET Assembly Import 기능을 통해 별도의 COM 등록 없이 바로 사용 가능합니다.
*   **TLS (STARTTLS) 암호화**: SMTP 통신 시 STARTTLS를 사용하여 안전하게 데이터를 전송합니다 (포트 25 지원).
*   **첨부파일 지원**: 여러 개의 첨부파일을 간편하게 추가하여 발송할 수 있습니다.
*   **유연한 인증서 검증**: 기본적으로 시스템 인증서 검증을 수행하며, 사설 인증서를 사용하는 경우 `AllowedCerts.txt` 파일에 지문(Thumbprint)을 등록하여 예외적으로 허용할 수 있습니다. (무조건적인 검증 생략이 아님)
*   **발신/수신자 표시 이름 지원**: 메일 주소와 함께 표시 이름(Alias)을 지정할 수 있습니다.
*   **.NET Framework 4.8 기반**: PowerBuilder 2019 R3가 지원하는 .NET Framework 버전으로 개발되었습니다.
*   **상세 로깅**: `C:\temp\TlsMailSender.log` 파일에 발송 성공/실패 및 인증서 검증 상세 내역을 기록합니다.

## 개발 환경

*   Visual Studio 2022 (또는 .NET Framework 4.8 개발이 가능한 환경)
*   .NET Framework 4.8 SDK
*   PowerBuilder 2019 R3 이상 버전

## 설치 및 사용 방법

이 라이브러리는 COM 방식이 아닌, PowerBuilder 2019 R3에 내장된 **.NET Assembly Import** 기능을 사용하여 등록하고 호출합니다.

### 1. C# 프로젝트 빌드

1.  Visual Studio 2022를 실행합니다.
2.  "Class Library (.NET Framework)" 템플릿으로 새 프로젝트를 생성합니다. (예: `SimpleNetMail`)
3.  **타겟 프레임워크**를 `.NET Framework 4.8`로 설정합니다.
4.  프로젝트 속성(Project Properties)에서 **Build** 탭으로 이동합니다.
5.  **Platform target**을 `x86`으로 설정합니다. (대부분의 PowerBuilder 2019 R3 환경이 32비트입니다.)
6.  **"Register for COM interop" 옵션은 체크 해제**합니다.
7.  `MailSender.cs` 파일을 프로젝트에 추가하고 코드를 작성합니다.
8.  `AssemblyInfo.cs` 파일에서 COM 관련 특성(`ComVisible`, `Guid` 등)을 제거합니다.
9.  솔루션을 빌드하여 `SimpleNetMail.dll` (또는 `TlsMailSender.dll`)을 생성합니다.
10. **(선택 사항)** 사설 인증서를 사용하는 경우, DLL과 동일한 경로에 `AllowedCerts.txt` 파일을 생성하고 허용할 인증서의 지문(SHA1/SHA256)을 입력합니다.

### 2. PowerBuilder 2019 R3+에서 .NET DLL Import

1.  PowerBuilder 2019 R3 이상 IDE를 실행하고 프로젝트를 엽니다.
2.  메뉴에서 `Project` -> `Import .NET Assembly...`를 선택합니다.
3.  `.NET DLL Import` 대화상자에서 빌드한 DLL 파일을 선택하고, Framework Type은 `.NET Framework`를 선택합니다.
4.  Destination PBT/PBL을 설정합니다.
5.  `SimpleNetMail.MailSender` 클래스를 선택하고 Import 합니다.
6.  System Tree에서 생성된 Proxy 오브젝트(예: `n_simplenetmail_mailsender`)를 확인합니다.

### 3. PowerBuilder 스크립트에서 MailSender 호출

#### 3.1. `SendMail` 메서드 사용 예제

```powerscript
// =================================================================================
// 1) DotNetObject 변수 선언
// =================================================================================
DotNetObject    dn_mailer
boolean         lb_result

// SMTP 설정 (예시: 포트 25 + STARTTLS)
string          ls_smtpServer = "smtp.your-server.com"
integer         li_smtpPort   = 25
string          ls_smtpUser   = "your_email@your-server.com"
string          ls_smtpPass   = "your_password"
boolean         lb_useTLS     = TRUE               // TRUE: STARTTLS 시도

// 메일 정보
string          ls_from       = "your_email@your-server.com"
string          ls_to         = "recipient1@domain.com;recipient2@domain.com"
string          ls_subject    = "PowerBuilder .NET Import 테스트 메일"
string          ls_body       = "이 메일은 PowerBuilder .NET Import로 발송되었습니다."

// 첨부파일
string[]        lsa_attachments
lsa_attachments = CREATE string[1]
lsa_attachments[1] = "C:\Path\To\Your\file.txt"

// =================================================================================
// 2) DotNetObject 인스턴스 생성
// =================================================================================
// Proxy 이름은 Import 결과에 따라 다를 수 있습니다.
string ls_proxy_name = "n_simplenetmail_mailsender" 

dn_mailer = CREATE DotNetObject(ls_proxy_name)
IF IsNull(dn_mailer) THEN
    MessageBox("오류", ".NET 객체 생성 실패: " + ls_proxy_name)
    RETURN
END IF

// =================================================================================
// 3) SendMail 메서드 호출
// =================================================================================
// 인증서 검증: 시스템 검증을 수행하며, 실패 시 AllowedCerts.txt의 지문과 대조하여 허용 여부 결정
lb_result = dn_mailer.SendMail( &
                ls_smtpServer, &
                li_smtpPort, &
                ls_smtpUser, &
                ls_smtpPass, &
                lb_useTLS, &
                ls_from, &
                ls_to, &
                ls_subject, &
                ls_body, &
                lsa_attachments )

IF lb_result = TRUE THEN
    MessageBox("성공", "메일이 성공적으로 전송되었습니다.")
ELSE
    MessageBox("실패", "메일 전송 실패. C:\temp\TlsMailSender.log 로그를 확인하세요.")
END IF

dn_mailer = NULL
```

#### 3.2. `SendMailWithAlias` 메서드 사용 예제 (표시 이름 포함)

```powerscript
// ... (변수 선언 및 설정은 위와 동일) ...

// 표시 이름 설정
string          ls_from_alias = "관리자"
string          ls_to_alias   = "수신자1;수신자2" // 수신자가 여러 명일 경우 세미콜론으로 구분

// ... (객체 생성) ...

// 메서드 호출
lb_result = dn_mailer.SendMailWithAlias( &
                ls_smtpServer, &
                li_smtpPort, &
                ls_smtpUser, &
                ls_smtpPass, &
                lb_useTLS, &
                ls_from, &
                ls_to, &
                ls_subject, &
                ls_body, &
                lsa_attachments, &
                ls_from_alias, &
                ls_to_alias )
```

## 주요 고려사항 및 주의사항

*   **인증서 검증 및 AllowedCerts.txt**: 
    *   이 라이브러리는 **기본적으로 시스템 인증서 검증**을 수행합니다 (유효하지 않은 인증서는 차단됨).
    *   사설 인증서나 만료된 인증서를 사용해야 하는 경우, 해당 인증서의 **지문(Thumbprint, SHA1 또는 SHA256)**을 확인하여 DLL 파일과 동일한 경로에 있는 `AllowedCerts.txt` 파일에 등록해야 합니다.
    *   `AllowedCerts.txt` 파일이 없거나 지문이 일치하지 않으면 전송이 실패합니다.
    *   로그 파일(`C:\temp\TlsMailSender.log`)에 검증 실패한 인증서의 지문이 기록되므로, 이를 복사하여 `AllowedCerts.txt`에 추가하면 됩니다.
*   **로그 파일**: 동작 중 발생하는 오류와 인증서 검증 정보는 `C:\temp\TlsMailSender.log` 파일에 기록됩니다. 문제 해결 시 이 로그를 먼저 확인하십시오.
*   **플랫폼(x86/x64)**: PowerBuilder IDE 및 실행 파일이 32비트인 경우, DLL도 `x86`으로 빌드되어야 합니다.
*   **SMTP 포트**: 포트 25 사용 시 서버가 STARTTLS를 지원해야 합니다.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE.txt` 파일을 참조하십시오.
