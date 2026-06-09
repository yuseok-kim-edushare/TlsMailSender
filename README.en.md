# SimpleNetMail: TLS (STARTTLS) Email Library for PowerBuilder (.NET DLL Import & COM)

TLS (STARTTLS) email library for PowerBuilder via .NET Assembly Import (2019 R3+) or COM interop.

*Read this in other languages: [한국어](README.md)*

[![CI Build](https://github.com/yuseok-kim-edushare/TlsMailSender/actions/workflows/ci.yaml/badge.svg)](https://github.com/yuseok-kim-edushare/TlsMailSender/actions/workflows/ci.yaml)

---

## Overview

This project is a .NET Framework Class Library that helps PowerBuilder applications send email over TLS (STARTTLS) encrypted connections using either .NET Assembly Import (2019 R3+) or COM. It supports attachments and sender/recipient display names (aliases), and includes a **certificate whitelist (`AllowedCerts.txt`)** so you can safely send mail in environments that use private certificates.

## Key Features

*   **Broad PowerBuilder support**:
    *   **.NET Assembly Import**: Works in PowerBuilder 2019 R3 and later without COM registration
    *   **COM**: Works in all PowerBuilder versions after COM registration via OLEObject
*   **Platform independent**: Built as MSIL (AnyCPU) for both 32-bit and 64-bit environments
*   **TLS (STARTTLS) encryption**: Uses STARTTLS for secure SMTP communication (port 25 supported)
*   **Attachment support**: Easily attach multiple files to outgoing messages
*   **Flexible certificate validation**: Performs system certificate validation by default; private certificates can be allowed by registering their thumbprints in `AllowedCerts.txt` (not a blanket bypass of validation)
*   **Sender/recipient display names**: Specify display names (aliases) alongside email addresses
*   **.NET Framework 4.8.1 based**: Targets the .NET Framework version supported by PowerBuilder 2019 R3
*   **Detailed logging**: Records send success/failure and certificate validation details to `C:\temp\TlsMailSender.log`

## Development Environment

*   Visual Studio 2022 (or any environment that supports .NET Framework 4.8 development)
*   .NET Framework 4.8 SDK
*   PowerBuilder 2019 R3 or later

## Installation and Usage

This library supports both **.NET Assembly Import** and **COM** approaches. Choose the method that fits your PowerBuilder version.

### 1. Build the C# Project

1.  Open Visual Studio 2022.
2.  Create a new project using the "Class Library (.NET Framework)" template (e.g., `TlsMailSender`).
3.  Set the **target framework** to `.NET Framework 4.8`.
4.  In Project Properties, go to the **Build** tab.
5.  Keep **Platform target** at the default **AnyCPU (MSIL)** (supports both 32-bit and 64-bit).
6.  Add the `MailSender.cs` file to the project and implement the code.
7.  Build the solution to produce `TlsMailSender.dll`.
8.  **(Optional)** If using private certificates, create an `AllowedCerts.txt` file in the same directory as the DLL and add the allowed certificate thumbprints (SHA1/SHA256).

### 2. Using in PowerBuilder

#### Method A: .NET Assembly Import (Recommended for PowerBuilder 2019 R3+)

1.  Open your project in PowerBuilder 2019 R3 or later.
2.  Select `Project` -> `Import .NET Assembly...` from the menu.
3.  In the `.NET DLL Import` dialog, select the built DLL and choose `.NET Framework` as the Framework Type.
4.  Set the Destination PBT/PBL.
5.  Select the `SimpleNetMail.MailSender` class and import it.
6.  Confirm the generated proxy object in the System Tree (e.g., `n_simplenetmail_mailsender`).

#### Method B: COM (All PowerBuilder Versions)

1.  **COM registration** (requires administrator privileges):
    ```powershell
    # For 32-bit PowerBuilder
    C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe TlsMailSender.dll /codebase /tlb:TlsMailSender.tlb
    
    # For 64-bit PowerBuilder
    C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe TlsMailSender.dll /codebase /tlb:TlsMailSender.tlb
    ```
    
    Alternatively, to register automatically on build in Visual Studio, check **"Register for COM interop"** on the **Build** tab in Project Properties (requires running Visual Studio as administrator).

2.  Create the COM object in PowerBuilder using OLEObject.

### 3. Calling MailSender from PowerBuilder Scripts

#### 3.1. .NET Assembly Import — `SendMail` Example

```powerscript
// =================================================================================
// 1) Declare DotNetObject variable
// =================================================================================
DotNetObject    dn_mailer
boolean         lb_result

// SMTP settings (example: port 25 + STARTTLS)
string          ls_smtpServer = "smtp.your-server.com"
integer         li_smtpPort   = 25
string          ls_smtpUser   = "your_email@your-server.com"
string          ls_smtpPass   = "your_password"
boolean         lb_useTLS     = TRUE               // TRUE: attempt STARTTLS

// Mail content
string          ls_from       = "your_email@your-server.com"
string          ls_to         = "recipient1@domain.com;recipient2@domain.com"
string          ls_subject    = "PowerBuilder .NET Import Test Mail"
string          ls_body       = "This mail was sent via PowerBuilder .NET Import."

// Attachments
string[]        lsa_attachments
lsa_attachments = CREATE string[1]
lsa_attachments[1] = "C:\Path\To\Your\file.txt"

// =================================================================================
// 2) Create DotNetObject instance
// =================================================================================
// Proxy name may vary depending on import results.
string ls_proxy_name = "n_simplenetmail_mailsender" 

dn_mailer = CREATE DotNetObject(ls_proxy_name)
IF IsNull(dn_mailer) THEN
    MessageBox("Error", ".NET object creation failed: " + ls_proxy_name)
    RETURN
END IF

// =================================================================================
// 3) Call SendMail
// =================================================================================
// Certificate validation: system validation first; on failure, thumbprints in AllowedCerts.txt are checked
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
    MessageBox("Success", "Mail sent successfully.")
ELSE
    MessageBox("Failure", "Mail send failed. Check C:\temp\TlsMailSender.log for details.")
END IF

dn_mailer = NULL
```

#### 3.2. COM — `SendMail` Example

```powerscript
// =================================================================================
// 1) Declare OLEObject variable
// =================================================================================
OLEObject    ole_mailer
boolean      lb_result

// SMTP settings (example: port 25 + STARTTLS)
string       ls_smtpServer = "smtp.your-server.com"
integer      li_smtpPort   = 25
string       ls_smtpUser   = "your_email@your-server.com"
string       ls_smtpPass   = "your_password"
boolean      lb_useTLS     = TRUE               // TRUE: attempt STARTTLS

// Mail content
string       ls_from       = "your_email@your-server.com"
string       ls_to         = "recipient1@domain.com;recipient2@domain.com"
string       ls_subject    = "PowerBuilder COM Test Mail"
string       ls_body       = "This mail was sent via PowerBuilder COM."

// Attachments
string[]     lsa_attachments
lsa_attachments = CREATE string[1]
lsa_attachments[1] = "C:\Path\To\Your\file.txt"

// =================================================================================
// 2) Create and connect COM object
// =================================================================================
ole_mailer = CREATE OLEObject

IF ole_mailer.ConnectToNewObject("SimpleNetMail.MailSender") <> 0 THEN
    MessageBox("Error", "COM object connection failed")
    RETURN
END IF

// =================================================================================
// 3) Call SendMail
// =================================================================================
// Certificate validation: system validation first; on failure, thumbprints in AllowedCerts.txt are checked
lb_result = ole_mailer.SendMail( &
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
    MessageBox("Success", "Mail sent successfully.")
ELSE
    MessageBox("Failure", "Mail send failed. Check C:\temp\TlsMailSender.log for details.")
END IF

// Disconnect
ole_mailer.DisconnectObject()
DESTROY ole_mailer
```

#### 3.3. `SendMailWithAlias` Example (with Display Names)

```powerscript
// ... (variable declarations and settings same as above) ...

// Display names
string          ls_from_alias = "Administrator"
string          ls_to_alias   = "Recipient1;Recipient2" // semicolon-separated for multiple recipients

// ... (object creation) ...

// Method call
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

## Important Notes and Considerations

*   **Certificate validation and AllowedCerts.txt**:
    *   This library **performs system certificate validation by default** (invalid certificates are blocked).
    *   If you must use a private or expired certificate, register its **thumbprint (SHA1 or SHA256)** in the `AllowedCerts.txt` file located in the same directory as the DLL.
    *   Sending fails if `AllowedCerts.txt` is missing or the thumbprint does not match.
    *   Failed certificate thumbprints are logged to `C:\temp\TlsMailSender.log` — copy them into `AllowedCerts.txt` as needed.
*   **Log file**: Errors and certificate validation details are written to `C:\temp\TlsMailSender.log`. Check this log first when troubleshooting.
*   **Platform compatibility**:
    *   The DLL is built as MSIL (AnyCPU) and loads appropriately in both 32-bit and 64-bit environments.
    *   For COM, register with the `regasm.exe` that matches your PowerBuilder bitness (32-bit PB → Framework folder, 64-bit PB → Framework64 folder).
*   **SMTP port**: When using port 25, the server must support STARTTLS.
*   **COM registration**: COM registration requires administrator privileges. Unregister with `regasm.exe /unregister`.

## License

This project is distributed under the MIT License. See `LICENSE.txt` for details.
