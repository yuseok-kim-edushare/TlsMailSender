# TlsMailSender C++ 샘플

`TlsMailSender.dll`(net10.0-windows, Native AOT `NativeLib=Shared` 게시물)을
`LoadLibraryW`/`GetProcAddress`로 동적 로드해 `SendMail`/`SendMailWithAliasReplyTo`/
`ReloadAllowedCerts` 내보내기 함수를 `__stdcall`로 호출하는 최소 예제입니다.

문자열은 UTF-16(`wchar_t*`)입니다. 반환값은 1(성공) 또는 0(실패)입니다.
`SendMailWithReplyTo`와 `SendMailWithAlias`는 `SendMailWithAliasReplyTo`에
해당 인자를 null로 넘겨 호출합니다.

## 사전 준비: 네이티브 라이브러리 게시

```powershell
dotnet publish ..\..\TlsMailSender\TlsMailSender.csproj -f net10.0-windows -r win-x64 -c Release
```

게시 결과(`TlsMailSender\bin\Release\net10.0-windows\win-x64\publish\TlsMailSender.dll`)를
이 샘플의 실행 파일과 같은 폴더로 복사해야 합니다. 32비트 호스트는 `-r win-x86` 게시물을 사용합니다.

사설 인증서를 허용하려면 `AllowedCerts.txt`를 네이티브 DLL과 같은 폴더(프로세스의 `AppContext.BaseDirectory`)에 둡니다.

## 빌드 (Visual Studio Developer Command Prompt + CMake)

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
```

또는 MSVC가 설치된 환경에서 `cl.exe`로 직접 빌드할 수도 있습니다.

```powershell
cl /EHsc /std:c++17 /utf-8 /DUNICODE /D_UNICODE main.cpp /Fe:TlsMailSenderCppSample.exe
```

## 실행

```powershell
copy ..\..\TlsMailSender\bin\Release\net10.0-windows\win-x64\publish\TlsMailSender.dll build\Release\
build\Release\TlsMailSenderCppSample.exe smtp.example.com 25 user@example.com secret 1 user@example.com dest@example.com "제목" "본문" "C:\temp\a.txt|C:\temp\b.txt"
```

첨부 경로는 `|`로 구분합니다. 표시 이름이나 Reply-To를 넘기면 `SendMailWithAliasReplyTo`를 호출합니다.

```powershell
build\Release\TlsMailSenderCppSample.exe smtp.example.com 25 user@example.com secret 1 user@example.com dest@example.com "제목" "본문" "" "보낸사람" "받는사람" "reply@example.com"
```
