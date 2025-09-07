Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                        (ByVal pCaller As Long, _
                         ByVal szURL As String, _
                         ByVal szFileName As String, _
                         ByVal dwReserved As Long, _
                         ByVal lpfnCB As Long) As Long
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" _
                        (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                        (ByVal pCaller As Long, _
                         ByVal szURL As String, _
                         ByVal szFileName As String, _
                         ByVal dwReserved As Long, _
                         ByVal lpfnCB As Long) As Long
Private Declare Function SHCreateDirectoryEx Lib "shell32.dll" Alias "SHCreateDirectoryExA" _
                        (ByVal hwnd As LongPtr, ByVal pszPath As String, ByVal psa As LongPtr) As Long
#End If

Private Type VersionType '---本当はクラスオブジェクトにしたいけどこれだけのためにモジュール作りたくない
    Major As Long
    Minor As Long
    Build As Long
    Revision As Long
    BuildVersion As String
    RevisionVersion As String
End Type
Const ZIP_FILE As String = "chromedriver.zip"
Const DRIVER_EXE As String = "chromedriver.exe"
Dim workPath As String
Public Function ChromeDriverAutoUpdate(Optional ByVal ForcedExecution As Boolean = False) As Boolean
'====================================================================================================
'chrome.exeとchromedriver.exeのバージョンを比較してchromedriverを自動更新する
'もしくは強制実行フラグ（ForcedExecution）がTrueでも実行する
'====================================================================================================
    Dim chromePath As String '---chrome.exeが保存されているパス
    Dim chromeFullpath As String '---chrome.exeまで含めたフルパス
    Dim chromeVersion As VersionType
    Dim chromedriverPath As String
    Dim chromedriverFullPath As String
    Dim objFso As New Scripting.FileSystemObject
    ' ---chromedriverをダウンロード用のフォルダを作成する　※Seleniumのキャッシュ構造に合わせている
    workPath = Environ("USERPROFILE") & "\.cache\selenium\seleniumbasic"
    Select Case SHCreateDirectoryEx(0&, workPath, 0&)
        Case 0:
            ' ---作成成功
        Case 183
            ' ---作成済み
        Case Else:
            ' ---作成できなかった時
            MsgBox "ダウンロード用フォルダを作成できませんでした" & vbCrLf & Error(Err), vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---chrome本体のフォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramW6432") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\Google\Chrome\Application")
            chromePath = Environ("ProgramFiles") & "\Google\Chrome\Application"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\Google\Chrome\Application")
            chromePath = Environ("LOCALAPPDATA") & "\Google\Chrome\Application"
        Case Else
            MsgBox "'chrome'フォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---念のためchrome.exeを確認する
    If objFso.FileExists(chromePath & "\chrome.exe") = True Then
        chromeFullpath = chromePath & "\chrome.exe"
    Else
        MsgBox "'chrome.exe'が見つかりません", vbCritical
        Exit Function
    End If
    
    '---SeleniumBasicのフォルダを探す
    Select Case True
        Case objFso.FolderExists(Environ("ProgramW6432") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramW6432") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("ProgramFiles") & "\SeleniumBasic")
            chromedriverPath = Environ("ProgramFiles") & "\SeleniumBasic"
        Case objFso.FolderExists(Environ("LOCALAPPDATA") & "\SeleniumBasic")
            chromedriverPath = Environ("LOCALAPPDATA") & "\SeleniumBasic"
        Case Else
            MsgBox "'SeleniumBasic'のフォルダが見つかりません", vbCritical
            ChromeDriverAutoUpdate = False
            Exit Function
    End Select
    
    '---念のためchromedriver.exeを確認する
    If objFso.FileExists(chromedriverPath & "\" & DRIVER_EXE) = True Then
        chromedriverFullPath = chromedriverPath & "\" & DRIVER_EXE
    Else
        MsgBox "'chromedriver.exe'が見つかりません", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
        
    '---chrome.exeのバージョンを取得する
    If GetChromeVersion(chromeFullpath, chromeVersion) = False Then '---chrome.exeのバージョンを取得する
        MsgBox "'chrome.exe'のバージョンが取得できませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---chrome.exeのバージョンに合わせたchromedriver.exeをダウンロードする
    If ChromedriverQuickCheck(chromedriverPath, chromeVersion) = False Then '---chromedriverのバージョンを取得する
        MsgBox "'chromedriver.exe'のバージョンが取得できませんでした", vbCritical
        ChromeDriverAutoUpdate = False
        Exit Function
    End If
    
    '---結果として更新していない場合もあるが、更新失敗じゃなくて更新不要な判定だからTrueを返す
    ChromeDriverAutoUpdate = True
Exit Function
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chromedriver の入替に失敗しました" & vbCrLf & Error(Err) & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
    ChromeDriverAutoUpdate = False
End Function
Private Function GetChromeVersion(ByVal chromeFullpath As String, ByRef chromeVersion As VersionType) As Boolean
'====================================================================================================
'PowerShellでchrome.exeのバージョン情報を取得する　※一瞬PowerShellが立ち上がる
'====================================================================================================
    Dim command As String
    Dim objRet As Object
    
    On Error GoTo ErrLabel
        '---chromeバージョン情報の初期値
        chromeVersion.Major = 1
        chromeVersion.Minor = 0
        chromeVersion.Build = 0
        chromeVersion.Revision = 0
        '---chrome.exeのバージョンを取得するPowerShellコマンド
        command = "powershell.exe -NoProfile -ExecutionPolicy Bypass (Get-Item -Path '" & chromeFullpath & "').VersionInfo.FileVersion"
        '---PowerShellの実行結果をセット
        Set objRet = CreateObject("WScript.Shell").Exec(command)
        '---PowerShellのコマンドレットの実行結果を取得
        chromeVersion.RevisionVersion = Trim(objRet.StdOut.ReadAll)
        '---情報の取得が終わったらオブジェクトをクリアする
        Set objRet = Nothing
        '---改行コードが含まれているから削除する
        chromeVersion.RevisionVersion = Trim(Replace(Replace(Replace(chromeVersion.RevisionVersion, vbCrLf, vbNullString), vbCr, vbNullString), vbLf, vbNullString))
        '---バージョン情報を分けて返す
        With CreateObject("VBScript.RegExp") '---正規表現の準備
            .Pattern = "\d+\.\d+\.\d+(\.\d+)?"
            .Global = True
            If .test(chromeVersion.RevisionVersion) Then '---念のため正規表現でバージョン情報をチェックする
                chromeVersion.Major = CLng(Split(chromeVersion.RevisionVersion, ".")(0))
                chromeVersion.Minor = CLng(Split(chromeVersion.RevisionVersion, ".")(1))
                chromeVersion.Build = CLng(Split(chromeVersion.RevisionVersion, ".")(2))
                If UBound(Split(chromeVersion.RevisionVersion, ".")) >= 3 Then '---リビジョン番号がなければ9999を仮でセット※基本あるはず
                    chromeVersion.Revision = CLng(Split(chromeVersion.RevisionVersion, ".")(3))
                Else
                    chromeVersion.Revision = 9999
                End If
                chromeVersion.BuildVersion = Join(Array(chromeVersion.Major, chromeVersion.Minor, chromeVersion.Build), ".") '---リビジョンを覗いたショートバージョン情報をセットする
                Debug.Print "Chromeのバージョン：" & chromeVersion.RevisionVersion
            Else '---正規表現不一致なら失敗で返す
                MsgBox "chrome.exe のバージョン情報取得に失敗しました" & vbCrLf & "[取得バージョン情報：" & chromeVersion.RevisionVersion & "]" & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
                GetChromeVersion = False
                Exit Function
            End If
        End With
        GetChromeVersion = True
    On Error GoTo 0
    Exit Function
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chrome.exe のバージョン情報取得に失敗しました" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
    GetChromeVersion = False
End Function
Private Function ChromedriverQuickCheck(chromedriverPath, chromeVersion As VersionType) As Boolean
    Dim objHttp As New MSXML2.XMLHTTP60
    Dim targetVarsion As String
    Dim uri As String
    Dim api_endpoints As String
    Dim downloadPath As String
    Dim objFso As New Scripting.FileSystemObject
    Const TARGET_PLATFORM As String = "win64"

    api_endpoints = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" & chromeVersion.BuildVersion
    On Error GoTo ErrLabel
        With objHttp
            .Open "GET", api_endpoints, False
            .Send
            targetVarsion = .responseText '---JSON endpoints から情報を収集する
            downloadPath = workPath & "\" & targetVarsion
            '--- 念のためリビジョンバージョンを比較する
            If chromeVersion.Revision >= CLng(Split(targetVarsion, ".")(3)) Then
                '--- まだダウンロードしていなかったらダウンロードする
                If objFso.FileExists(downloadPath & "\" & DRIVER_EXE) = False Then
                    uri = "https://storage.googleapis.com/chrome-for-testing-public/" & targetVarsion & "/" & TARGET_PLATFORM & "/chromedriver-" & TARGET_PLATFORM & ".zip"
                    Call DownloadChromedriver(uri, targetVarsion)
                    Call objFso.DeleteFile(chromedriverPath & "\" & DRIVER_EXE, True)
                    Call objFso.GetFile(downloadPath & "\" & DRIVER_EXE).Copy(chromedriverPath & "\" & DRIVER_EXE, True)
                    Debug.Print "インストールしたChromedriverのバージョン：" & targetVarsion
                Else
                    Debug.Print "Chromedriverは最新です"
                End If
            Else
                Debug.Print "Chromeのバージョンが古いためChromedriverは更新しません"
            End If
        End With
    On Error GoTo 0
    ChromedriverQuickCheck = True
    Exit Function
ErrLabel:     '---予期せぬエラーの分岐
    MsgBox "chromedriver.exe の更新に失敗しました" & vbCrLf & "[" & Error(Err) & "]" & vbCrLf & "※この画面のキャプチャを作成者へ送ってください"
    ChromedriverQuickCheck = False
End Function
Private Function DownloadChromedriver(ByVal url As String, targetVersion As String) As Boolean
    Dim rc As Long
    Dim downloadPath As String
    Dim newDriverPath As String
    Dim objFso As New Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    downloadPath = workPath & "\" & targetVersion
    ' ---chromedriverのフォルダを作成する
    Select Case SHCreateDirectoryEx(0&, downloadPath, 0&)
        Case 0:
            ' ---作成成功
        Case 183
            ' ---作成済み
        Case Else:
            ' ---作成できなかった時
            MsgBox "ChromeDriver用フォルダを作成できませんでした" & vbCrLf & Error(Err), vbCritical
            DownloadChromedriver = False
            Exit Function
    End Select
    
    '---ファイルをダウンロードする
    If URLDownloadToFile(0, url, workPath & "\" & ZIP_FILE, 0, 0) <> 0 Then
        MsgBox "ChromeDriverをダウンロードできませんでした" & vbCrLf & Error(Err), vbCritical
        DownloadChromedriver = False
        Exit Function
    End If
    Application.DisplayAlerts = False
    '---zipを既定のフォルダに向けて解凍する
    With CreateObject("Shell.Application") '---zipを既定のフォルダに向けて解凍する
        .Namespace((downloadPath)).CopyHere .Namespace((workPath & "\" & ZIP_FILE)).Items
    End With
    '--- 解凍したフォルダから再起処理してchromedriver.exeのフルパスを取得する
    newDriverPath = SearchFilesRecursively(downloadPath & "\", "chromedriver.exe")
    If newDriverPath = "" Then
        MsgBox "chromedriver.exe の更新に失敗しました"
        DownloadChromedriver = False
    End If
    '---chromedriverをバージョンフォルダ直下に移動する
    Call objFso.MoveFile(newDriverPath, downloadPath & "\")
    '---chromedriverがなくなった不要フォルダを削除する
    For Each objFolder In objFso.GetFolder(downloadPath).SubFolders
        objFolder.Delete True
    Next
    '---zipファイルを削除する
    Call objFso.DeleteFile(workPath & "\" & ZIP_FILE, True)
    Application.DisplayAlerts = True
    DownloadChromedriver = True
End Function
Function SearchFilesRecursively(ByVal folderPath As String, fileName As String) As String
    '====================================================================================================
    ' folderPathを起点に再帰処理でサブフォルダまで対象にしてfileNameを探してフルパスを返す
    '====================================================================================================
    Dim objFso As New Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim objFile As Scripting.File
    Dim result As String

    ' ファイル一覧をチェック
    For Each objFile In objFso.GetFolder(folderPath).Files
        If objFile.Name = fileName Then
            SearchFilesRecursively = objFile.Path
            Exit Function
        End If
    Next objFile

    ' サブフォルダを再帰的に探索
    For Each subFolder In objFso.GetFolder(folderPath).SubFolders
        result = SearchFilesRecursively(subFolder.Path, fileName)
        If result <> "" Then
            SearchFilesRecursively = result
            Exit Function
        End If
    Next subFolder

    ' 見つからなかった場合
    SearchFilesRecursively = ""
End Function

