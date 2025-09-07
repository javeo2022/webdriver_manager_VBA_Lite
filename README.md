# webdriver_manager_VBA_Lite

このリポジトリは **VBA (Visual Basic for Applications)** で書かれた  
**Google Chrome のバージョンに合わせて Chromedriver を自動更新するスクリプト** です。  
主に **SeleniumBasic** を利用する環境で、Chromedriver のバージョン不一致エラーを防ぐために作成されています。

---

## 特徴

- Chrome のインストールパスを自動検出
- Chrome のバージョンを PowerShell 経由で取得
- Google 提供の API から該当する Chromedriver を自動ダウンロード
- ZIP 解凍 → `chromedriver.exe` 更新まで自動で実行
- SeleniumBasic の標準フォルダ構成に対応
- chromeが自動更新されている前提です ※意図的にダウングレードされない前提

---

## 動作環境

- Windows 10 / 11
- Microsoft Office (VBAが実行可能な環境)
- SeleniumBasic がインストール済みであること
- PowerShell 利用可能な環境

---

## インストール方法

1. `モジュール` にこのコードを貼り付けます。
2. `ツール > 参照設定` から以下を参照設定してください:
   - Microsoft Scripting Runtime  
   - Microsoft XML, v6.0
3. 任意のプロシージャから以下を実行します:

```vba
Sub TestUpdate()
    If ChromeDriverAutoUpdate Then
        Debug.Print "Chromedriverの更新成功"
    Else
        Debug.Print "Chromedriverの更新失敗"
    End If
End Sub
