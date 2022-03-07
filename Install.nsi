; インストーラーの識別子
!define PRODUCT_NAME "Ladex"
; インストーラーのバージョン。
!define PRODUCT_VERSION "1.2.1.0"

; 多言語で使用する場合はここをUnicodeにすることを推奨
Unicode true

; インストーラーのアイコン
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\win-install.ico"

; アンインストーラーのアイコン
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\win-uninstall.ico"

; インストーラの見た目
; !define MUI_HEADERIMAGE
; !define MUI_HEADERIMAGE_RIGHT
; !define MUI_HEADERIMAGE_BITMAP          "${NSISDIR}\Contrib\Graphics\Header\win.bmp"
; !define MUI_HEADERIMAGE_UNBITMAP        "${NSISDIR}\Contrib\Graphics\Header\win.bmp"

; !define MUI_WELCOMEFINISHPAGE_BITMAP    "${NSISDIR}\Contrib\Graphics\Wizard\win.bmp"
; !define MUI_UNWELCOMEFINISHPAGE_BITMAP  "${NSISDIR}\Contrib\Graphics\Wizard\win.bmp"


; 使用する外部ライブラリ
!include Sections.nsh
!include MUI2.nsh
!include LogicLib.nsh
; !include nsProcess.nsh


; 圧縮設定。通常は/solid lzmaが最も圧縮率が高い
SetCompressor /solid lzma

; インストーラー名
Name "Ladex [Addin For Excel Library]"
; 出力されるファイル名
OutFile "${PRODUCT_NAME}-${PRODUCT_VERSION}-setup.exe"

; インストール/アンインストール時の進捗画面
ShowInstDetails   show
ShowUnInstDetails show


; インストーラーフィアルのバージョン情報記述
VIProductVersion ${PRODUCT_VERSION}
VIAddVersionKey ProductName     "${PRODUCT_NAME}"
VIAddVersionKey ProductVersion  "${PRODUCT_VERSION}"
VIAddVersionKey Comments        "Addin For Excel Library"
VIAddVersionKey LegalTrademarks ""
VIAddVersionKey LegalCopyright  "Copyright 2020 Bumpei.Koizumi"
VIAddVersionKey FileDescription "${PRODUCT_NAME}"
VIAddVersionKey FileVersion     "${PRODUCT_VERSION}"

; デフォルトのファイルのインストール先
 InstallDir "$appData\Bkoizumi\Ladex"

;実行権限 [user/admin]
RequestExecutionLevel user

;インストール画面構成
; !define MUI_LICENSEPAGE_RADIOBUTTONS      ; 「ライセンスに同意する」をラジオボタンにする
; !define MUI_FINISHPAGE_NOAUTOCLOSE        ; インストール完了後自動的に完了画面に遷移しないようにする

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
; !insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; アンインストール画面構成
UninstPage uninstConfirm
UninstPage instfiles

!insertmacro MUI_LANGUAGE "Japanese"

; インストール処理---------------------------------------------------------------------------------------
Section "Ladex" sec_Main
  SetOutPath "$appData\Microsoft\AddIns"
  File    "Ladex.xlam"

  SetOutPath $INSTDIR
  ; ディレクトリ/ファイルをコピー
  File    "ExcelOpen_ViewProtected.vbs"
  File    "README.pdf"
  ; File    "メンテナンス用.xlsm"
  CreateDirectory $INSTDIR\Images
  CreateDirectory $INSTDIR\log
  File /r "Ladex\RibbonImg"
  File /r "Ladex\RibbonSrc"
  File /r "Ladex\RibbonSrc"
  File    "Ladex\スタイル情報.xlsx"


  ; レジストリキーの設定
  WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstVersion" ${PRODUCT_VERSION}

  ; アドイン登録
  call AddinInstalled

  ; アンインストーラを出力
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ;スタートメニューの作成
  SetShellVarContext current
  CreateDirectory "$SMPROGRAMS\Bkoizumi"
  CreateDirectory "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}"
  CreateShortCut  "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\スタイル情報.lnk"   "$INSTDIR\スタイル情報.xlsx"
  CreateShortCut  "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\README.lnk"         "$INSTDIR\README.pdf"
  CreateShortCut  "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\Uninstall.lnk"      "$INSTDIR\Uninstall.exe"
SectionEnd

Section  "読み取り専用で開く" addReadOnly
  File    "ExcelOpen_ReadOnly.vbs"

  GetTempFileName $0
  File /oname=$0 `ExcelOpen_ReadOnly.vbs`
  nsExec::ExecToStack `"$SYSDIR\CScript.exe" $0  //e:vbscript //B //NOLOGO`
  ## Get & Test Return Code
  Pop $0
  DetailPrint `Return Code = $0`

  SetShellVarContext current
  CreateShortCut  "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\読み取り専用で開く.lnk"   "$INSTDIR\ExcelOpen_ReadOnly.vbs"

SectionEnd


Section "Uninstall"
  ; アドイン登録解除
  call un.AddinUninstalled

  ;スタートメニューから削除
  SetShellVarContext current
  Delete "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\Uninstall.lnk"
  Delete "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\読み取り専用で開く.lnk"
  Delete "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\スタイル情報.xlsx"
  Delete "$SMPROGRAMS\Bkoizumi\${PRODUCT_NAME}\README.lnk"
  RMDir /r  "$SMPROGRAMSBkoizumi\${PRODUCT_NAME}"
  RMDir /r  "$SMPROGRAMSBkoizumi"

  ; ディレクトリ削除
  RMDir /r "$INSTDIR"

  ; レジストリキー削除
  DeleteRegKey HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}"
SectionEnd

; セクションの説明文を入力
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
!insertmacro MUI_DESCRIPTION_TEXT ${sec_Main}       "Ladex インストール"
!insertmacro MUI_DESCRIPTION_TEXT ${addReadOnly}    "Excelファイルの右クリック「読み取り専用で開く」を有効にします。実行には管理者権限が必要です"
!insertmacro MUI_FUNCTION_DESCRIPTION_END


Function .onInit
  call  BootingCheck
  call  isInstalled
FunctionEnd


; Excelの起動確認---------------------------------------------------------------------------------------
Function BootingCheck

  ; reCheck:
  ; ${nsProcess::FindProcess} "EXCEL.EXE" $R0
  ; MessageBox MB_OK "[$R0]"
  ; ${If} $R0 = 0
  ; ${Else}
  ;   MessageBox MB_OK "Excel が起動しています"
  ;   ; nsProcess::_KillProcess "$1"
  ;   Pop $R0
  ;   Sleep 500
  ;   Goto reCheck
  ; ${EndIf}
FunctionEnd


; インストール済みかどうか------------------------------------------------------------------------------
Function isInstalled
  ; ReadRegStr $0 HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "InstVersion"

  ; ${If} $0 == ${PRODUCT_VERSION}
  ;   MessageBox MB_OK "同一バージョンがインストールされています"
    ; Abort

  ; ${Else}
  ;   SetOutPath $1
  ;   File "${APPDIR}\WebTools.xlsm"
  ;   File "${APPDIR}\var\WebCapture\新規Book.xlsm"
  ;   WriteRegStr HKCU "Software\VB and VBA Program Settings\${PRODUCT_NAME}\Main" "Version" ${PRODUCT_VERSION}
  ;   MessageBox MB_OK "既にバージョン $0 がインストールされているため、Excelのみ更新しました"
  ;   Abort
  ; ${EndIf}

FunctionEnd


; アドインの登録------------------------------------------------------------------------------------------
Function AddinInstalled
  GetTempFileName $0
  File /oname=$0 `install.vbs`
  nsExec::ExecToStack `"$SYSDIR\CScript.exe" $0  //e:vbscript //B //NOLOGO`
  ## Get & Test Return Code
  Pop $0
  DetailPrint `Return Code = $0`

FunctionEnd

; アドインの登録解除----------------------------------------------------------------------------------------
Function un.AddinUninstalled
  GetTempFileName $0
  File /oname=$0 `uninstall.vbs`
  nsExec::ExecToStack `"$SYSDIR\CScript.exe" $0  //e:vbscript //B //NOLOGO`
  ## Get & Test Return Code
  Pop $0
  DetailPrint `Return Code = $0`

FunctionEnd
