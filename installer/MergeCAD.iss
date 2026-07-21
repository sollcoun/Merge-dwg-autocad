; ============================================================
; MergeCAD Installer
; Inno Setup 6.x
; ============================================================

#define MyAppName "MergeCAD"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "MergeCAD"
#define MyAppExeName "MergeCAD.exe"

[Setup]

;--------------------------------------------------------------
; Основная информация
;--------------------------------------------------------------

AppId={{0CDBD347-D5E4-47D4-BD89-7A3C5E934A51}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}

;--------------------------------------------------------------
; Установка
;--------------------------------------------------------------

DefaultDirName={autopf}\MergeCAD
DefaultGroupName=MergeCAD
UsePreviousAppDir=yes
UsePreviousGroup=yes

DisableProgramGroupPage=yes

;--------------------------------------------------------------
; Вывод
;--------------------------------------------------------------

OutputDir=Output
OutputBaseFilename=MergeCAD_Setup_v1.0.0

;--------------------------------------------------------------
; Сжатие
;--------------------------------------------------------------

Compression=lzma2/max
SolidCompression=yes
LZMANumBlockThreads=4

;--------------------------------------------------------------
; Интерфейс
;--------------------------------------------------------------

WizardStyle=modern

;--------------------------------------------------------------
; Архитектура
;--------------------------------------------------------------

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

;--------------------------------------------------------------
; Права
;--------------------------------------------------------------

PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

;--------------------------------------------------------------
; Иконка и лицензия
;--------------------------------------------------------------

SetupIconFile=..\MergeCAD.ico
LicenseFile=..\license.txt

;--------------------------------------------------------------
; Информация о программе
;--------------------------------------------------------------

UninstallDisplayIcon={app}\MergeCAD.exe
UninstallDisplayName=MergeCAD

VersionInfoVersion={#MyAppVersion}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoProductName=MergeCAD
VersionInfoCompany=MergeCAD
VersionInfoDescription=Professional DWG Merge Tool
VersionInfoCopyright=Copyright © 2026 MergeCAD

;--------------------------------------------------------------
; Поведение
;--------------------------------------------------------------

AppMutex=MergeCAD

CloseApplications=yes
RestartApplications=no

ChangesAssociations=yes

;==============================================================
; Язык
;==============================================================

[Languages]

Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

;==============================================================
; Дополнительные задачи
;==============================================================

[Tasks]

Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные задачи:"

;==============================================================
; Папки
;==============================================================

[Dirs]

Name: "{userappdata}\MergeCAD"

;==============================================================
; Файлы
;==============================================================

[Files]

Source: "..\dist\MergeCAD.exe"; DestDir: "{app}"; Flags: ignoreversion

;==============================================================
; Ярлыки
;==============================================================

[Icons]

Name: "{group}\MergeCAD"; Filename: "{app}\MergeCAD.exe"

Name: "{group}\Удалить MergeCAD"; Filename: "{uninstallexe}"

Name: "{autodesktop}\MergeCAD"; Filename: "{app}\MergeCAD.exe"; Tasks: desktopicon

;==============================================================
; Запуск после установки
;==============================================================

[Run]

Filename: "{app}\MergeCAD.exe"; \
Description: "Запустить MergeCAD"; \
Flags: nowait postinstall skipifsilent

;==============================================================
; Удаление
;==============================================================

[UninstallDelete]

Type: filesandordirs; Name: "{userappdata}\MergeCAD"