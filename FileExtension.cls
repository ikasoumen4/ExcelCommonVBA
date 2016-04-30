VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function GetLinkAsString(LnkFilePathName As String) As String
'=========================================================================================
'機能      ：.lnkファイルが存在していれば、リンク先を返す
'引数      ：.lnkの拡張子をつけてパスを指定する
'作成日    ：2014/12/1
'未完成
'=========================================================================================


    Dim WSH As Object
    Dim LnkFile As Object
    
    Set WSH = CreateObject("WScript.Shell")
    
    Set LnkFile = WSH.CreateShortcut(LnkFilePathName)
    
    GetLinkAsString = LnkFile.TargetPath
    
    Set LnkFile = Nothing
    Set WSH = Nothing
    
End Function



Public Function ExistFile(TargetFilePath As String) As String
'=========================================================================================
'機能      ：ファイルが存在していれば True 存在していなければ False
'引数      ：ファイルのフルパス as String
'作成日    ：2014/12/1
'=========================================================================================

ExistFile = False
Dim FileName As String

FileName = Dir(TargetFilePath)

If FileName <> "" Then ExistFile = True


End Function

