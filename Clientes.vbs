' =========================================================
' CorpSync Solutions - Criador de Atalho (Versao de Teste)
' Cria um atalho .lnk na Area de Trabalho que executa
' um comando PowerShell inofensivo para diagnostico.
' =========================================================

Option Explicit

Dim oWS, sLinkFile, oLink, sLogFile, oFSO, oLog

Set oWS = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

sLinkFile = oWS.SpecialFolders("Desktop") & "\CorpSync_Teste.lnk"
sLogFile  = oWS.SpecialFolders("Desktop") & "\CorpSync_Log.txt"

Set oLog = oFSO.CreateTextFile(sLogFile, True)
oLog.WriteLine "===== LOG DE DIAGNOSTICO CORPSYNC ====="
oLog.WriteLine "Data: " & Now
oLog.WriteLine "Caminho do atalho: " & sLinkFile
oLog.WriteLine "----------------------------------------"

On Error Resume Next
Set oLink = oWS.CreateShortcut(sLinkFile)

oLink.TargetPath = "powershell.exe"
oLog.WriteLine "TargetPath: " & oLink.TargetPath

oLink.Arguments = "-NoExit -Command ""Write-Host 'ATALHO FUNCIONANDO! Mecanica OK.' -ForegroundColor Green"""
oLog.WriteLine "Arguments: " & oLink.Arguments

oLink.Description = "Atalho de teste CorpSync - Diagnostico"
oLog.WriteLine "Description: " & oLink.Description

oLink.WindowStyle = 1
oLink.IconLocation = "powershell.exe, 0"
oLink.WorkingDirectory = oWS.SpecialFolders("Desktop")

oLink.Save

If Err.Number <> 0 Then
    oLog.WriteLine "ERRO ao salvar atalho: " & Err.Description
    MsgBox "Erro ao criar atalho: " & Err.Description, vbCritical
Else
    oLog.WriteLine "STATUS: Atalho criado com sucesso."
End If

oLog.WriteLine "----------------------------------------"
oLog.WriteLine "FIM DO LOG"
oLog.Close

Set oLink = Nothing
Set oWS = Nothing
Set oFSO = Nothing

MsgBox "Atalho 'CorpSync_Teste.lnk' criado na Area de Trabalho!" & vbCrLf & _
       "Verifique tambem o arquivo 'CorpSync_Log.txt' para diagnostico.", _
       vbInformation, "CorpSync - Teste"
