Class DatabaseConnector

'Uživatel, pod kterým se bude hlasít
Dim usr
'Heslo pro uživatele
Dim passwd
'Soubor, kam se budou ukládát logy
Dim logF
'Cesta k BTEQ
Dim bteq

Private Sub Class_Initialize()
 usr = ""
 passwd = ""
 logF = "log.txt"
 bteq = "c:\Program Files\Teradata\Client\13.10\bin\bteq.exe"
End Sub

'Nastavení uživatele
Property Let user(u)
  usr = u  
End Property

'Nastavení hesla
Property Let password(p)
  passwd = p  
End Property

'Nastavení logovacího souboru
Property Let logFile(l)
  logF = l  
End Property

'Připojí se k databázi a pustí příkaz
Sub executeCommands(s)

    'Musí být nastavený uživatel
    If  (usr = "") Then
      MsgBox "UserName is not set"
    Else 
      'Musí být nastaveno heslo
      If (passwd = "") Then
        MsgBox "Password is not set"
      Else
        Dim tmpFile
        Dim objOutFile
        Dim objFS
        Dim shell
        'Pouští command line, přes kterou pouští exe
        Set objFS = CreateObject("Scripting.FileSystemObject")

	'Celý skript uloží do souboru, který poté spustí.
        tmpFile = "tmp.txt"
        Set objOutFile = objFS.CreateTextFile(tmpFile,True)
        objOutFile.Write(s)
        objOutFile.Close
        Set shell = CreateObject("WScript.Shell")
	'Pustí BTEQ s přihlášením pro uživatele
        shell.Run Chr(34) & bteq & Chr(34) & ".logon " & usr & "," & passwd & "<" & tmpFile & ">" & logF, 1, True
        Set shell = Nothing
        objFS.DeleteFile(tmpFile)
      End If
    End If
    
End Sub


End Class
