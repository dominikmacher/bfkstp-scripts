Windows 11:

Dieser Fehler wird durch die PowerShell-Sicherheitsrichtlinie verursacht, die standardmäßig Skripte blockiert. 

Um Skripte auszuführen:
1) öffnen Sie PowerShell als Administrator
2) Get-ExecutionPolicy 
3) Set-ExecutionPolicy RemoteSigned 

Alternativ - nur für aktuelle Powershell-Sitzung setzen: 
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass 