Attribute VB_Name = "StartWatch"
Option Explicit
Public Sub Go()
    Shell "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ""D:\Onedrive\MCI Excel\2_Dataset\_generator\snapwatch.ps1""", vbHide
End Sub
