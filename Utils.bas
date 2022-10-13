Attribute VB_Name = "Utils"
Option Explicit

Sub Init()
    ActiveSheet.UnProtect "password"
    Sheets(ActiveSheet.Name).Cells.Clear
    Sheets(ActiveSheet.Name).Cells.Interior.ColorIndex = 2
    Sheets(ActiveSheet.Name).Cells.Locked = True
End Sub
