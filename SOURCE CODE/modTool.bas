Attribute VB_Name = "modTool"
Option Explicit

'Modulbeschreibung:
'Globale Variablen festlegen, Tool starten und GUI aufrufen
'----------------------------------------------------------

Public Const g_strTool As String = "XL export" 'Tool-Name
Public Const g_strVersion As String = "Version 1.0" 'Tool-Version
Public g_strSprache As String 'Kennzeichen für die Sprache der GUI
Public g_strSaveDateipfad As String 'Dateipfad für die Textdatei mit den Tool-Einstellungen
Public g_strSaveOptionen As String 'Dateiname der Textdatei mit den Tool-Einstellungen

Sub ToolStartenIconXLexport(control As IRibbonControl) 'Aufruf durch den Button im Ribbon
    Call ToolStarten
End Sub

Sub ToolStarten() 'Diese Prozedur manuell starten zum Testen der Entwicklung
    'Sprache
    g_strSprache = "DE"
    'GUI laden und starten
    Load frmGUI
    frmGUI.Show
End Sub
