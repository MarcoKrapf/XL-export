VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVersionshinweise 
   Caption         =   "[XL export - Features und Versionshistorie]"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   OleObjectBlob   =   "frmVersionshinweise.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmVersionshinweise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Texte für die Versionshinweise, die beim Aufrufen gezogen werden
'----------------------------------------------------------------

Private Sub UserForm_Initialize()

    Select Case g_strSprache
        Case "DE"
            With frmVersionshinweise
                .Caption = g_strTool & " - Features und Versionshistorie"
                .lblVersionsInfo10a.Caption = "Version 1.0 (07.01.2016)"
                .lblVersionsInfo10b.Caption = "- Auswahl des Bereichs, der exportiert wird" & vbNewLine & _
                                                "- Export als Textdatei, CSV oder eigenes Format" & vbNewLine & _
                                                "- Trennzeichen frei wählbar" & vbNewLine & _
                                                "- Trennzeichen am Zeilenende können entfernt werden" & vbNewLine & _
                                                "- Unnötige Leerzeichen können entfernt werden" & vbNewLine & _
                                                "- Steuerzeichen können entfernt werden" & vbNewLine & _
                                                "- Konvertierung in Groß- oder Kleinbuchstaben möglich" & vbNewLine & _
                                                "- Vorschau der Exportdatei" & vbNewLine & _
                                                "- Automatisches Speichern der Tool-Einstellungen"
            End With
        Case "EN"
            With frmVersionshinweise
                .Caption = g_strTool & " - Features and version history"
                .lblVersionsInfo10a.Caption = "Version 1.0 (07.01.2016)"
                .lblVersionsInfo10b.Caption = "- Choice of the area to be exported" & vbNewLine & _
                                                "- Export as text file, CSV or custom file type" & vbNewLine & _
                                                "- Free choice of delimiter" & vbNewLine & _
                                                "- Delimiter at the end of a row can be deleted" & vbNewLine & _
                                                "- Unnecessary spaces can be deleted" & vbNewLine & _
                                                "- Control characters can be deleted" & vbNewLine & _
                                                "- Conversion to upper-case or lower-case letters possible" & vbNewLine & _
                                                "- Preview of the export file" & vbNewLine & _
                                                "- Automatically saving the tool settings"
            End With
        Case Else
        
    End Select
End Sub
