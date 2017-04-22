Attribute VB_Name = "modTOOLTIPtexte"
Option Explicit

'Modulbeschreibung:
'Ein- bzw. Ausblenden der Tooltips
'---------------------------------

Public Sub TooltipsON()
    Select Case g_strSprache
        Case "DE"
            With frmGUI
                'Allgemeine Elemente
                .imgLogo.ControlTipText = "minimieren/maximieren"
                .imgFlaggeDE.ControlTipText = ""
                .imgFlaggeEN.ControlTipText = "Switch to English"
                .imgRefresh.ControlTipText = "Vorschau aktualisieren"
                
                'Page "Optionen"
                .optBereichSheet.ControlTipText = "Aktuelles Tabellenblatt von A1 bis zur letzten benutzten Zeile/Spalte"
                .optBereichAlle.ControlTipText = "Alle Tabellenblätter von A1 bis zur letzten benutzten Zeile/Spalte in jeweils einer eigenen Datei"
                .optBereichSelektion.ControlTipText = "Selektierte(r) Bereich(e) des aktuellen Tabellenblatts"
                .optBereichUsed.ControlTipText = "Aktuelles Tabellenblatt von der ersten bis zur letzten benutzten Zeile/Spalte"
                .btnReset1.ControlTipText = "Alle Einstellungen auf Standardwerte zurücksetzen"
                
                'Page "Extras"
                .chkboxLeerzeichen.ControlTipText = "Entfernen von geschützten Leerzeichen (Unicode-Zeichen 160) und unnötigen Leerzeichen (am Anfang oder Ende einer Zelle und mehrfach aufeinanderfolgende Leerzeichen) in der Exportdatei"
                .chkboxSteuerzeichen.ControlTipText = "Entfernen von Steuerzeichen (7-Bit-ASCII-Zeichen 0-31 und Unicode-Zeichen 127, 129, 141, 143, 144 und 157) in der Exportdatei"
                .optSpeichernJa.ControlTipText = "Speichern der Tool-Einstellungen unter " & g_strSaveDateipfad & g_strSaveOptionen
                .optSpeichernNein.ControlTipText = "Tool-Einstellungen nicht speichern"
                .btnReset2.ControlTipText = "Tool-Einstellungen auf Standardwerte zurücksetzen"
                
                'Page "Spenden"
                .imgPrinz.ControlTipText = "Aufrufen der Website der Stiftung 'Große Hilfe für kleine Helden' im Standardbrowser"
                .imgQRcode.ControlTipText = "QR-Code scannen zum Aufruf des Online-Spendenformulars"
            End With
        Case "EN"
            With frmGUI
                'Allgemeine Elemente
                .imgLogo.ControlTipText = "minimize/maximize"
                .imgFlaggeDE.ControlTipText = "Auf Deutsch anzeigen"
                .imgFlaggeEN.ControlTipText = ""
                .imgRefresh.ControlTipText = "Refresh preview"
                
                'Page "Optionen"
                .optBereichSheet.ControlTipText = "Current worksheet from A1 to the last used row/column"
                .optBereichAlle.ControlTipText = "All worksheets from A1 to the last used row/column in separate files"
                .optBereichSelektion.ControlTipText = "Seleceted region(s) in the current worksheet"
                .optBereichUsed.ControlTipText = "Current worksheet from the first to the last used row/column"
                .btnReset1.ControlTipText = "Reset tool settings to default values"
                
                'Page "Extras"
                .chkboxLeerzeichen.ControlTipText = "Remove non-breaking spaces (Unicode-Zeichen 160) and unnecessary spaces (at the beginning or the end of a cell and multiple spaces within a cell) in the export file"
                .chkboxSteuerzeichen.ControlTipText = "Remove control characters (7-bit ASCII code characters 0-31 and unicode characters 127, 129, 141, 143, 144 and 157) in the export file"
                .optSpeichernJa.ControlTipText = "Save tool settings to " & g_strSaveDateipfad & g_strSaveOptionen
                .optSpeichernNein.ControlTipText = "Do not save tool settings"
                .btnReset2.ControlTipText = "Reset tool settings to default values"
                
                 'Page "Spenden"
                .imgPrinz.ControlTipText = "Open the web site of the foundation 'Große Hilfe für kleine Helden' in your standard browser"
                .imgQRcode.ControlTipText = "Scan this QR code for online donation"
            End With
        Case Else
    End Select
End Sub

Public Sub TooltipsOFF()
    With frmGUI
        'Allgemeine Elemente
        .imgLogo.ControlTipText = ""
        .imgFlaggeDE.ControlTipText = ""
        .imgFlaggeEN.ControlTipText = ""
        .imgRefresh.ControlTipText = ""
        
        'Page "Optionen"
        .optBereichSheet.ControlTipText = ""
        .optBereichAlle.ControlTipText = ""
        .optBereichSelektion.ControlTipText = ""
        .optBereichUsed.ControlTipText = ""
        .btnReset1.ControlTipText = ""
        
        'Page "Extras"
        .chkboxLeerzeichen.ControlTipText = ""
        .chkboxSteuerzeichen.ControlTipText = ""
        .optSpeichernJa.ControlTipText = ""
        .optSpeichernNein.ControlTipText = ""
        .btnReset2.ControlTipText = ""
        
        'Page "Spenden"
        .imgPrinz.ControlTipText = ""
        .imgQRcode.ControlTipText = ""
    End With
End Sub
