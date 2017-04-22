Attribute VB_Name = "modGUItexte"
Option Explicit

'Modulbeschreibung:
'Anpassung der GUI-Beschriftungen je nach gewählter Sprache
'----------------------------------------------------------

Public Sub Sprache(strSprachwahl As String)

    Select Case strSprachwahl
        Case "DE"
            With frmGUI
                'Allgemeine Elemente
                .Caption = g_strTool & " - " & g_strVersion
                .checkboxTooltip.Caption = "Tooltips"
                .frmVorschau.Caption = "Vorschau"
                .lblAnzahl.Caption = "Zeilen"
                
                'MultiPage
                .MultiPageGUI.Pages(0).Caption = "Optionen"
                .MultiPageGUI.Pages(1).Caption = "Extras"
                .MultiPageGUI.Pages(2).Caption = "Info"
                .MultiPageGUI.Pages(3).Caption = "Spende"
                
                'Page "Optionen"
                .frmTrennzeichen.Caption = "Trennzeichen"
                    .optTrennSemikolon.Caption = "Semikolon [;]"
                    .optTrennKomma.Caption = "Komma [,]"
                    .optTrennTabstopp.Caption = "Tabulator"
                    .optTrennLeer.Caption = "Leerzeichen"
                    .optTrennCustom.Caption = "Anderes"
                .frmZeilenende.Caption = "Am Zeilenende"
                    .optEndeEntfernen.Caption = "Trennzeichen entfernen"
                    .optEndeBehalten.Caption = "Trennzeichen erhalten"
                .frmDateityp.Caption = "Dateityp"
                    .optDateitypCSV.Caption = "CSV-Datei"
                    .optDateitypTXT.Caption = "Textdatei"
                    .optDateitypCustom.Caption = "Eigener"
                .frmExportbereich.Caption = "Exportbereich"
                    .optBereichSheet.Caption = "Dieses Tabellenblatt"
                    .optBereichAlle.Caption = "Alle Tabellenblätter"
                    .optBereichSelektion.Caption = "Selektierter Bereich"
                    .optBereichUsed.Caption = "Benutzter Bereich"
                .btnStartExport1.Caption = "Export starten"
                .btnReset1.Caption = "Reset"
                
                'Page "Extras"
                .frmOptionen.Caption = "Datenbereinigung beim Export"
                    .chkboxLeerzeichen.Caption = "Unnötige und geschützte Leerzeichen entfernen"
                    .chkboxSteuerzeichen.Caption = "Steuerzeichen entfernen"
                .frmKonvertieren.Caption = "Konvertierung beim Export"
                    .optKonvertGross.Caption = "Alle Buchstaben zu Großbuchstaben umwandeln"
                    .optKonvertKlein.Caption = "Alle Buchstaben zu Kleinbuchstaben umwandeln"
                    .optKonvertNicht.Caption = "Nicht konvertieren"
                .frmSpeichern.Caption = "Beim Beenden des Tools"
                    .optSpeichernJa.Caption = "Einstellungen speichern"
                    .optSpeichernNein.Caption = "Einstellungen verwerfen"
                .btnStartExport2.Caption = "Export starten"
                .btnReset2.Caption = "Reset"
                
                'Page "Info"
                .btnFeatures.Caption = "Features und Versionshistorie"
                .btnSourceCode.Caption = "Quellcode auf GitHub"
                .btnDisclaimer.Caption = "Nutzungsbedingungen"
                .btnFeedback.Caption = "Feedback"
                .lblInfo1.Caption = g_strTool & " - " & g_strVersion & " (Januar 2017)"
                .lblInfo2.Caption = "Autor: Marco Krapf - E-Mail: excel@marco-krapf.de"
                
                'Page "Spenden"
                .lblSpendenLink.Caption = "Info und Spende an die Stiftung 'Große Hilfe für kleine Helden'"
                .lblSpendenText.Caption = "Das Excel-Add-in 'XL export' wird privat entwickelt und unter " & _
                    "http://marco-krapf.de/excel/ kostenlos zum Download angeboten." & vbNewLine & vbNewLine & _
                    "Über eine kleine Spende an die Stiftung 'Große Hilfe für kleine Helden' für kranke Kinder " & _
                    "in der Region Heilbronn würde ich mich sehr freuen."
            End With
        Case "EN"
            With frmGUI
                'Allgemeine Elemente
                .Caption = g_strTool & " - " & g_strVersion
                .checkboxTooltip.Caption = "Tooltips"
                .frmVorschau.Caption = "Preview"
                .lblAnzahl.Caption = "Rows"
                
                'MultiPage
                .MultiPageGUI.Pages(0).Caption = "Options"
                .MultiPageGUI.Pages(1).Caption = "Extras"
                .MultiPageGUI.Pages(2).Caption = "Info"
                .MultiPageGUI.Pages(3).Caption = "Donation"
                
                'Page "Optionen"
                .frmTrennzeichen.Caption = "Delimiter"
                    .optTrennSemikolon.Caption = "Semicolon [;]"
                    .optTrennKomma.Caption = "Comma [,]"
                    .optTrennTabstopp.Caption = "Tabulator"
                    .optTrennLeer.Caption = "Space"
                    .optTrennCustom.Caption = "Other"
                .frmZeilenende.Caption = "At the end of a line"
                    .optEndeEntfernen.Caption = "Remove delimiter"
                    .optEndeBehalten.Caption = "Keep delimiter"
                .frmDateityp.Caption = "File format"
                    .optDateitypCSV.Caption = "CSV file"
                    .optDateitypTXT.Caption = "Text file"
                    .optDateitypCustom.Caption = "Custom ."
                .frmExportbereich.Caption = "Area to export"
                    .optBereichSheet.Caption = "This sheet"
                    .optBereichAlle.Caption = "All sheets"
                    .optBereichSelektion.Caption = "Selected area"
                    .optBereichUsed.Caption = "Used range"
                .btnStartExport1.Caption = "Start export"
                .btnReset1.Caption = "Reset"
                
                'Page "Extras"
                .frmOptionen.Caption = "Data cleanup on export"
                    .chkboxLeerzeichen.Caption = "Remove unnecessary and non-breaking spaces"
                    .chkboxSteuerzeichen.Caption = "Remove control characters"
                .frmKonvertieren.Caption = "Conversion on export"
                    .optKonvertGross.Caption = "Convert all characters to upper-case letters"
                    .optKonvertKlein.Caption = "Convert all characters to lower-case letters"
                    .optKonvertNicht.Caption = "No conversion"
                .frmSpeichern.Caption = "When closing the tool"
                    .optSpeichernJa.Caption = "Save settings"
                    .optSpeichernNein.Caption = "Discard settings"
                .btnStartExport2.Caption = "Start export"
                .btnReset2.Caption = "Reset"
                
                'Page "Info"
                .btnFeatures.Caption = "Features and version history"
                .btnSourceCode.Caption = "Source code on GitHub"
                .btnDisclaimer.Caption = "Terms of use"
                .btnFeedback.Caption = "Feedback"
                .lblInfo1.Caption = g_strTool & " - " & g_strVersion & " (Jan 2017)"
                .lblInfo2.Caption = "Author: Marco Krapf - email: excel@marco-krapf.de"
                
                'Page "Spenden"
                .lblSpendenLink.Caption = "Info and donation to the foundation"
                .lblSpendenText.Caption = "This add-in is being developed and maintained with private effort " & _
                    "and provided for free download on http://marco-krapf.de/excel/" & vbNewLine & vbNewLine & _
                    "I would be very happy about a small donation to this foundation for sick children in the " & _
                    "region of Heilbronn/Germany."
            End With
        Case Else
    End Select
End Sub
