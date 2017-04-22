VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGUI 
   Caption         =   "[Titel]"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "frmGUI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Haupt-Code des Tools mit den Funktionen der GUI
'-----------------------------------------------

'Variablen definieren
Dim varSave As Variant 'Infos aus Dialogbox zum Speichern
Dim strSheetSave As String 'Speichern von allen Tabellenblättern
Dim strSheetSaveInfo As String 'Ausgabeinfo für Export von allen Tabellenblättern
Dim rngSelection As Range 'Selektion auf dem Tabellenblatt
Dim strOutputZeile As String 'Zeile, die in die Ausgabedatei geschrieben wird
Dim strFileName As String 'Default-Filename
Dim intFileNr As Integer 'Nächste freie Nummer beim Export
Dim lngTrenn As Long 'Trennzeichen als Zeichencode
Dim strTyp As String 'Dateityp für den Export
Dim strTypText As String 'Bezeichnung im SaveAs-Dialog
Dim blnZeilenende As Boolean 'Kennzeichen ob am Ende der Zeile ein Trennzeichen steht
Dim strBereich As String 'Code für den Exportbereich
Dim blnEntfLeer As Boolean 'Kennzeichen ob überflüssige Leerzeichen entfernt werden sollen
Dim blnEntfSteuer As Boolean 'Kennzeichen ob Steuerzeichen entfernt werden sollen
Dim strKonvert As String 'Code für die Konvertierung von Buchstaben
Dim lngCheck As Long, strcheckZelle As String, strCheckZeichen As String 'Variablen für die Datenbereinigung
Dim lngAnzVorschau As Long, lngAnzVorAkt As Long 'Anzahl der Zeilen im Vorschaufenster und aktuelle Anzahl
Dim i As Long, j As Long, m As Long 'Zählvariablen für Schleifen
Dim objMail As Object 'Shell-Objekt für E-Mail
Dim opt1 As String, opt2 As String, opt3 As String, opt4 As String, opt5 As String, _
        opt6 As String, opt7 As String, opt8 As String, opt9 As String, opt10 As String, _
        opt11 As String, opt12 As String, opt13 As String, opt14 As String, opt15 As String, _
        opt16 As String, opt17 As String, opt18 As String, opt19 As String, opt20 As String, _
        opt21 As String, opt22 As String, opt23 As String, opt24 As String, opt25 As String _
        'Variablen mit den Werten der Tool-Einstellungen


Sub Export() 'Start des Exports

    'Wenn ein Fehler auftritt
    On Error GoTo Fehlerbehandlung
    
    'Optionen auslesen
    lngTrenn = Trennzeichen()
    strTyp = Dateityp()
    strTypText = DateitypText()
    blnZeilenende = Zeilenende()
    strBereich = Exportbereich()
    blnEntfLeer = EntfernenLeer()
    blnEntfSteuer = EntfernenSteuer()
    strKonvert = Konvertierung()

    'Kanal für den Output
    intFileNr = FreeFile 'Nächste freie Nummer zuweisen
    
    'Default-Name für die Ausgabedatei
    If InStr(ActiveWorkbook.Name, ".") > 0 Then
        strFileName = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    Else
        strFileName = ActiveWorkbook.Name
    End If
    
    'Speichern unter - Dialog
    varSave = Application.GetSaveAsFilename( _
        InitialFileName:=strFileName, _
        FileFilter:=DateitypText, _
        Title:=g_strTool & " - " & g_strVersion)
    
    'Wenn Dialog nicht abgebrochen wurde, dann exportieren
    If varSave <> False Then
        'Exportbereich auslesen
        Select Case strBereich
            Case "SHEET"
                Open varSave For Output As #intFileNr 'Ausgangskanal öffnen
                    Call expSHEET("hot")
                Close #intFileNr 'Ausgangskanal schließen
                'Pop-up
                MsgBox ("Export-Datei wurde erzeugt." & vbNewLine & vbNewLine & varSave), _
                    vbInformation, g_strTool & " - " & g_strVersion
            Case "ALL"
                Call expALL("hot")
                'Pop-up
                MsgBox ("Export-Dateien wurden erzeugt." & vbNewLine & vbNewLine & strSheetSaveInfo), _
                    vbInformation, g_strTool & " - " & g_strVersion
            Case "SEL"
                Open varSave For Output As #intFileNr 'Ausgangskanal öffnen
                    Call expSEL("hot")
                Close #intFileNr 'Ausgangskanal schließen
                'Pop-up
                MsgBox ("Export-Datei wurde erzeugt." & vbNewLine & vbNewLine & varSave), _
                    vbInformation, g_strTool & " - " & g_strVersion
            Case "USED"
                Open varSave For Output As #intFileNr 'Ausgangskanal öffnen
                    Call expUSED("hot")
                Close #intFileNr 'Ausgangskanal schließen
                'Pop-up
                MsgBox ("Export-Datei wurde erzeugt." & vbNewLine & vbNewLine & varSave), _
                    vbInformation, g_strTool & " - " & g_strVersion
        End Select
        
        Exit Sub
        
    End If
    
    Exit Sub
    
Fehlerbehandlung:
    MsgBox ("Fehler-Nr: " & Err.Number & vbNewLine & _
        "Beschreibung: " & Err.Description), vbExclamation, "Fehler beim Export"
End Sub


'Export
'------

Sub expSHEET(strCASE As String)
        For i = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            'Zeile zurücksetzen
            strOutputZeile = ""
            For j = 1 To ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
                'Zeile zusammenbauen
                strcheckZelle = checkLeerzeichen(ActiveSheet.Cells(i, j).Value)
                strcheckZelle = checkSteuerzeichen(strcheckZelle)
                strOutputZeile = strOutputZeile & strcheckZelle & Chr(lngTrenn)
            Next j
            strOutputZeile = konvertOutput(strOutputZeile)
            strOutputZeile = checkZeilenende(strOutputZeile)

            If strCASE = "hot" Then 'In Exportfile schreiben
                Print #intFileNr, strOutputZeile
            Else 'In Vorschaufenster schreiben
                listboxVorschau.AddItem
                listboxVorschau.List(listboxVorschau.ListCount - 1) = strOutputZeile
                lngAnzVorAkt = lngAnzVorAkt + 1 'Zähler hochzählen
                If lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl erreicht
            End If
        Next i
End Sub

Sub expALL(strCASE As String)
    strSheetSaveInfo = "" 'String für die Ausgabeinfo zurücksetzen
    'Alle Tabellenblätter durchlaufen
    For m = 1 To ActiveWorkbook.Worksheets.Count
    
        If strCASE = "hot" Then
            strSheetSave = Left(varSave, InStrRev(varSave, ".") - 1) & _
                "(" & ActiveWorkbook.Worksheets(m).Name & ")" & _
                Right(varSave, Len(varSave) - InStrRev(varSave, ".") + 1)
            strSheetSaveInfo = strSheetSaveInfo & strSheetSave & vbNewLine
            Open strSheetSave For Output As #intFileNr 'Ausgangskanal öffnen
        End If
    
        For i = 1 To ActiveWorkbook.Worksheets(m).UsedRange.SpecialCells(xlCellTypeLastCell).Row
            'Zeile zurücksetzen
            strOutputZeile = ""
            For j = 1 To ActiveWorkbook.Worksheets(m).UsedRange.SpecialCells(xlCellTypeLastCell).Column
                'Zeile zusammenbauen
                strcheckZelle = checkLeerzeichen(ActiveWorkbook.Worksheets(m).Cells(i, j).Value)
                strcheckZelle = checkSteuerzeichen(strcheckZelle)
                strOutputZeile = strOutputZeile & strcheckZelle & Chr(lngTrenn)
            Next j
            strOutputZeile = konvertOutput(strOutputZeile)
            strOutputZeile = checkZeilenende(strOutputZeile)

            If strCASE = "hot" Then 'In Exportfile schreiben
                Print #intFileNr, strOutputZeile
            Else 'In Vorschaufenster schreiben
                listboxVorschau.AddItem
                listboxVorschau.List(listboxVorschau.ListCount - 1) = strOutputZeile
                lngAnzVorAkt = lngAnzVorAkt + 1 'Zähler hochzählen
                If lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl erreicht
            End If
        Next i
        
        If strCASE = "hot" Then
            Close #intFileNr 'Ausgangskanal schließen
        Else
            If lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl erreicht
        End If
        
    Next m
End Sub

Sub expSEL(strCASE As String)
    Set rngSelection = Selection
    'Alle selektierten Bereiche durchlaufen
    For m = 1 To rngSelection.Areas.Count
    
        For i = rngSelection.Areas(m).Rows.Row To rngSelection.Areas(m).Rows.Row + rngSelection.Areas(m).Rows.Count - 1
        'Zeile zurücksetzen
        strOutputZeile = ""
            For j = rngSelection.Areas(m).Columns.Column To rngSelection.Areas(m).Columns.Column + rngSelection.Areas(m).Columns.Count - 1
                'Zeile zusammenbauen
                strcheckZelle = checkLeerzeichen(ActiveSheet.Cells(i, j).Value)
                strcheckZelle = checkSteuerzeichen(strcheckZelle)
                strOutputZeile = strOutputZeile & strcheckZelle & Chr(lngTrenn)
            Next j
            strOutputZeile = konvertOutput(strOutputZeile)
            strOutputZeile = checkZeilenende(strOutputZeile)

            If strCASE = "hot" Then 'In Exportfile schreiben
                Print #intFileNr, strOutputZeile
            Else 'In Vorschaufenster schreiben
                listboxVorschau.AddItem
                listboxVorschau.List(listboxVorschau.ListCount - 1) = strOutputZeile
                lngAnzVorAkt = lngAnzVorAkt + 1 'Zähler hochzählen
                If lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl der Vorschau erreicht
            End If
        Next i
        If strCASE = "cold" And lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl der Vorschau erreicht
    Next m
End Sub

Sub expUSED(strCASE As String)
    Set rngSelection = ActiveSheet.UsedRange
    For i = rngSelection.Rows.Row To rngSelection.Rows.Count + rngSelection.Rows.Row - 1
        'Zeile zurücksetzen
        strOutputZeile = ""
        For j = rngSelection.Columns.Column To rngSelection.Columns.Count + rngSelection.Columns.Column - 1
            'Zeile zusammenbauen
            strcheckZelle = checkLeerzeichen(ActiveSheet.Cells(i, j).Value)
            strcheckZelle = checkSteuerzeichen(strcheckZelle)
            strOutputZeile = strOutputZeile & strcheckZelle & Chr(lngTrenn)
        Next j
        strOutputZeile = konvertOutput(strOutputZeile)
        strOutputZeile = checkZeilenende(strOutputZeile)

        If strCASE = "hot" Then 'In Exportfile schreiben
            Print #intFileNr, strOutputZeile
        Else 'In Vorschaufenster schreiben
            listboxVorschau.AddItem
            listboxVorschau.List(listboxVorschau.ListCount - 1) = strOutputZeile
            lngAnzVorAkt = lngAnzVorAkt + 1 'Zähler hochzählen
            If lngAnzVorAkt > lngAnzVorschau Then Exit For 'maximale Anzahl erreicht
        End If
    Next i
End Sub


'Funktionen zum Anwenden der Exportfunktionen
'--------------------------------------------

Private Function checkZeilenende(strZeile As String) As String
    On Error Resume Next
    If blnZeilenende = True Then
        checkZeilenende = Left(strZeile, Len(strZeile) - 1) 'Trennzeichen am Ende abschneiden
    Else
        checkZeilenende = strZeile
    End If
    On Error GoTo 0
End Function

Private Function checkLeerzeichen(strZelle As String) As String
    If blnEntfLeer = True Then
        'String der Zelle durchlaufen
        For lngCheck = 1 To Len(strZelle)
            strCheckZeichen = Mid(strZelle, lngCheck, 1) 'Einzelnes Zeichen das untersucht wird
            If Asc(strCheckZeichen) = 160 Then 'Geschütztes Leerzeichen gefunden
                strZelle = Application.WorksheetFunction _
                    .Replace(strZelle, lngCheck, 1, Chr(32)) 'Geschütztes Leerzeichen durch normales ersetzen
            End If
        Next lngCheck
        'Überflüssige Leerzeichen entfernen und an Rückgabestring übergeben
        checkLeerzeichen = Application.WorksheetFunction.Trim(strZelle)
    Else
        checkLeerzeichen = strZelle
    End If
End Function

Private Function checkSteuerzeichen(strZelle As String) As String
    If blnEntfSteuer = True Then
        'String der Zelle durchlaufen
        For lngCheck = 1 To Len(strZelle)
            strCheckZeichen = Mid(strZelle, lngCheck, 1) 'Einzelnes Zeichen das untersucht wird
            Select Case Asc(strCheckZeichen) 'Asc gibt den Zeichencode des ersten Buchstabens zurück
                Case 1 To 31, 127, 129, 141, 143, 144, 157 'Steuerzeichen gefunden
                    strZelle = Application.WorksheetFunction _
                        .Replace(strZelle, lngCheck, 1, Chr(9)) 'Steuerzeichen durch horizontalen Tab ersetzen
            End Select
        Next lngCheck
        'Horizontale Tabs entfernen und an Rückgabestring übergeben
        checkSteuerzeichen = Application.WorksheetFunction.Clean(strZelle)
    Else
        checkSteuerzeichen = strZelle
    End If
End Function

Private Function konvertOutput(strZeile As String) As String
    Select Case strKonvert
        Case "GROSS"
            konvertOutput = UCase(strZeile)
        Case "KLEIN"
            konvertOutput = LCase(strZeile)
        Case "NICHT"
            konvertOutput = strZeile
    End Select
End Function


'Funtionen zum Auslesen der Exportoptionen
'-----------------------------------------

Private Function Trennzeichen()
    Select Case True
        Case optTrennSemikolon.Value
            Trennzeichen = 59 'Semikolon
        Case optTrennKomma.Value
            Trennzeichen = 44 'Komma
        Case optTrennTabstopp.Value
            Trennzeichen = 9 'Horizontaler Tabulator
        Case optTrennLeer.Value
            Trennzeichen = 32 'Leerzeichen
        Case optTrennCustom.Value
            On Error Resume Next
            Trennzeichen = Asc(txtboxTrennCustom.Value) 'Eigenes Zeichen
            On Error GoTo 0
    End Select
    If Trennzeichen = "" Then Trennzeichen = 32 'Default-Trennzeichen, falls nichts eingetragen ist
End Function

Private Function Dateityp()
    Select Case True
        Case optDateitypCSV.Value
            Dateityp = ".csv"
        Case optDateitypTXT.Value
            Dateityp = ".txt"
        Case optDateitypCustom.Value
            Dateityp = "." & txtboxDateitypCustom.Value 'Eigene Endung
    End Select
End Function

Private Function DateitypText()
    Select Case g_strSprache
        Case "DE"
            Select Case True
                Case optDateitypCSV.Value
                    DateitypText = "CSV-Datei (*.csv), *.csv" & "," & _
                                    "Textatei (*.txt), *.txt"
                Case optDateitypTXT.Value
                    DateitypText = "Textatei (*.txt), *.txt" & "," & _
                                    "CSV-Datei (*.csv), *.csv"
                Case optDateitypCustom.Value
                    DateitypText = "Eigener Dateityp (*" & strTyp & "), *" & strTyp & "," & _
                                    "CSV-Datei (*.csv), *.csv" & "," & _
                                    "Textatei (*.txt), *.txt"
            End Select
        Case "EN"
            Select Case True
                Case optDateitypCSV.Value
                    DateitypText = "CSV file (*.csv), *.csv" & "," & _
                                    "Text file (*.txt), *.txt"
                Case optDateitypTXT.Value
                    DateitypText = "Text file (*.txt), *.txt" & "," & _
                                    "CSV file (*.csv), *.csv"
                Case optDateitypCustom.Value
                    DateitypText = "Custom format (*" & strTyp & "), *" & strTyp & "," & _
                                    "CSV file (*.csv), *.csv" & "," & _
                                    "Text file (*.txt), *.txt"
            End Select
    End Select
End Function

Private Function Zeilenende()
    Select Case True
        Case optEndeEntfernen.Value
            Zeilenende = True
        Case optEndeBehalten.Value
            Zeilenende = False
    End Select
End Function

Private Function Exportbereich()
    Select Case True
        Case optBereichSheet.Value
            Exportbereich = "SHEET"
        Case optBereichAlle.Value
            Exportbereich = "ALL"
        Case optBereichSelektion.Value
            Exportbereich = "SEL"
        Case optBereichUsed.Value
            Exportbereich = "USED"
    End Select
End Function

Private Function EntfernenLeer()
        EntfernenLeer = chkboxLeerzeichen.Value
End Function

Private Function EntfernenSteuer()
        EntfernenSteuer = chkboxSteuerzeichen.Value
End Function

Private Function Konvertierung()
    Select Case True
        Case optKonvertGross.Value
            Konvertierung = "GROSS"
        Case optKonvertKlein.Value
            Konvertierung = "KLEIN"
        Case optKonvertNicht.Value
            Konvertierung = "NICHT"
    End Select
End Function


'Vorschaufenster
'---------------

Private Sub Vorschau()

        'Zähler zurücksetzen
        lngAnzVorAkt = 1

        'Vorschaufenster leeren
        listboxVorschau.Clear

        'Optionen auslesen
        lngTrenn = Trennzeichen()
        blnZeilenende = Zeilenende()
        strBereich = Exportbereich()
        blnEntfLeer = EntfernenLeer()
        blnEntfSteuer = EntfernenSteuer()
        strKonvert = Konvertierung()
        
        'Vorschau ausgeben
        Select Case strBereich
            Case "SHEET"
                Call expSHEET("cold")
            Case "ALL"
                Call expALL("cold")
            Case "SEL"
                Call expSEL("cold")
            Case "USED"
                Call expUSED("cold")
        End Select
End Sub


'Popups
'------

Private Sub FeaturesAnzeigen() 'Öffnen bzw. schließen des Popups
    If frmVersionshinweise.Visible = False Then
        Load frmVersionshinweise
        frmVersionshinweise.StartUpPosition = 2 'Zentriert auf dem gesamten Bildschirm
        frmVersionshinweise.Show
    Else
        Unload frmVersionshinweise
    End If
End Sub

Private Sub DisclaimerAnzeigen() 'Öffnen bzw. schließen des Popups
    If frmDisclaimer.Visible = False Then
        Load frmDisclaimer
        frmDisclaimer.StartUpPosition = 2 'Zentriert auf dem gesamten Bildschirm
        frmDisclaimer.Show
    Else
        Unload frmDisclaimer
    End If
End Sub


'Tooltips
'--------

Private Sub Tooltips()
    If checkboxTooltip.Value = True Then
        Call modTOOLTIPtexte.TooltipsON
    Else
        Call modTOOLTIPtexte.TooltipsOFF
    End If
End Sub


'Reset
'-----

Private Sub Reset() 'Default-Werte setzen
    'Registerkarte "Export"
    txtboxTrennCustom.Value = ""
    txtboxDateitypCustom.Value = ""
    optTrennSemikolon.Value = True
    optDateitypCSV.Value = True
    optEndeEntfernen.Value = True
    optBereichSheet.Value = True
    'Registerkarte "Extras"
    chkboxLeerzeichen.Value = False
    chkboxSteuerzeichen.Value = False
    optKonvertNicht.Value = True
    optSpeichernJa.Value = True
End Sub


'Einstellungen laden
'-------------------

Private Sub EinstellungenLesen()
    On Error Resume Next 'Falls ein Fehler auftritt: Anweisung überspringen
    
    'Kanal für den Input
    intFileNr = FreeFile 'Nächste freie Nummer zuweisen
    
    Open g_strSaveDateipfad & g_strSaveOptionen For Input As #intFileNr 'Eingangskanal öffnen
        Line Input #intFileNr, opt1
        Line Input #intFileNr, opt2
        Line Input #intFileNr, opt3
        Line Input #intFileNr, opt4
        Line Input #intFileNr, opt5
        Line Input #intFileNr, opt6
        Line Input #intFileNr, opt7
        Line Input #intFileNr, opt8
        Line Input #intFileNr, opt9
        Line Input #intFileNr, opt10
        Line Input #intFileNr, opt11
        Line Input #intFileNr, opt12
        Line Input #intFileNr, opt13
        Line Input #intFileNr, opt14
        Line Input #intFileNr, opt15
        Line Input #intFileNr, opt16
        Line Input #intFileNr, opt17
        Line Input #intFileNr, opt18
        Line Input #intFileNr, opt19
        Line Input #intFileNr, opt20
        Line Input #intFileNr, opt21
        Line Input #intFileNr, opt22
        Line Input #intFileNr, opt23
        Line Input #intFileNr, opt24
        Line Input #intFileNr, opt25
    Close #intFileNr 'Eingangskanal schließen
    
    'Ausgelesene Werte in GUI setzen
    g_strSprache = opt1
    checkboxTooltip.Value = CBool(opt2)
    txtboxTrennCustom.Value = opt8
    optTrennSemikolon.Value = CBool(opt3)
    optTrennLeer.Value = CBool(opt4)
    optTrennTabstopp.Value = CBool(opt5)
    optTrennKomma.Value = CBool(opt6)
    optTrennCustom.Value = CBool(opt7)
    txtboxDateitypCustom.Value = opt12
    optDateitypCSV.Value = CBool(opt9)
    optDateitypTXT.Value = CBool(opt10)
    optDateitypCustom.Value = CBool(opt11)
    optEndeEntfernen.Value = CBool(opt13)
    optEndeBehalten.Value = CBool(opt14)
    optBereichSheet.Value = CBool(opt15)
    optBereichAlle.Value = CBool(opt16)
    optBereichSelektion.Value = CBool(opt17)
    optBereichUsed.Value = CBool(opt18)
    chkboxLeerzeichen.Value = CBool(opt19)
    chkboxSteuerzeichen.Value = CBool(opt20)
    optKonvertGross.Value = CBool(opt21)
    optKonvertKlein.Value = CBool(opt22)
    optKonvertNicht.Value = CBool(opt23)
    optSpeichernJa.Value = CBool(opt24)
    optSpeichernNein.Value = CBool(opt25)
    
    On Error GoTo 0
End Sub


'Einstellungen speichern
'-----------------------

Private Sub EinstellungenSpeichern()
    On Error Resume Next 'Falls ein Fehler auftritt: Anweisung überspringen
    
    'Kanal für den Output
    intFileNr = FreeFile 'Nächste freie Nummer zuweisen
    
    'Tool-Einstellungen speichern
    Open g_strSaveDateipfad & g_strSaveOptionen For Output As #intFileNr 'Ausgangskanal öffnen
        'Werte in Textdatei schreiben (-1 für TRUE, 0 für FALSE)
        Print #intFileNr, g_strSprache
        Print #intFileNr, CInt(checkboxTooltip.Value)
        Print #intFileNr, CInt(optTrennSemikolon.Value)
        Print #intFileNr, CInt(optTrennLeer.Value)
        Print #intFileNr, CInt(optTrennTabstopp.Value)
        Print #intFileNr, CInt(optTrennKomma.Value)
        Print #intFileNr, CInt(optTrennCustom.Value)
        Print #intFileNr, txtboxTrennCustom.Value
        Print #intFileNr, CInt(optDateitypCSV.Value)
        Print #intFileNr, CInt(optDateitypTXT.Value)
        Print #intFileNr, CInt(optDateitypCustom.Value)
        Print #intFileNr, txtboxDateitypCustom.Value
        Print #intFileNr, CInt(optEndeEntfernen.Value)
        Print #intFileNr, CInt(optEndeBehalten.Value)
        Print #intFileNr, CInt(optBereichSheet.Value)
        Print #intFileNr, CInt(optBereichAlle.Value)
        Print #intFileNr, CInt(optBereichSelektion.Value)
        Print #intFileNr, CInt(optBereichUsed.Value)
        Print #intFileNr, CInt(chkboxLeerzeichen.Value)
        Print #intFileNr, CInt(chkboxSteuerzeichen.Value)
        Print #intFileNr, CInt(optKonvertGross.Value)
        Print #intFileNr, CInt(optKonvertKlein.Value)
        Print #intFileNr, CInt(optKonvertNicht.Value)
        Print #intFileNr, CInt(optSpeichernJa.Value)
        Print #intFileNr, CInt(optSpeichernNein.Value)
    Close #intFileNr 'Ausgangskanal schließen
    
    On Error GoTo 0
End Sub


'Hyperlinks
'----------

Private Sub SourceCodeURL()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:="https://github.com/MarcoKrapf/XL-export"
    On Error GoTo 0
End Sub

Private Sub SpendenLinkURLaufrufen()
    On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:="http://www.ghfkh.de/"
    On Error GoTo 0
End Sub


'E-Mails
'-------

Private Sub eMail() 'Feedback E-Mail
    On Error Resume Next
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & "excel@marco-krapf.de" _
            & "&subject=" & "Feedback: " & g_strTool & " - " & g_strVersion & " / " _
            & Application.OperatingSystem & " / Excel-Version " & Application.Version
    On Error GoTo 0
End Sub


'Change-Ereignisse
'-----------------

Private Sub txtboxTrennCustom_Change() 'Eintrag in das Custom-Feld für das Trennzeichen
    optTrennCustom.Value = True 'Radio-Button aktivieren
    If txtboxTrennCustom.Value <> "" Then
        lblTrennzeichen.Caption = Chr(Asc(txtboxTrennCustom.Value)) 'Gibt für ungültige Zeichen ein ? aus
        lblTrennzeichen.ControlTipText = "ASCII-Code " & Asc(txtboxTrennCustom.Value)
    Else
        lblTrennzeichen.Caption = ""
        lblTrennzeichen.ControlTipText = "ASCII-Code 32"
    End If
    Call Vorschau
End Sub

Private Sub txtboxDateitypCustom_Change() 'Eintrag in das Custom-Feld für den Dateityp
    On Error Resume Next 'Bei Fehler (z.B. Leerstring): Anweisung überspringen
    
    optDateitypCustom.Value = True 'Radio-Button aktivieren
    
    'Alle Leerzeichen entfernen
    strTyp = txtboxDateitypCustom.Value 'Custom-Dateityp auslesen
    m = 0 'Zähler zurücksetzen
    For lngCheck = 1 To Len(strTyp)
        strCheckZeichen = Mid(strTyp, lngCheck - m, 1) 'Einzelnes Zeichen das untersucht wird
            If Asc(strCheckZeichen) = 32 Then 'Leerzeichen gefunden
                strTyp = Application.WorksheetFunction _
                    .Replace(strTyp, lngCheck - m, 1, "") 'Leerzeichen entfernen
                m = m + 1 'Zähler hochzählen
            End If
    Next lngCheck
    
    lblDateityp.Caption = "." & strTyp 'Dateityp in der GUI anzeigen
    
    On Error GoTo 0
End Sub

Private Sub spinVorschau_Change() 'Anzahl der Zeilen in der Vorschau
    lngAnzVorschau = spinVorschau.Value
    If lngAnzVorschau < 5 Then
        lblAnzVorZahl.Caption = 1
    Else
        lblAnzVorZahl.Caption = lngAnzVorschau
    End If
    Call Vorschau
End Sub


'Klick-Ereignisse
'----------------

Private Sub imgLogo_Click() 'GUI minimieren/maximieren
    If frmGUI.Width = 300 Then
        frmGUI.Width = 240
        frmGUI.Height = 66
    Else:
        frmGUI.Width = 300
        frmGUI.Height = 382
    End If
End Sub

Private Sub btnReset1_Click()
    Call Reset
End Sub

Private Sub btnReset2_Click()
    Call Reset
End Sub

Private Sub btnStartExport1_Click()
    Call Export
End Sub

Private Sub btnStartExport2_Click()
    Call Export
End Sub

Private Sub checkboxTooltip_Click()
    Call Tooltips
End Sub

Private Sub imgFlaggeDE_Click()
    g_strSprache = "DE"
    Call modGUItexte.Sprache(g_strSprache)
    If checkboxTooltip.Value = True Then Call modTOOLTIPtexte.TooltipsON
End Sub

Private Sub imgFlaggeEN_Click()
    g_strSprache = "EN"
    Call modGUItexte.Sprache(g_strSprache)
    If checkboxTooltip.Value = True Then Call modTOOLTIPtexte.TooltipsON
End Sub

Private Sub btnAnleitung_Click()
MsgBox "richtige anleitung reinmachen"
    Call AnleitungAnzeigen
End Sub

Private Sub btnFeatures_Click()
    Call FeaturesAnzeigen
End Sub

Private Sub btnDisclaimer_Click()
    Call DisclaimerAnzeigen
End Sub

Private Sub btnSourceCode_Click()
    Call SourceCodeURL
End Sub

Private Sub btnFeedback_Click()
    Call eMail
End Sub

Private Sub imgPrinz_Click()
    Call SpendenLinkURLaufrufen
End Sub

Private Sub lblSpendenLink_Click()
    Call SpendenLinkURLaufrufen
End Sub

Private Sub imgRefresh_Click()
    Call Vorschau
End Sub

Private Sub chkboxLeerzeichen_Click()
    Call Vorschau
End Sub

Private Sub chkboxSteuerzeichen_Click()
    Call Vorschau
End Sub

Private Sub optBereichAlle_Click()
    Call Vorschau
End Sub

Private Sub optBereichSelektion_Click()
    Call Vorschau
End Sub

Private Sub optBereichSheet_Click()
    Call Vorschau
End Sub

Private Sub optBereichUsed_Click()
    Call Vorschau
End Sub

Private Sub optEndeBehalten_Click()
    Call Vorschau
End Sub

Private Sub optEndeEntfernen_Click()
    Call Vorschau
End Sub

Private Sub optKonvertGross_Click()
    Call Vorschau
End Sub

Private Sub optKonvertKlein_Click()
    Call Vorschau
End Sub

Private Sub optKonvertNicht_Click()
    Call Vorschau
End Sub

Private Sub optTrennCustom_Click()
    lblTrennzeichen.Caption = txtboxTrennCustom.Value
    If txtboxTrennCustom.Value <> "" Then
        lblTrennzeichen.ControlTipText = "ASCII-Code " & Asc(txtboxTrennCustom.Value)
    Else
        lblTrennzeichen.ControlTipText = "ASCII-Code 32"
    End If
    Call Vorschau
End Sub

Private Sub optTrennKomma_Click()
    lblTrennzeichen.Caption = ","
    lblTrennzeichen.ControlTipText = "ASCII-Code " & Asc(",")
    Call Vorschau
End Sub

Private Sub optTrennLeer_Click()
    lblTrennzeichen.Caption = " "
    lblTrennzeichen.ControlTipText = "ASCII-Code " & Asc(" ")
    Call Vorschau
End Sub

Private Sub optTrennSemikolon_Click()
    lblTrennzeichen.Caption = ";"
    lblTrennzeichen.ControlTipText = "ASCII-Code " & Asc(";")
    Call Vorschau
End Sub

Private Sub optTrennTabstopp_Click()
    lblTrennzeichen.Caption = "TAB"
    lblTrennzeichen.ControlTipText = "ASCII-Code 9"
    Call Vorschau
End Sub

Private Sub optDateitypCSV_Click()
    lblDateityp.Caption = ".csv"
End Sub

Private Sub optDateitypTXT_Click()
    lblDateityp.Caption = ".txt"
End Sub

Private Sub optDateitypCustom_Click()
    lblDateityp.Caption = "." & strTyp
End Sub


'GUI initialisieren
'------------------

Private Sub UserForm_Initialize()

    'Standardordner für Office-Add-Ins und Dateinamen für Systemdateien setzen
    g_strSaveDateipfad = Application.UserLibraryPath
    g_strSaveOptionen = "settings.xlexport"
    
    'Gespeicherte Tool-Einstellungen laden wenn Datei vorhanden
    If Dir(g_strSaveDateipfad & g_strSaveOptionen) <> "" Then
        Call EinstellungenLesen
    Else
        Call Reset
    End If
    
    'Größe der GUI
    frmGUI.Width = 300
    frmGUI.Height = 382

    'Erste Seite aktivieren
    MultiPageGUI.Value = 0
    
    'GUI-Beschriftungen
    Call modGUItexte.Sprache(g_strSprache)
    
    'Default-Werte setzen
        'Registerkarte "Export"
        txtboxTrennCustom.MaxLength = 1 'Max. 1 Zeichen
        txtboxDateitypCustom.MaxLength = 6 'Max. 6 Zeichen
        'Registerkarte "Extras"
        spinVorschau.Value = 5
        lngAnzVorschau = spinVorschau.Value
        lblAnzVorZahl.Caption = lngAnzVorschau
    
    'Vorschaufenster füllen
    Call Vorschau
        
End Sub


'GUI schließen
'-------------

Private Sub UserForm_Terminate()
    If optSpeichernJa.Value = True Then Call EinstellungenSpeichern
End Sub
