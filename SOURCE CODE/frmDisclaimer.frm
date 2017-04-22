VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDisclaimer 
   Caption         =   "[XL export - Nutzungsbedingungen]"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   OleObjectBlob   =   "frmDisclaimer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Modulbeschreibung:
'Texte f�r die Nutzungsbedingungen, die beim Aufrufen gezogen werden
'-------------------------------------------------------------------

Private Sub UserForm_Initialize()

    Select Case g_strSprache
        Case "DE"
            With frmDisclaimer
                .Caption = g_strTool & " - Nutzungsbedingungen"
                .lblDisclaimer.Caption = "Das Excel-Add-In 'XL export' darf ohne Einschr�nkung privat und " & _
                    "gewerblich verwendet werden. " & vbNewLine & vbNewLine & _
                    "Die Software wird mit gr��tm�glicher Sorgfalt entwickelt und getestet. " & _
                    "F�r Fehler im Code, die unkorrekte Ergebnisse liefern, Abst�rze des Programms oder des Systems " & _
                    "verursachen k�nnen, sowie f�r eventuellen Datenverlust durch Anwendung der Tools wird keine " & _
                    "Haftung �bernommen."
            End With
        Case "EN"
            With frmDisclaimer
                .Caption = g_strTool & " - Terms of use"
                .lblDisclaimer.Caption = "This tool is allowed for private and commercial use without any " & _
                    "restriction." & vbNewLine & vbNewLine & _
                    "The software is developed and tested with the utmost care. " & _
                    "There is no liability for code errors which can provide incorrect results, " & _
                    "crashes the program or the system as well as for potential data loss by using this tool."
            End With
        Case Else
        
    End Select
End Sub
