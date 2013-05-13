VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kreiraj smetka"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim izlez As String
Dim status As String

COMportnum = 1      ' COM1

' WriteCommand(Br_na_COM_port, Sekventen_Broj, komanda+vlezni_parametri, izlez, status)

x = WriteCommand(COMportnum, 45, ">", izlez, status)     'Procitaj Datum/Cas od printerot
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
x = WriteCommand(COMportnum, 46, Chr(48)+"1,0000,1", izlez, status) 'Otvori fiskalna smetka
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
x = WriteCommand(COMportnum, 47, Chr(49)+"Kompjuter" + Chr(10) + "Pentium 4" + Chr(9) + Chr(192) + "16600.00*2.00,-10", izlez, status) ' Reg. prodazba
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
x = WriteCommand(COMportnum, 48, Chr(49)+"Monitor 'Soni'" + Chr(9) + Chr(192) + "12000", izlez, status) ' Reg. prodazba
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
x = WriteCommand(COMportnum, 49, Chr(53)+ Chr(9), izlez, status) ' Plakjanje na smetkata
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
x = WriteCommand(COMportnum, 50, Chr(56), izlez, status) ' Zatvaranje na smetkata
Text1.Text = Text1.Text + izlez + " | " + status + Chr(13) + Chr(10)
End Sub
