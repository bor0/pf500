Attribute VB_Name = "FiscalModule"
' $Id: FiscalModule.bas,v 1.5.0.stable 2002/07/29 18:55:15$
' Description: Modul za VB6 za povrzuvanje na pf500.dll so Visual Basic 6
' Author: Vasko Mitanov (vasko@accent.com.mk), Accent Computers
'
' Public Function WriteCommand(ByVal com_port As Byte, ByVal seq As Byte, ByVal sVlez As String, ByRef sIzlez As String, ByRef sStatus As String) As Integer
'
' Parametri:
'            [IN]   com_port : Broj na COM portot (1-255)
'            [IN]   seq      : Sekvencen broj na komandata, (20h-7Fh)
'            [IN]   sVlez    : Komandata + Vleznite podatoci do printerot
'            [OUT]  sIzlez   : Odgovor od komandata
'            [OUT]  sStatus  : Status na printerot po izvrsenata komanda
'

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal source As Long, ByVal bytes As Long)

'__declspec(dllexport) HANDLE InitComm(unsigned char portnum)
'__declspec(dllexport) int WriteCommand(HANDLE porthnd, unsigned char seq, LPCTSTR cmd, LPTSTR output, LPTSTR status)
'__declspec(dllexport) int CloseComm(HANDLE porthnd)

Private Declare Function InitComm Lib "pf500.dll" (ByVal portnum As Byte) As Long
Private Declare Function SendCommand Lib "pf500.dll" Alias "WriteCommand" (ByVal porthnd As Long, ByVal seq As Byte, ByVal vlez As String, ByVal izlez As Long, ByVal stats As Long) As Integer
Private Declare Function CloseComm Lib "pf500.dll" (ByVal porthnd As Long) As Integer

Public Function AnsiZtoString(ByVal strz As Long) As String
    ' Funkcija za konverzija od bafer so bajti vo VB string
    Dim byteCh(1) As Byte
    Dim bOK As Boolean
    bOK = True
    Dim ptrByte As Long
    ptrByte = VarPtr(byteCh(0))
    Dim j As Long
    j = 0
    Dim str As String
    While bOK
        ' kopiranje na blok memorija
        CopyMemory ptrByte, strz + j, 1
        If (byteCh(0) = 0) Or (j = 255) Then
            bOK = False
        Else
            str = str + Chr(byteCh(0))
        End If
        j = j + 1
    Wend
    AnsiZtoString = str
End Function

Public Function WriteCommand(ByVal com_port As Byte, ByVal seq As Byte, ByVal sVlez As String, ByRef sIzlez As String, ByRef sStatus As String) As Integer
    
    sIzlez = String$(256, vbNullChar)   ' Alociranje 255 bajti memorija
    sStatus = String$(256, vbNullChar)  ' Alociranje 255 bajti memorija
    
    Dim lngPortHandle As Long           ' handle do otvoreniot COM port
    Dim lngIzlez As Long                ' pointer do izlezot
    Dim lngStatus As Long               ' pointer do statusot
    
    lngIzlez = StrPtr(sIzlez)
    lngStatus = StrPtr(sStatus)
    
    If ((seq < 32) Or (seq > 127)) Then
        sIzlez = "Sekventniot broj MORA da pripagja vo intervalot od [32,127]"
        sStatus = "GRESEN Sekventen broj"
        Exit Function
    End If
    
    If (Len(sVlez) = 0) Then
        sIzlez = "Parametarot vlez e cmd+data, znaci -mora- da bide min. 1 znak dolg"
        sStatus = "GRESNA vrednost za parametarot vlez"
        Exit Function
    End If
    
    lngPortHandle = InitComm(com_port)  ' Dokolku ima problem so otvaranje na portot, rezultatot ke bide 0
    If lngPortHandle > 0 Then
        If (SendCommand(lngPortHandle, seq, sVlez, lngIzlez, lngStatus) >= 0) Then 'prakjanje na komandata
            sIzlez = AnsiZtoString(lngIzlez)    ' konverzija na bafer->string
            sStatus = AnsiZtoString(lngStatus)  ' konverzija na bafer->string
        Else
            sIzlez = "Neuspeshno prakjanje/primanje na podatoci od COM" + str(com_port)
            sStatus = "GRESKA"
            ' Menuvanjeto na gornive poraki za greski ne vlijae vrz pravilnoto
            ' funkcioniranje na programot, smenete gi vo oblik sto
            ' najdobro vi odgovara (na pr. DialogBox ili sl.)
        End If
        CloseComm (lngPortHandle)           ' Zatvaranje na portot
    Else
        sIzlez = "Portot COM" + str(com_port) + " ne e pronajden!"
        sStatus = "GRESKA"
            ' Menuvanjeto na gornive poraki za greski ne vlijae vrz pravilnoto
            ' funkcioniranje na programot, smenete gi vo oblik sto
            ' najdobro vi odgovara (na pr. DialogBox ili sl.)
    End If
    
    ' Zabeleska:
    ' Kako sto moze da se zabelezi so izvrsuvanje na sekoja komanda WriteCommand
    ' povtorno se vrsi Otvaranje i Zatvaranje na seriskiot port (f-ciite InitComm i CloseComm)
    ' postapka sto moze vo nekoi slucai da dovede do blokiranje na seriskiot port
    ' zatoa najdobro e InitComm i CloseComm da gi izvlecete -nadvor- od ovaa funkcija

End Function
