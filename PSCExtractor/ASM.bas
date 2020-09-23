Attribute VB_Name = "ASM"
' ASM BAS


Option Base 1  ' Arrays base 1
DefLng A-W     ' All variable Long
DefSng X-Z     ' unless singles
               ' unless otherwise defined
'-----------------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long
'-----------------------------------------------------------------
Public ptrMC, ptrStruc        ' Ptrs to Machine Code & Structure

'MCode Structure
Public Type PStruc
   picWd As Long
   picHt As Long
   Ptrpic1mem As Long      ' pic1mem(1-4,1-picWd,1-picHt) bytes
   Ptrpic2mem As Long      ' pic2mem(1-4,1-picWd,1-picHt) bytes
   StartColor As Long      ' RGB clicked
   PtrDELRGB As Long       ' DELRGB(0,1,2,3)  ' Grey,+/- R,G,B
End Type
Public DStruc As PStruc

Public ExtractMC() As Byte

Public Sub ASM_Extract()
res = CallWindowProc(ptrMC, ptrStruc, ptrMC, 3&, 4&)
'Stop
End Sub

Public Sub Loadmcode(InFile$, MCCode() As Byte)
'Load machine code into InCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox InFile$ & " missing", , "Extractor"
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox InFile$ & " missing", , "Extractor"
   DoEvents
   Unload Form1
   End
End If
ReDim MCCode(MCSize&)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

