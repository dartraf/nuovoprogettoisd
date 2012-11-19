Attribute VB_Name = "modBackup"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Function CopiaFile(origine As String, destinazione As String, prbarra As Object) As Long
prbarra.max = 100
prbarra.min = 0
Const BUFSIZE = 1024 'grandezza del buffer
Static Buf$
Dim BTest!, FSize!
Dim Chunk%, F1%, F2%
LunghezzaFileDestinazione = 0
prbarra.Value = 0
Open origine For Binary As #1 ' Apre il file.
Flunghezza = LOF(1)  ' Ottiene la lunghezza del file.
Open destinazione For Binary As #2 ' Apre il file.
BTest = Flunghezza - LOF(2)
Do
If BTest < BUFSIZE Then
   Chunk = BTest
Else
   Chunk = BUFSIZE
End If
Buf = String(Chunk, " ")
Get 1, , Buf
Put 2, , Buf
BTest = Flunghezza - LOF(2)
prbarra.Value = (100 - Int(100 * BTest / Flunghezza)) 'avanzamento progressbar
Loop Until BTest = 0
Close 1 'chiude il file di origine
Close 2 'chiude il file di destinazione
End Function

