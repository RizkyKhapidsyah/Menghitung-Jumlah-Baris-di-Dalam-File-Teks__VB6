VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Jumlah Baris di Dalam File Teks"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function HitungBarisFileTeks(strFileName As String) As Long
On Error GoTo ErrHandler
  Dim fso As FileSystemObject
  Dim TextStream As TextStream
  Dim lngBaris As Long, sLine As String
  'Buat object dengan menggunakan FSO
  Set fso = CreateObject("Scripting.FileSystemObject")
  'Buka file dan tampung ke dalam TextStream
  Set TextStream = fso.OpenTextFile(strFileName)
  'Ulangi selama belum mencapai akhir baris
  '(akhir dari stream).
  Do While TextStream.AtEndOfStream = False
    'Baca setiap satu baris
    sLine = TextStream.ReadLine
    'Update counter baris
    lngBaris = lngBaris + 1
  Loop
  'Setelah selesai, tutup file
  TextStream.Close
  'Kembalikan jumlah baris yang diperoleh
  HitungBarisFileTeks = lngBaris
  Exit Function
ErrHandler:
  MsgBox Err.Number & " - " & _
         Err.Description, _
         vbExclamation, _
         "Error HitungBarisFileTeks"
End Function

Private Sub Command1_Click()
  MsgBox "Jumlah baris dalam file = " & _
         HitungBarisFileTeks(App.Path & _
         "\FileTeks.txt"), vbInformation, _
         "Jumlah Baris"
End Sub


