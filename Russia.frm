VERSION 5.00
Begin VB.Form Russia 
   Caption         =   "Электронный репетитор"
   ClientHeight    =   11010
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.TextBox txtSpell 
      Height          =   9135
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   14775
   End
   Begin VB.CommandButton cmdSpell 
      Caption         =   "Проверить"
      Height          =   495
      Left            =   13440
      TabIndex        =   0
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label Ochek 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   9960
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Введите текст для проверки:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Russia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mdocSpell As New Document
Dim mblnVisible As Boolean
Dim intUpper As Integer
Dim intCount As Integer
Dim intMistake As Integer
Dim intOnk As Integer
Dim strForward() As String
Dim strForward2() As String
Const KEY_F7 = 118
Const KEY_F2 = 113
Private Sub cmdSpell_Click()
Ochek.Caption = ""
' начальное состояние
 strForward2 = Split(txtSpell, " ")
' добавляем текст в объект Range текстового процессора Word
mdocSpell.Range.Text = txtSpell
' ВНИМАНИЕ: перед вызовом метода CheckSpelling
' надо выполнить следующие две операции:
' 1) сделать окно Word видимым
mdocSpell.Application.Visible = True
' 2) активизировать Word
AppActivate mdocSpell.Application.Caption
' проверяем правописание
mdocSpell.Range.CheckSpelling
'проверяем синтаксис
  mdocSpell.Range.CheckGrammar
' обновляем содержимое текстового поля
' в соответствии с тем, что получено от Word
txtSpell = mdocSpell.Range.Text
' отсекаем нулевой символ, добавленный Word
txtSpell = Left(txtSpell, Len(txtSpell) - 1)
' активизируем свою программу
AppActivate Caption
' последующее состояние
strForward = Split(txtSpell, " ")
For intCount = 0 To UBound(strForward)
' проверяем изменение
If strForward(intCount) <> strForward2(intCount) Then intMistake = intMistake + 1
 Next intCount
 'определяем оценку
 intOnk = Int(5 - intMistake / 2)
 Ochek.Caption = "Ваша оценка/ " & Str(intOnk)
End Sub
Private Sub Form_Load()
' переменная mblnVisible используется в процедуре Form_Unload,
' чтобы определить, запускала ли эта программа Word
mblnVisible = mdocSpell.Application.Visible
End Sub
' очистка
Private Sub Form_Unload(Cancel As Integer)
' проверяем, запускала ли эта программа Word
If mblnVisible Then
' закрываем документ
mdocSpell.Close savechanges:=False
Else
' закрываем Word
mdocSpell.Application.Quit savechanges:=False
End If
End Sub

