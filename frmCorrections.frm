VERSION 5.00
Begin VB.Form frmCorrections 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   4980
   ClientTop       =   4230
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.ListBox LstCorrections 
      Height          =   1815
      ItemData        =   "frmCorrections.frx":0000
      Left            =   120
      List            =   "frmCorrections.frx":0002
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "������"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "��������"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "frmCorrections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ���������, ���������� ������ frmSpell ��� �����������
' ���������, ������������ ���������� Word
Friend Sub Display(Corrections)
Dim Word
For Each Word In Corrections
LstCorrections.AddItem Word
Next Word
' �������� ������ �������
LstCorrections.Selected(0) = True
' ���������� �����
Show vbModal
End Sub
' ������ ����� ��������� ���������
Public Sub cmdReplace_Click()
Russia.Replac LstCorrections.List(LstCorrections.ListIndex)
Unload Me
End Sub
' ������ ���������
Private Sub cmdCancel_Click()
Unload Me
End Sub


