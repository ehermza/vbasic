VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Dificultad"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   2685
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opcDific 
      Caption         =   "Difícil"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtBombas 
      BackColor       =   &H80000004&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtFilas 
      BackColor       =   &H80000004&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtColum 
      BackColor       =   &H80000004&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton opcPers 
      Caption         =   "Personal"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton opcMedio 
      Caption         =   "Medio"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton opcFacil 
      Caption         =   "Fácil"
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Bomba"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Filas"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub opcDific_Click()
    txtBombas = 70
    txtColum = 20
    txtFilas = 23
End Sub

Private Sub opcFacil_Click()
    txtBombas = 8
    txtColum = 10
    txtFilas = 10
End Sub

Private Sub opcMedio_Click()
    txtBombas = 40
    txtColum = 15
    txtFilas = 18
End Sub

Private Sub txtFilas_Click()
    opcPers = True
End Sub

Private Sub txtColum_Click()
    opcPers = True
End Sub

Private Sub txtBombas_Click()
    opcPers = True
End Sub
