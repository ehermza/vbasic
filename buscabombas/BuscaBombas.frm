VERSION 5.00
Begin VB.Form frmJuego 
   AutoRedraw      =   -1  'True
   Caption         =   "BuscaBombas"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "BuscaBombas.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgCheck 
      Height          =   225
      Left            =   6840
      Picture         =   "BuscaBombas.frx":281A6
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   225
   End
   Begin VB.Image imgBandErr 
      Height          =   225
      Left            =   6840
      Picture         =   "BuscaBombas.frx":284B8
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   225
   End
   Begin VB.Image imgBomba 
      Height          =   240
      Left            =   6840
      Picture         =   "BuscaBombas.frx":289EA
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image imgBoton 
      Height          =   225
      Left            =   6840
      Picture         =   "BuscaBombas.frx":28C94
      Stretch         =   -1  'True
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgBand 
      Height          =   225
      Left            =   6840
      Picture         =   "BuscaBombas.frx":291C6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   225
   End
   Begin VB.Image imgPress 
      Height          =   225
      Left            =   6840
      Picture         =   "BuscaBombas.frx":296F8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Menu menuNuevo 
      Caption         =   "Nuevo"
   End
   Begin VB.Menu menuSize 
      Caption         =   "Tamano"
   End
   Begin VB.Menu salir 
      Caption         =   "Cerrar"
   End
End
Attribute VB_Name = "frmJuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private objPanel As New frmConfig
Private objBusca As clsBusca
Private Sub form_Load()

    Set objBusca = New clsBusca
    Set objBusca.Init_Size = objPanel
    Set objBusca.Initialize = Me
End Sub
Private Sub Form_MouseDown(Boton As Integer, Shift As Integer, X As Single, Y As Single)

    objBusca.MouseDown Boton, X, Y
End Sub
Private Sub Form_MouseUp(Boton As Integer, Shift As Integer, X As Single, Y As Single)

    objBusca.MouseUp Boton, X, Y
End Sub
Private Sub Form_MouseMove(Boton As Integer, Shift As Integer, X As Single, Y As Single)

        objBusca.MueveMouse Boton, X, Y
End Sub
Private Sub menuNuevo_Click()

    Set objBusca = New clsBusca
    Set objBusca.Init_Size = objPanel
    Set objBusca.frmInstancia = Me
End Sub
Private Sub menuSize_Click()

    objBusca.frmMostrarPanel ' frmConfig
    menuNuevo_Click
End Sub
Private Sub salir_Click()

    objBusca.frmDesinst
End Sub
