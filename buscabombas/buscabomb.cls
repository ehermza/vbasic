VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Hecho por Ezequiel Hernandez 21-Oct 2006
'Basandome en winmine.vbp de S.Ramani

Option Explicit

Private Const SIZE_TECLA As Integer = 225
Private Const MAX_BOMBAS As Integer = 99

Private Const MAX_FILAS As Integer = 25
Private Const MIN_FILAS As Integer = 6
Private Const MAX_COLUM As Integer = 30
Private Const MIN_COLUM As Integer = 8

Private Const CLICK_IZQ As Byte = 1

Private Const NONE As Byte = 0
Private Const BAND As Byte = 13
Private Const CHECK As Byte = 16
Private Const BOMBA As Byte = 36

Private btnMouseDown As Boolean
Private Jugando As Boolean
Private Comenzo As Boolean

Private maxFilas As Integer
Private maxColum As Integer
Private cantBombas As Integer
Private cantDescub As Integer

Private tclBand As New Collection
Private tclBomba As New Collection
Private idBandera() As Integer

'Cada tecla se diferencia de otro en sus Coord.
'Asique para referenciar alguna tcl creo
'2 matrices que me digan:
Private tclConten() As Byte  ' que contiene (bomba, pista, null)
Private tclEstado() As Byte  ' estado actual (bandera, visto)

Private frmBusca As Form
Private objSize As frmConfig
Private partido As clsBusca

Property Set Init_Size(objPanel As frmConfig)

    Set objSize = objPanel
End Property
Property Set Initialize(frmBuscaBomb As Form)

    Set frmBusca = frmBuscaBomb
    objSize.txtBombas = 5
    objSize.txtFilas = 10
    objSize.txtColum = 10
    JuegoNuevo
End Property
Property Set frmInstancia(frmBuscaBomb As Form)
    
    Set frmBusca = frmBuscaBomb
    JuegoNuevo
End Property

Public Sub JuegoNuevo()

    maxFilas = Val(objSize.txtFilas)
    maxColum = Val(objSize.txtColum)
    cantBombas = Val(objSize.txtBombas)
    frmCambiarTamano
        
    frmBusca.Cls
    
    Jugando = True
    Comenzo = False
    frmBusca.FontBold = True
    
    Erase tclConten()
    Erase tclEstado()
    Erase idBandera()
'Ahora sí puedo completar las declaración de matrices hechas al comienzo

    ReDim tclEstado(maxColum - 1, maxFilas - 1)
    ReDim tclConten(maxColum - 1, maxFilas - 1)
    ReDim idBandera(maxColum - 1, maxFilas - 1)

End Sub
Public Sub frmMostrarPanel()

Dim Filas As Integer, Colum As Integer, Bombas As Integer
Dim permBom As Boolean, permMin As Boolean, permMax As Boolean

    frmBusca.Cls

VOLVER:
    objSize.Show 1

    Bombas = Val(objSize.txtBombas)
    Filas = Val(objSize.txtFilas)
    Colum = Val(objSize.txtColum)
    
    'Chequeando que los valores ingresados
    'sean los permitidos
    permBom = Bombas > MAX_BOMBAS
    permMax = Filas > MAX_FILAS Or Colum > MAX_COLUM
    permMin = Colum < MIN_COLUM Or Filas < MIN_FILAS
    
    Dim prohibido As Boolean
    prohibido = permBom Or permMax Or permMin

    If prohibido = True Then
        MsgBox "Valores NO permitidos", vbCritical, "Datos mal ingresados"
        GoTo VOLVER
    Else
        Jugando = False
    End If
End Sub

Private Sub frmCambiarTamano()
    
    frmBusca.Height = (SIZE_TECLA * maxFilas) + 700
     frmBusca.Width = (SIZE_TECLA * maxColum) + 120

End Sub

Private Sub ColocarBombas(ByVal prohX As Integer, ByVal prohY As Integer)
Dim i As Integer
Dim tclClick As Boolean
Dim tclX, tclY As Integer
    
    For i = 1 To cantBombas
    
VOLVER: Randomize Timer
        tclX = Int(Rnd * maxColum)
        tclY = Int(Rnd * maxFilas)

        tclClick = (tclX = prohX) And (tclY = prohY)
        
        If tclConten(tclX, tclY) = BOMBA Or tclClick Then GoTo VOLVER
        
        tclConten(tclX, tclY) = BOMBA   'Pone Bomba en (tclX, tclY)
        
        Dim objCoord() As New clsCoord
        ReDim objCoord(cantBombas)
        objCoord(i).ptoX = tclX
        objCoord(i).ptoY = tclY
        
        tclBomba.Add objCoord(i)
        
        'COLOCANDO PISTAS (NUMEROS)
        If Not ExcedeLimite(tclX - 1, tclY - 1) Then SumarPista tclX - 1, tclY - 1
        If Not ExcedeLimite(tclX + 1, tclY - 1) Then SumarPista tclX + 1, tclY - 1
        If Not ExcedeLimite(tclX - 1, tclY + 1) Then SumarPista tclX - 1, tclY + 1
        If Not ExcedeLimite(tclX + 1, tclY + 1) Then SumarPista tclX + 1, tclY + 1
        
        If Not ExcedeLimite(tclX - 1, tclY) Then SumarPista tclX - 1, tclY
        If Not ExcedeLimite(tclX + 1, tclY) Then SumarPista tclX + 1, tclY
        
        If Not ExcedeLimite(tclX, tclY - 1) Then SumarPista tclX, tclY - 1
        If Not ExcedeLimite(tclX, tclY + 1) Then SumarPista tclX, tclY + 1
        
    Next i
    
End Sub

Private Sub SumarPista(ByVal tclX As Integer, ByVal tclY As Integer)

    If tclConten(tclX, tclY) <> BOMBA Then tclConten(tclX, tclY) = tclConten(tclX, tclY) + 1

End Sub

Public Sub MouseDown(Boton As Integer, X As Single, Y As Single)
Dim tclX, tclY As Integer
Dim coordX, coordY As Integer

    If Not Jugando Then Exit Sub
    If Not Comenzo Then Exit Sub

    '(X, Y) = Coord. Mouse en pixeles
    'Para saber cuál fue la tcl a la que hice referencia
    'divido (X / Tamaño tcl en pixeles) = Posic. horiz de tcl

    tclX = Int(X / SIZE_TECLA)
    tclY = Int(Y / SIZE_TECLA)
    
    coordX = tclX * SIZE_TECLA
    coordY = tclY * SIZE_TECLA
    
    If (Boton = CLICK_IZQ) Then
    
        If (tclEstado(tclX, tclY) <> NONE) Then Exit Sub
        
        btnMouseDown = True
        frmBusca.imgPress.Left = coordX
        frmBusca.imgPress.Top = coordY
        frmBusca.imgPress.Visible = True

    Else
        Select Case tclEstado(tclX, tclY)

        Case NONE:  'CLAVANDO BANDERA
            If cantBombas <= tclBand.Count Then Exit Sub
            
            Dim objBand As New clsCoord
            objBand.ptoX = tclX
            objBand.ptoY = tclY
            tclBand.Add objBand
            idBandera(tclX, tclY) = tclBand.Count
            
            frmBusca.PaintPicture frmBusca.imgBand, coordX, coordY
            tclEstado(tclX, tclY) = BAND

        Case BAND:  'QUITANDO BANDERA
            tclBand.Remove idBandera(tclX, tclY)
            frmBusca.PaintPicture frmBusca.imgBoton, coordX, coordY
            tclEstado(tclX, tclY) = NONE
        End Select

    End If
        
End Sub

Public Sub MouseUp(Boton As Integer, X As Single, Y As Single)
Dim igualBombas As Boolean
Dim igualDescub As Boolean
Dim tclX, tclY As Integer

    If Not Jugando Then Exit Sub
    btnMouseDown = False
    frmBusca.imgPress.Visible = False
    
    tclX = Int(X / SIZE_TECLA)
    tclY = Int(Y / SIZE_TECLA)
        
    If Not Boton = CLICK_IZQ Then Exit Sub
    
    If Not Comenzo Then
        Comenzo = True
        ColocarBombas tclX, tclY
    End If
    
    If ExcedeLimite(tclX, tclY) Then Exit Sub
    
    If tclEstado(tclX, tclY) = BAND Then Exit Sub
    
    If tclConten(tclX, tclY) = BOMBA Then
        EstallaBomba
        
    ElseIf (tclConten(tclX, tclY) = NONE) Then
        AbrirEspVacio tclX, tclY    'ABRE CAMINO (TECLAS VACIAS)
    
    Else
        MostrarConten tclX, tclY  'MUESTRA Pista
    End If
    
    igualDescub = ((maxFilas * maxColum - cantBombas) = cantDescub)
    igualBombas = (tclBand.Count = cantBombas)
    
    If (igualDescub And igualBombas) Then
        Set frmBusca = Nothing   'GANO EL JUEGO
        MsgBox "Gracias a Dios! Evitaste una explosion", vbInformation, "Ganador"
    End If
    
End Sub

Private Sub AbrirEspVacio(ByVal tclX As Integer, ByVal tclY As Integer)
'Verifica contenido en cada tcl vecina

    If Not ExcedeLimite(tclX - 1, tclY - 1) Then Verif_Conten tclX - 1, tclY - 1
    If Not ExcedeLimite(tclX + 1, tclY - 1) Then Verif_Conten tclX + 1, tclY - 1
    If Not ExcedeLimite(tclX - 1, tclY + 1) Then Verif_Conten tclX - 1, tclY + 1
    If Not ExcedeLimite(tclX + 1, tclY + 1) Then Verif_Conten tclX + 1, tclY + 1
    
    If Not ExcedeLimite(tclX - 1, tclY) Then Verif_Conten tclX - 1, tclY
    If Not ExcedeLimite(tclX + 1, tclY) Then Verif_Conten tclX + 1, tclY
    
    If Not ExcedeLimite(tclX, tclY - 1) Then Verif_Conten tclX, tclY - 1
    If Not ExcedeLimite(tclX, tclY + 1) Then Verif_Conten tclX, tclY + 1

End Sub

Private Sub Verif_Conten(ByVal tclX As Integer, ByVal tclY As Integer)
    
    If tclEstado(tclX, tclY) = CHECK Then Exit Sub  'YA ESTA DESCUB. NO ME INTERESA
    
    MostrarConten tclX, tclY    'DESCUBRIR
    
    If tclConten(tclX, tclY) = NONE Then AbrirEspVacio tclX, tclY   'SI HAY ESPACIO EN BLANCO
    
End Sub

Private Sub MostrarConten(ByVal tclX As Integer, ByVal tclY As Integer)
Dim coordX As Single
Dim coordY As Single
    
    If tclEstado(tclX, tclY) = BAND Then Exit Sub
    
    coordX = tclX * SIZE_TECLA
    coordY = tclY * SIZE_TECLA
    
    If tclConten(tclX, tclY) = NONE Then 'CASILLERO EN BLANCO
        frmBusca.PaintPicture frmBusca.imgCheck, coordX, coordY
    
    ElseIf tclConten(tclX, tclY) = BOMBA Then
        frmBusca.PaintPicture frmBusca.imgBomba, coordX, coordY
        
    Else   'IMPRIME VALOR DE LA PISTA
        frmBusca.PaintPicture frmBusca.imgCheck, coordX, coordY
        frmBusca.CurrentX = coordX          'CALIBRANDO PARA
        frmBusca.CurrentY = coordY + 22     'IMPRIMIR CIFRA
        frmBusca.ForeColor = QBColor(tclConten(tclX, tclY))
        frmBusca.Print tclConten(tclX, tclY)
    End If
    
    tclEstado(tclX, tclY) = CHECK
    cantDescub = cantDescub + 1

End Sub

Private Sub EstallaBomba()
Dim i As Integer
Dim tclX, tclY As Integer
Dim BandMal As Boolean
Dim SinBand As Boolean
        
    For i = 1 To tclBomba.Count
        tclX = tclBomba.Item(i).ptoX
        tclY = tclBomba.Item(i).ptoY
        
        'REVISANDO EN CADA BOMBA SI TENIA UNA BAND CLAVADA
        SinBand = tclEstado(tclX, tclY) = NONE And tclConten(tclX, tclY) = BOMBA
        If SinBand Then MostrarConten tclX, tclY   'NO HAY BAND => MOSTRAR BOMBA
    Next i
    
    For i = 1 To tclBand.Count
        tclX = tclBand(i).ptoX
        tclY = tclBand(i).ptoY
        
        'REVISANDO SI HAY BANDERAS MAL PUESTAS
        BandMal = tclEstado(tclX, tclY) = BAND And tclConten(tclX, tclY) <> BOMBA
        If BandMal Then frmBusca.PaintPicture frmBusca.imgBandErr, (tclX * SIZE_TECLA), (tclY * SIZE_TECLA)
    Next i
    
    MsgBox "Estalló todo..! Perdiste", vbCritical, "Boom..!"
    
    Jugando = False
End Sub

Public Sub MueveMouse(Boton As Integer, X As Single, Y As Single)
Dim tclX, tclY As Integer
    
    If Not Jugando Then Exit Sub
    If Not btnMouseDown Then Exit Sub

    tclX = Int(X / SIZE_TECLA) ' - 1
    tclY = Int(Y / SIZE_TECLA) ' - 1
    If Not ExcedeLimite(tclX, tclY) Then

        If btnMouseDown Then
            frmBusca.imgPress.Left = tclX * SIZE_TECLA
            frmBusca.imgPress.Top = tclY * SIZE_TECLA
            frmBusca.imgPress.Visible = True
        End If
    Else
        frmBusca.imgPress.Visible = False
    End If
End Sub

Private Property Get ExcedeLimite(ByVal tclX As Integer, ByVal tclY As Integer)
Dim Condic As Boolean
    'REVISA QUE LA COORD INDICADA EN ARGUM.
    'ESTA DENTRO DE LA MATRIZ DE TECLAS
    'SI NO ES ASI DEVUELVE TRUE
    
    Condic = tclX < 0
    Condic = Condic Or tclY < 0
    Condic = Condic Or tclX > maxColum - 1
    Condic = Condic Or tclY > maxFilas - 1
    
    ExcedeLimite = Condic
    
End Property

Public Sub frmDesinst()
    Set frmBusca = Nothing
    End
End Sub
