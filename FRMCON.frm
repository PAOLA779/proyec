VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMCON 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   17130
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "FRMCON.frx":0000
      Left            =   7560
      List            =   "FRMCON.frx":0010
      TabIndex        =   8
      Text            =   "AÑO"
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "FRMCON.frx":002C
      Left            =   7560
      List            =   "FRMCON.frx":0054
      TabIndex        =   7
      Text            =   "MES"
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "FRMCON.frx":0088
      Left            =   7560
      List            =   "FRMCON.frx":00E9
      TabIndex        =   6
      Text            =   "DIA"
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "FRMCON.frx":0160
      Left            =   6000
      List            =   "FRMCON.frx":0170
      TabIndex        =   5
      Text            =   "AÑO"
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FRMCON.frx":018C
      Left            =   6000
      List            =   "FRMCON.frx":01B4
      TabIndex        =   4
      Text            =   "MES"
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FRMCON.frx":01E8
      Left            =   6000
      List            =   "FRMCON.frx":0249
      TabIndex        =   3
      Text            =   "DIA"
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FRMCON.frx":02C0
      Left            =   6360
      List            =   "FRMCON.frx":02CD
      TabIndex        =   2
      Top             =   5880
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3735
      Left            =   7560
      TabIndex        =   1
      Top             =   1920
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   15000
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   11520
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   3240
      Picture         =   "FRMCON.frx":02EC
      Top             =   8160
      Width           =   1650
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   5280
      Picture         =   "FRMCON.frx":0FDF
      Top             =   8280
      Width           =   1800
   End
   Begin VB.Image Image4 
      Height          =   1860
      Left            =   840
      Picture         =   "FRMCON.frx":301A
      Top             =   0
      Width           =   3825
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   7560
      Picture         =   "FRMCON.frx":5339
      Top             =   8280
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   16395
      Left            =   -480
      Picture         =   "FRMCON.frx":71A7
      Top             =   0
      Width           =   17625
   End
End
Attribute VB_Name = "FRMCON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer
Dim s, a As Integer


Private Sub Combo1_Click()
If Combo1.Text = "Desde" Then
        Combo2.Visible = True
        Combo3.Visible = True
        Combo4.Visible = True
        Combo5.Visible = False
        Combo6.Visible = False
        Combo7.Visible = False
    End If
    If Combo1.Text = "Hasta" Then
        Combo5.Visible = True
        Combo6.Visible = True
        Combo7.Visible = True
        Combo2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    End If
    If Combo1.Text = "Desde/Hasta" Then
        Combo2.Visible = True
        Combo3.Visible = True
        Combo4.Visible = True
        Combo5.Visible = True
        Combo6.Visible = True
        Combo7.Visible = True
        
    End If
End Sub

Private Sub Command1_Click()
FRMVENTAS_ELI.Show
Unload Me
End Sub

Private Sub DataGrid1_Click()

  Label2 = DataGrid1.Columns(0).Text
    With rsFactura
        Dim s As String
        s = Label2.Caption
        If .State = 1 Then .Close
        .Open "Select * From FACTURA Where [IDVENTAS] Like '" & s & "'"
        Set DataGrid2.DataSource = rsFactura
    End With
End Sub


Private Sub Form_Load()
FACTURA_ELIMINADA
VENTAS_ELIMINADAS
tablaVENTAS

factura
Set DataGrid1.DataSource = RSVEN
RSVEN.MoveFirst

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\img\bus1.jpg")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\img\bus1.jpg")

Dim s, a As String
    With RSVEN
If .State = 1 Then .Close
        If Combo1.Text = "Desde" Then
            s = "#" & Combo3.Text & "/" & Combo2.Text & "/" & Combo4.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA) >= " & s & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
        If Combo1.Text = "Hasta" Then
        a = "#" & Combo3.Text & "/" & Combo2.Text & "/" & Combo4.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA)<= " & a & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
        If Combo1.Text = "Desde/Hasta" Then
            If .State = 1 Then .Close
            s = "#" & desde.Text & "#"
            a = "#" & Hasta.Text & "#"
            .Open "Select * From VENTAS Where FECHA >= " & s & ") AND ((VENTAS.[FECHA])<= " & a, CN, adOpenStatic, adLockBatchOptimistic
            Set DataGrid1.DataSource = RSVEN
        End If
End With
Set DataGrid1.DataSource = RSVEN
End Sub




Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img\VEN_EL.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img\VEN_EL.jpg")
FRMVENTAS_ELI.Show
Unload Me
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\img\ELI0.jpg")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\img\ELI0.jpg")
With RSVENTAS_ELIMINADAS
    .AddNew
    !IDVENTAS = DataGrid1.Columns(0).Text
    !FECHA = DataGrid1.Columns(1).Text
    !CEDULACLIENTE = DataGrid1.Columns(2).Text
    !CEDULADUENO = DataGrid1.Columns(3).Text
 
    End With
    q = rsFactura.RecordCount
    For X = 1 To q
    With RSFACTURA_ELIMINADAS
    If .RecordCount = 0 Then Exit Sub
    !IDFACTURA = DataGrid2.Columns(0).Text
    !IDPRODUCTO = DataGrid2.Columns(1).Text
    !CANTIDAD = DataGrid2.Columns(2).Text
    !PRECIO = DataGrid2.Columns(3).Text
    !IDVENTAS = DataGrid2.Columns(4).Text
    
    End With
    rsFactura.MoveNext
    Next
    With RSVEN
    .Delete
    .MoveFirst
    End With
    For X = 1 To q
    With rsFactura
 
    .Requery
    .Delete
    .MoveNext
    If .EOF Then Exit Sub
    End With
    Next
End Sub
