VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMINV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INVENTARIO - Bazar Jessica"
   ClientHeight    =   12525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19875
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   19875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BUSCAR"
      Height          =   495
      Left            =   15240
      TabIndex        =   15
      Top             =   8880
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dgpro 
      Height          =   3495
      Left            =   11400
      TabIndex        =   13
      Top             =   5280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
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
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   13080
      TabIndex        =   12
      Top             =   3840
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   13080
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox TXTNUMP 
      DataField       =   "NOMBRE"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox TXTCAN 
      DataField       =   "CANTIDAD"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TXTCOS 
      DataField       =   "PRECIO"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   6015
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin MSAdodcLib.Adodc ADODCINV 
      Height          =   615
      Left            =   11520
      Top             =   12000
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BASEINV.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BASEINV.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "INVENTARIO"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TXTIDPRO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   495
      Left            =   13320
      TabIndex        =   14
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Image Image19 
      Height          =   1185
      Left            =   0
      Picture         =   "Form1.frx":0017
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
   End
   Begin VB.Image ImageM 
      Height          =   675
      Left            =   2400
      Picture         =   "Form1.frx":18BF
      Top             =   11760
      Width           =   1725
   End
   Begin VB.Image Image18 
      Height          =   690
      Left            =   11640
      Picture         =   "Form1.frx":2135
      Stretch         =   -1  'True
      Top             =   9600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image Image17 
      Height          =   630
      Left            =   9960
      Picture         =   "Form1.frx":3C00
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image16 
      Height          =   690
      Left            =   6960
      Picture         =   "Form1.frx":4692
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image15 
      Height          =   690
      Left            =   9720
      Picture         =   "Form1.frx":56A8
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   8760
      Picture         =   "Form1.frx":66C6
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image13 
      Height          =   690
      Left            =   7920
      Picture         =   "Form1.frx":7446
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image12 
      Height          =   1185
      Left            =   9000
      Picture         =   "Form1.frx":81BE
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Image Image11 
      Height          =   750
      Left            =   10680
      Picture         =   "Form1.frx":9A66
      Top             =   10320
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image10 
      Height          =   750
      Left            =   360
      Picture         =   "Form1.frx":BDC2
      Top             =   11760
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Top             =   11400
      Width           =   10695
   End
   Begin VB.Image Image9 
      Height          =   1575
      Left            =   -120
      Picture         =   "Form1.frx":DB5A
      Stretch         =   -1  'True
      Top             =   11280
      Width           =   11145
   End
   Begin VB.Image Image8 
      Height          =   3495
      Left            =   8160
      Picture         =   "Form1.frx":DD97
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image Image7 
      Height          =   750
      Left            =   8760
      Picture         =   "Form1.frx":DFD4
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image6 
      Height          =   750
      Left            =   5880
      Picture         =   "Form1.frx":FE42
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   3960
      Picture         =   "Form1.frx":11E7D
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   2040
      Picture         =   "Form1.frx":13AFA
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   120
      Picture         =   "Form1.frx":15B37
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   4680
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE LOS PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IDPRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "FRMINV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer
Dim buscar As String
Dim CN As New ADODB.Connection
Private WithEvents RS As ADODB.Recordset
Attribute RS.VB_VarHelpID = -1
Dim C As Integer

Private Sub CMDBUSCAR_Click()
ADODCINV.Refresh
DataGrid1.Refresh
ADODCINV.Recordset.Find "idproducto=" & Val(TXTIDPRO.Text)

End Sub

Private Sub CMDMOD_Click()
    MsgBox "Llenar todos los campos de datos de los productos y guardar para agregar correctamente al inventario.", vbInformation, "Dialogo"
    RSINV.MoveLast
    RSINV.AddNew
    RSINV("NOMBRE") = TXTNUMP.Text
    RSINV("PRECIO") = TXTCOS.Text
    RSINV("CANTIDAD") = TXTCAN.Text
    RSINV("IDPROVEEDORES") = 0
    RSINV.Update
    ADODCINV.Refresh
    ADODCINV.Recordset.MoveLast
    
    
End Sub
'
Private Sub CMDSAL_Click()
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
            End
    End If
    ADODCINV.Recordset.MoveFirst
'BUTONES DE MOVIMIENTO INICIO
End Sub



Private Sub Command2_Click()

ADODCINV.Recordset.MovePrevious

End Sub

Private Sub Command3_Click()

ADODCINV.Recordset.MoveNext

End Sub

Private Sub Command4_Click()

ADODCINV.Recordset.MoveLast

End Sub
'BUTONES DE MOVIMIENTO END


Private Sub Combo1_Change()
    If Combo1.Text <> "" Then
        Text1.Enabled = True
    End If
    
End Sub


Private Sub Command1_Click()

DataGrid1 = "select * from INVENTARIOS where IDPROVEEDORES Like '" & Text1.Text & "'"
DataGrid1.Refresh

End Sub

Private Sub DataGrid1_DblClick()
If DataGrid1.ApproxCount < 1 Then

MsgBox "no ha seleccionado ningun registro", vbExclamation
Exit Sub
Else

    
      TXTIDPRO.Text = DataGrid1.Columns(0).Text
     TXTNUMP.Text = DataGrid1.Columns(1).Text
     TXTCAN.Text = DataGrid1.Columns(2).Text
     TXTCOS.Text = DataGrid1.Columns(3).Text
     
     'TXTIDPROV.Text = DataGrid1.Columns(4).Text

End If
End Sub

Private Sub dgpro_Click()
    C = RSNOM!idproveedor
    Label1.Caption = C
End Sub

Private Sub Form_Load()
Set RS = New ADODB.Recordset

If CN.State = 0 Then CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & " Data Source= " & App.Path & "\DATA\BASEINV.mdb;Persist Security Info=False"
    '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAOLA\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
    
    Combo1.AddItem "PROVEEDORES"
    Combo1.AddItem "NOMBRE"
    
    tablaPROVEEDORES
    TABLANOMBRE
    Text1.Enabled = True
    
    ADODCINV.Recordset.MoveLast
    
    DataGrid1.Enabled = False
    
    
    FRMINV.Picture = LoadPicture(App.Path & "\IMG\tst.jpg")
    Image1.Picture = LoadPicture(App.Path & "\IMG\logob.gif")
    Image2.Picture = LoadPicture(App.Path & "\IMG\logoinv.gif")
    DataGrid1.Columns(1).Width = 5600
    
    Image4.Picture = LoadPicture(App.Path & "\IMG\guad2.jpg")
    'Image5.Picture = LoadPicture(App.Path & "\IMG\ed2.jpg")
    B = 0

    

    
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image18.Picture = LoadPicture(App.Path & "\img0\pro1.jpg")
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image18.Picture = LoadPicture(App.Path & "\img0\pro0.jpg")
    FRMNUELO.Show
    Unload Me
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\img\agr1.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\img\agr0.jpg")
    If MsgBox("Estas seguro que deseas guardar ", vbQuestion + vbYesNo, "Inventario") = vbYes Then
    RSINV.MoveLast
    RSINV.AddNew
    
    RSINV("NOMBRE") = TXTNUMP.Text
    RSINV("PRECIO") = TXTCOS.Text
    RSINV("CANTIDAD") = TXTCAN.Text
    RSINV("IDPROVEEDORES") = 0
    RSINV.Update
    ADODCINV.Refresh
    ADODCINV.Recordset.MoveLast
    
    
    Image4.Picture = LoadPicture(App.Path & "\IMG\gua0.jpg")
    'Image5.Picture = LoadPicture(App.Path & "\IMG\ed0.jpg")
    Else
    MsgBox "No se ha guardado el registro.", vbInformation, "Dialogo"
    End If
    
End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\img\gua1.jpg")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If B = 1 Then
    If TXTNUMP.Text = "" Or TXTCAN.Text = "" Or TXTCOS.Text = "" Then
    Image4.Picture = LoadPicture(App.Path & "\img\gua0.jpg")
    MsgBox "Llenar todos los campos de datos de los productos", vbInformation, "Dialogo"
    Exit Sub
    
    Else
    
    ADODCINV.Recordset.Fields("NOMBRE") = TXTNUMP.Text
    ADODCINV.Recordset.Fields("CANTIDAD") = TXTCAN.Text
    ADODCINV.Recordset.Fields("PRECIO") = TXTCOS.Text
    ADODCINV.Recordset.Update
    MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
    B = 0
    DataGrid1.Enabled = True
    Image13.Enabled = True
    Image14.Enabled = True
    Image15.Enabled = True
    Image16.Enabled = True
    Image6.Enabled = True
    Image7.Enabled = True
    End If
    Else
    Image4.Picture = LoadPicture(App.Path & "\IMG\guad2.jpg")
    
    End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\img\ed1.jpg")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image5.Picture = LoadPicture(App.Path & "\img\ed0.jpg")
     If MsgBox("Esta seguro que desea editar un registro?", vbQuestion + vbYesNo) = vbYes Then
     
    DataGrid1.Enabled = False
    Image13.Enabled = False
    Image14.Enabled = False
    Image15.Enabled = False
    Image16.Enabled = False
    Image6.Enabled = False
    Image7.Enabled = False
    Image4.Picture = LoadPicture(App.Path & "\img\gua0.jpg")
    B = 1
     
    End If

End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli1.jpg")
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli0.jpg")
    If MsgBox("Esta seguro que desea eliminar un registro?", vbQuestion + vbYesNo) = vbYes Then
        ADODCINV.Recordset.Delete
    End If
End Sub


Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture(App.Path & "\img\bus1.jpg")
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture(App.Path & "\img\bus0.jpg")
    
    ADODCINV.Refresh
    DataGrid1.Refresh
    ADODCINV.Recordset.Find "idproducto=" & Val(TXTIDPRO.Text)
End Sub


Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.Picture = LoadPicture(App.Path & "\img\ven1.jpg")
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.Picture = LoadPicture(App.Path & "\img\ven0.jpg")
    FRMVENTAS.Show
    Unload Me
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.Picture = LoadPicture(App.Path & "\img\da1.jpg")
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.Picture = LoadPicture(App.Path & "\img\da0.jpg")
    Set RS = CN.Execute("select *from inventario")
    If RS.EOF = False Then
    Set DRINV.DataSource = RS
    DRINV.Show
End If
End Sub
Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri1.jpg")
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri0.jpg")
    ADODCINV.Recordset.MovePrevious
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig1.jpg")
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig0.jpg")
    ADODCINV.Recordset.MoveNext
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi1.jpg")
End Sub

Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi0.jpg")
    ADODCINV.Recordset.MoveLast
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in1.jpg")
End Sub

Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in0.jpg")
    ADODCINV.Recordset.MoveFirst
End Sub
Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image17.Picture = LoadPicture(App.Path & "\img\X1.jpg")
End Sub

Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image17.Picture = LoadPicture(App.Path & "\img\X0.jpg")
    If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo, "Inventario") = vbYes Then
        
            End
    End If
    ADODCINV.Recordset.MoveFirst
End Sub
Private Sub ImageM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu1.jpg")
End Sub

Private Sub ImageM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu0.jpg")
If MsgBox("Esta seguro que desea regresar al menu?", vbQuestion + vbYesNo, "Inventario") = vbYes Then
    FRMMENU.Show
    
    
    Unload Me
    
    End If
End Sub

Private Sub Label1_Click()
    Set dgpro.DataSource = RSNOM
    Find idproveedores
    
    dgpro = "select * from INVENTARIOS where IDPROVEEDORES Like '" & Text1.Text & "'"
End Sub


Private Sub Text1_Change()
   buscar = Text1.Text
    If Combo1.Text = "IDPROVEEDOR" Then buscarIDPROVEEDOR
    If Combo1.Text = "NOMBRE" Then buscarNOMBRE
    Set dgpro.DataSource = RSNOM

    
End Sub
Sub BUSCARINV()
If RSINV.State = 1 Then RSINV.Close
    RSINV.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RSINV.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    RSINV.Open = "select * from INVENTARIO where IDPROVEEDORES like '" & Text1.Text & "'"
End Sub

Sub buscarIDPROVEEDOR()

    If RSPRO.State = 1 Then RSPRO.Close
    RSPRO.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RSPRO.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RSPRO.Open "select * from PROVEEDORES  where IDPROVEEDO =  BUSCAR  , CN"
    
End Sub

Sub buscarNOMBRE()
    If RSNOM.State = 1 Then RSNOM.Close
    RSNOM.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RSNOM.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.

    RSNOM.Open "select * from PROVEEDORES  where NOMBRE like '%" & buscar & "%'", CN
End Sub


