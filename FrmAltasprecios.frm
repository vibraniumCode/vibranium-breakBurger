VERSION 5.00
Begin VB.Form FrmAltasprecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas de precios"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   Icon            =   "FrmAltasprecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox CBPrecios 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   3
         Text            =   "Listado de precios"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox precio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "$00,00"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingresar precios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1545
      End
   End
End
Attribute VB_Name = "FRMAltasprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CBPrecios_KeyPress(KeyAscii As Integer)
If IsNumeric(Chr(KeyAscii)) _
  Or KeyAscii = 8 _
  Or KeyAscii = 32 _
  Or KeyAscii = 46 _
  Then
    KeyAscii = 0
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
sql = "select id_envios,precios from tp_envios"
 Call BuscaConexion(sql)
 Do While Not rs.EOF
     CBPrecios.AddItem "$" & rs.Fields("precios").Value
     CBPrecios.ItemData(CBPrecios.NewIndex) = rs.Fields("id_envios").Value
     rs.MoveNext
 Loop
 rs.Close
 cn.Close
 Set cn = Nothing
 Set rs = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Inicio.Enabled = True
Unload Me
End Sub



Private Sub precio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sql = "IF NOT EXISTS(SELECT id_envios FROM tp_envios WHERE precios = " & precio.Text & " ) insert into tp_envios select " & precio.Text
    Call BuscaConexion(sql)
    Set cn = Nothing
    Set rs = Nothing
    CBPrecios.Clear
    sql = "select id_envios,precios from tp_envios"
    Call BuscaConexion(sql)
    Do While Not rs.EOF
        CBPrecios.AddItem "$" & rs.Fields("precios").Value
        CBPrecios.ItemData(CBPrecios.NewIndex) = rs.Fields("id_envios").Value
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
    Set cn = Nothing
    Set rs = Nothing
End If
If IsNumeric(Chr(KeyAscii)) _
  Or KeyAscii = 8 _
  Or KeyAscii = 32 _
  Or KeyAscii = 46 _
  Then
  Else
    KeyAscii = 0
  End If
End Sub

