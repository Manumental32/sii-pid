VERSION 5.00
Begin VB.Form FormPid 
   BackColor       =   &H8000000B&
   Caption         =   "Trabajo Final SII / Ejercicio 1"
   ClientHeight    =   8715
   ClientLeft      =   3045
   ClientTop       =   2415
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10485
   Begin VB.TextBox zoomMax 
      Height          =   375
      Left            =   3480
      TabIndex        =   48
      Text            =   "600"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox zoomMin 
      Height          =   285
      Left            =   840
      TabIndex        =   47
      Text            =   "0"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox PorcSalida 
      DataSource      =   "PorcSalida"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   45
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox ValorSP 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   3360
      TabIndex        =   44
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox GraficoPV 
      BackColor       =   &H00000000&
      Height          =   3165
      Left            =   600
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   35
      Top             =   1560
      Width           =   4245
   End
   Begin VB.PictureBox GraficoSalida 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   3165
      Left            =   5880
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   34
      Top             =   1560
      Width           =   4125
   End
   Begin VB.Timer Control 
      Interval        =   50
      Left            =   9360
      Top             =   360
   End
   Begin VB.Timer Simu 
      Interval        =   10
      Left            =   8640
      Top             =   360
   End
   Begin VB.CommandButton CargarSP 
      Caption         =   "CargarSP"
      Height          =   375
      Left            =   12120
      TabIndex        =   33
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox tiempolazo 
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
      Height          =   400
      Left            =   960
      TabIndex        =   21
      Text            =   "500"
      Top             =   7800
      Width           =   915
   End
   Begin VB.HScrollBar BarraSP 
      Height          =   255
      Left            =   3360
      Max             =   400
      TabIndex        =   19
      Top             =   5880
      Value           =   400
      Width           =   1185
   End
   Begin VB.HScrollBar BarraSalida 
      Height          =   255
      Left            =   3240
      Max             =   2000
      TabIndex        =   18
      Top             =   7080
      Width           =   1185
   End
   Begin VB.CommandButton BtnAuto 
      BackColor       =   &H8000000D&
      Caption         =   "AUTOMÁTICO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   360
      Width           =   1740
   End
   Begin VB.CommandButton BtnManual 
      BackColor       =   &H8000000B&
      Caption         =   "MANUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   420
      Width           =   1740
   End
   Begin VB.TextBox ValorDeriv 
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
      Height          =   400
      Left            =   960
      TabIndex        =   6
      Text            =   "0"
      Top             =   6960
      Width           =   915
   End
   Begin VB.TextBox ValorInt 
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
      Height          =   400
      Left            =   840
      TabIndex        =   5
      Text            =   "10"
      Top             =   6120
      Width           =   915
   End
   Begin VB.TextBox ValorProp 
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
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Text            =   "0.8"
      Top             =   5280
      Width           =   915
   End
   Begin VB.HScrollBar BarraEntrada 
      Height          =   255
      Left            =   11760
      Max             =   100
      TabIndex        =   1
      Top             =   7440
      Width           =   1185
   End
   Begin VB.Label Label19 
      Caption         =   "zoomMax"
      Height          =   255
      Left            =   3480
      TabIndex        =   50
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Zoom Min"
      Height          =   255
      Left            =   840
      TabIndex        =   49
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Ingresar valores decimales solo con (punto) '.' Ejemplo: 0.1"
      Height          =   615
      Left            =   5400
      TabIndex        =   46
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   43
      Top             =   4440
      Width           =   285
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   42
      Top             =   1605
      Width           =   435
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "OP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   41
      Top             =   4680
      Width           =   4155
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "SP - PV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   40
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   39
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   38
      Top             =   4440
      Width           =   285
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "600"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   37
      Top             =   1605
      Width           =   525
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   36
      Top             =   3000
      Width           =   525
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3240
      TabIndex        =   32
      Top             =   7320
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4080
      TabIndex        =   31
      Top             =   7320
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400 l. max."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3480
      TabIndex        =   30
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4080
      TabIndex        =   29
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3360
      TabIndex        =   28
      Top             =   6120
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   12600
      TabIndex        =   27
      Top             =   7680
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   11760
      TabIndex        =   26
      Top             =   7680
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000 l/min."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3375
      TabIndex        =   25
      Top             =   6360
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1500 l/min."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   11880
      TabIndex        =   24
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label TxtSalida 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13440
      TabIndex        =   23
      Top             =   5640
      Width           =   1185
   End
   Begin VB.Label ValorPV 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   22
      Top             =   7680
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   20
      Top             =   7440
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALIDA (n)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11730
      TabIndex        =   17
      Top             =   5685
      Width           =   1755
   End
   Begin VB.Label suministrocap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   195
      Left            =   9960
      TabIndex        =   16
      Top             =   240
      Width           =   525
   End
   Begin VB.Label ValorError 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      TabIndex        =   15
      Top             =   7440
      Width           =   915
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6240
      TabIndex        =   14
      Top             =   7080
      Width           =   165
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2880
      TabIndex        =   13
      Top             =   5400
      Width           =   435
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2760
      TabIndex        =   12
      Top             =   7680
      Width           =   435
   End
   Begin VB.Shape IndAuto 
      FillColor       =   &H0000C000&
      Height          =   285
      Left            =   1110
      Shape           =   5  'Rounded Square
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape IndManual 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   3630
      Shape           =   5  'Rounded Square
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      Top             =   6600
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   8
      Top             =   5760
      Width           =   75
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1320
      TabIndex        =   7
      Top             =   4920
      Width           =   435
   End
   Begin VB.Label PorcEntrada 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11760
      TabIndex        =   3
      Top             =   6960
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2745
      TabIndex        =   2
      Top             =   6600
      Width           =   435
   End
   Begin VB.Label PorcSalidaLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11280
      TabIndex        =   0
      Top             =   6960
      Width           =   285
   End
End
Attribute VB_Name = "FormPid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim entradavalv, salidavalv, error, error1, error2, kp, ki, kd, salida, salida1, pv, incpro, incint, incder As Double
Dim modo, suministro, x, y, tlazo, n, a, b As Integer
Dim grafsalida(1000), grafpv(1000), tiempo, sp, deltaT, PvPunto As Long

Private Sub BtnManual_Click()
IndManual.FillStyle = 0
IndAuto.FillStyle = 1
modo = 0
BarraSalida.Enabled = True
PorcSalida.Enabled = True
End Sub

Private Sub BtnAuto_Click()
IndManual.FillStyle = 1
IndAuto.FillStyle = 0
modo = 1
BarraSalida.Enabled = False
PorcSalida.Enabled = False
End Sub


Private Sub CargarSP_Click()

    Simu_Timer

  If modo = 1 Then
    pidloop
  Else
    calcerror
    graficar
  End If

End Sub


Private Sub Form_load()
FormPid.Left = (Screen.Width / 2) - (FormPid.Width / 2)
FormPid.Top = (Screen.Height / 2) - (FormPid.Height / 2)

tiempo = 6000
error = 0
error1 = 0
error2 = 0
salida1 = 0
BarraSP.Value = 0
ValorSP = BarraSP.Value

pv = 0

'variables para calculo de la planta
n = 0
a = -5
b = 1
deltaT = 0.015
PvPunto = 0

zoomMin = 0
zoomMax = 600


suministro = 1000
suministrocap.Caption = suministro
BarraEntrada.Value = 100
PorcEntrada.Caption = BarraEntrada.Value
entradavalv = (BarraEntrada.Value * (suministro / 100)) / tiempo
salidavalv = 0
PorcSalida = 0
ValorPV.Caption = Conversion.Int(pv)

IndManual.FillStyle = 1
IndAuto.FillStyle = 0
modo = 1
BarraSalida.Enabled = False

kp = ValorProp.Text
ki = ValorInt.Text
kd = ValorDeriv.Text

'BarraSP.Value = 1000
BarraSP.Value = 300

ValorSP = BarraSP.Value
sp = BarraSP.Value

GraficoPV.Cls
GraficoPV.ScaleMode = 3
GraficoPV.ScaleHeight = 600
GraficoPV.ScaleWidth = 1000
GraficoPV.AutoRedraw = True
GraficoPV.ForeColor = vbRed
GraficoPV.DrawStyle = 0
GraficoPV.DrawWidth = 2

GraficoSalida.Cls
GraficoSalida.ScaleMode = 3
GraficoSalida.ScaleHeight = 2200
GraficoSalida.ScaleWidth = 1000
GraficoSalida.AutoRedraw = True
GraficoSalida.ForeColor = vbBlue
GraficoSalida.DrawStyle = 0
GraficoSalida.DrawWidth = 2
End Sub


Private Sub BarraEntrada_Change()
    PorcEntrada.Caption = BarraEntrada.Value
End Sub

Private Sub BarraSalida_Change()
    PorcSalida = BarraSalida.Value
End Sub


Private Sub PorcSalida_Change()
    If PorcSalida = "" Then PorcSalida = 0
    If PorcSalida > 2000 Then PorcSalida = 2000
    If PorcSalida < 0 Then PorcSalida = 0
    BarraSalida.Value = PorcSalida
End Sub

Private Sub Simu_Timer()

entradavalv = (BarraEntrada.Value * (suministro / 100)) / tiempo

salidavalv = (BarraSalida.Value) / tiempo


If pv < 0 Then pv = 0

ValorPV.Caption = Conversion.Int(pv)

suministrocap.Caption = suministro

End Sub

Private Sub BarraSP_Change()
ValorSP = BarraSP.Value
sp = BarraSP.Value
End Sub

Private Sub pidloop()
On Error Resume Next

   If ValorProp = "" Then ValorProp = 0
   If ValorProp < 0 Then
   ValorProp = 0
   End If
      If ValorProp > 10000 Then
      ValorProp = 10000
      End If
kp = Val(ValorProp)
    
   If ValorInt = "" Then ValorInt = 0
   If ValorInt < 0 Then
   ValorInt = 0
   End If
      If ValorInt > 10000 Then
      ValorInt = 10000
      End If
ki = Val(ValorInt)

   If ValorDeriv = "" Then ValorDeriv = 0
   If ValorDeriv < 0 Then
   ValorDeriv = 0
   End If
      If ValorDeriv > 10000 Then
      ValorDeriv = 10000
      End If
kd = Val(ValorDeriv)

   If tiempolazo < 300 Then
   tiempolazo = 300
   End If
      If tiempolazo > 1000 Then
      tiempolazo = 1000
      End If
tlazo = Val(tiempolazo)

error2 = error1
error1 = error
calcerror

incpro = (error - error1)
incint = (error + error1) / (2 * tlazo)
incder = (error - 2 * error1 + error2) / tlazo

salida = kp * (incpro + ki * incint + kd * incder) + salida1
salida1 = salida

TxtSalida.Caption = Round(salida, 2)


If salida > 2000 Then
  salida = 2000
End If
If salida < 0 Then
  salida = 0
End If


If salida1 > 2000 Then
  salida1 = 2000
End If
If salida1 < 0 Then
  salida1 = 0
End If


BarraSalida.Value = salida
'PorcSalida = Round(2000 - salida, 2)
PorcSalida = BarraSalida.Value

PvPunto = a * pv + b * PorcSalida
pv = pv + PvPunto * deltaT

graficar
End Sub

Private Sub Control_Timer()
  If modo = 1 Then
    pidloop
  Else
    calcerror
    graficar
  End If
End Sub

Private Sub graficar()
GraficoPV.Cls
grafpv(1000) = pv
For x = 0 To 999
    grafpv(x) = grafpv(x + 1)
    GraficoPV.PSet (x, zoomMax - (grafpv(x)))
Next x
GraficoPV.Line (0, zoomMax - sp)-(1000, zoomMax - sp), vbYellow

GraficoSalida.Cls

grafsalida(1000) = PorcSalida
For x = 0 To 999
    grafsalida(x) = grafsalida(x + 1)
    GraficoSalida.PSet (x, 2200 - (grafsalida(x)))
Next x
End Sub

Private Sub calcerror()
error = sp - pv
ValorError.Caption = Conversion.Int(error)
End Sub


Private Sub ValorSP_Change()
    If ValorSP = "" Then
        ValorSP = 0
    End If
    
    If ValorSP > 400 Then
        ValorSP = 400
    End If
    
    BarraSP.Value = ValorSP
End Sub

