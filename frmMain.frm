VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WXLite - using www.Weather.com"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWind 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5925
      Top             =   1950
   End
   Begin MSComctlLib.ImageList imlWX31 
      Left            =   6300
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   31
      ImageHeight     =   31
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1546
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1713
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A83
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2134
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2475
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2961
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B39
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3161
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":329A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3572
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3730
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":390C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C91
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":402D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4397
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":456C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4714
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5070
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5237
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5417
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5475
      Top             =   1950
   End
   Begin MSComctlLib.ImageList imlPressure 
      Left            =   5100
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   5
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6149
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61A0
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61F7
            Key             =   "steady"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlWX52 
      Left            =   5700
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   52
      ImageHeight     =   52
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   48
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6246
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6934
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7022
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7743
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8415
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9368
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9705
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A074
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A314
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A5FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B1F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B607
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA18
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BD64
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C223
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C508
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CA8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CD97
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D1F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D568
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D992
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DEA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E713
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ED6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F086
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F41F
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F7C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FBB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1045B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   7350
      TabIndex        =   3
      Top             =   75
      Width           =   765
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   6600
      TabIndex        =   2
      Top             =   75
      Width           =   765
   End
   Begin VB.Frame fraF3 
      Caption         =   "Next 36 Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   6450
      TabIndex        =   35
      Top             =   825
      Width           =   1665
      Begin VB.Label lblF3 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   3120
         Left            =   75
         TabIndex        =   38
         Top             =   225
         Width           =   1515
      End
      Begin VB.Image imgF3 
         Height          =   465
         Left            =   1125
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame fraF2 
      Caption         =   "Next 24 Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   4725
      TabIndex        =   34
      Top             =   825
      Width           =   1665
      Begin VB.Label lblF2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   3120
         Left            =   75
         TabIndex        =   37
         Top             =   225
         Width           =   1515
      End
      Begin VB.Image imgF2 
         Height          =   465
         Left            =   1125
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame fraF1 
      Caption         =   "Next 12 Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   3000
      TabIndex        =   33
      Top             =   825
      Width           =   1665
      Begin VB.Label lblF1 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   3120
         Left            =   75
         TabIndex        =   36
         Top             =   225
         Width           =   1515
      End
      Begin VB.Image imgF1 
         Height          =   465
         Left            =   1125
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.Frame fraCurrent 
      Caption         =   "Current Conditions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   0
      TabIndex        =   7
      Top             =   825
      Width           =   2940
      Begin VB.Label lblCurrent 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   75
         TabIndex        =   40
         Top             =   225
         Width           =   60
      End
      Begin VB.Label lblDew 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   39
         Top             =   1575
         Width           =   60
      End
      Begin VB.Label lblflTemp 
         AutoSize        =   -1  'True
         Caption         =   "0°F"
         Height          =   195
         Left            =   1275
         TabIndex        =   32
         Top             =   450
         Width           =   240
      End
      Begin VB.Label lblUdated 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   75
         TabIndex        =   31
         Top             =   3150
         Width           =   60
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "UV Index:"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   30
         Top             =   675
         Width           =   705
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Wind:"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   29
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lblUV 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   28
         Top             =   675
         Width           =   60
      End
      Begin VB.Label lblWind 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   27
         Top             =   900
         Width           =   60
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Humidity:"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   26
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label lblHumidity 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   25
         Top             =   1125
         Width           =   60
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Pressure:"
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   24
         Top             =   1350
         Width           =   660
      End
      Begin VB.Label lblPressure 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   23
         Top             =   1350
         Width           =   60
      End
      Begin VB.Image imgPressure 
         Height          =   120
         Left            =   1125
         Top             =   1380
         Width           =   75
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Dew Point:"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   22
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Visibility:"
         Height          =   195
         Index           =   5
         Left            =   75
         TabIndex        =   21
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblVisibility 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   20
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Last Update:"
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   19
         Top             =   2925
         Width           =   915
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Sunrise:"
         Height          =   195
         Index           =   7
         Left            =   75
         TabIndex        =   18
         Top             =   2250
         Width           =   570
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Sunset:"
         Height          =   195
         Index           =   8
         Left            =   75
         TabIndex        =   17
         Top             =   2475
         Width           =   540
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Today:"
         Height          =   195
         Index           =   9
         Left            =   1275
         TabIndex        =   16
         Top             =   2025
         Width           =   495
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Tomorrow:"
         Height          =   195
         Index           =   10
         Left            =   2025
         TabIndex        =   15
         Top             =   2025
         Width           =   750
      End
      Begin VB.Label lblSR0 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   14
         Top             =   2250
         Width           =   60
      End
      Begin VB.Label lblSS0 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   1275
         TabIndex        =   13
         Top             =   2475
         Width           =   60
      End
      Begin VB.Label lblSR1 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   2025
         TabIndex        =   12
         Top             =   2250
         Width           =   60
      End
      Begin VB.Label lblSS1 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   2025
         TabIndex        =   11
         Top             =   2475
         Width           =   60
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Feels Like:"
         Height          =   195
         Index           =   11
         Left            =   75
         TabIndex        =   10
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         Caption         =   "Daylight Remaining:"
         Height          =   195
         Index           =   12
         Left            =   75
         TabIndex        =   9
         Top             =   2700
         Width           =   1410
      End
      Begin VB.Label lblDaylight 
         AutoSize        =   -1  'True
         Caption         =   "*"
         Height          =   195
         Left            =   2025
         TabIndex        =   8
         Top             =   2700
         Width           =   60
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   315
      Left            =   5850
      TabIndex        =   1
      Top             =   75
      Width           =   765
   End
   Begin VB.ComboBox cboZip 
      Height          =   315
      Left            =   3525
      TabIndex        =   0
      Top             =   75
      Width           =   2265
   End
   Begin VB.Label lblRefresh 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Refreshing now..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   6165
      TabIndex        =   41
      Top             =   450
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Caption         =   "0°F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   780
      TabIndex        =   6
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zip Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2550
      TabIndex        =   5
      Top             =   150
      Width           =   840
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2550
      TabIndex        =   4
      Top             =   450
      Width           =   2925
   End
   Begin VB.Image imgWX 
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboZip_GotFocus()
  tmrCheck.Enabled = False
End Sub

Private Sub cboZip_KeyPress(KeyAscii As Integer)
  'allow for backspace
  If KeyAscii = vbKeyBack Then Exit Sub
  'if entire text is selected, then clear the text
  If cboZip.SelLength = Len(cboZip.Text) Then
    cboZip.Text = ""
  End If
  'allow only numeric keystrokes
  If Not IsNumeric(Chr(KeyAscii)) Then
    KeyAscii = 0
  End If
  'allow only 5 digits to be entered
  If Len(cboZip.Text) = 5 Then
    KeyAscii = 0
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdGo_Click()
  'make sure we have a 5-digit zip
  If cboZip.Text = "" Then
    MsgBox "You must enter a 5-digit Zip Code.", vbInformation, "WXLite"
    cboZip.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(Left(cboZip.Text, 5)) Then
    MsgBox "You must enter a 5-digit Zip Code.", vbInformation, "WXLite"
    cboZip.SetFocus
    Exit Sub
  End If
  'change mouse to HOurglass for both the command button and form
  'cmdGo.MousePointer = vbHourglass
  'Me.MousePointer = vbHourglass
  cmdGo.MousePointer = vbCustom
  cmdGo.MouseIcon = Me.Icon
  Me.MousePointer = vbCustom
  Me.MouseIcon = Me.Icon
  lblRefresh.Visible = True
  'get the weather checking to see that we found weather for the given zip
  If Not GetWX(Left(cboZip.Text, 5)) Then GoTo NO_WEATHER
  tmrCheck.Enabled = True
  'display the weather
  With Me
    .Caption = "WXLite - " & localWX.Temp & " - " & localWX.Name
    .lblName = localWX.Name
    .imgWX.Picture = .imlWX52.ListImages(localWX.IMG).Picture
    .lblCurrent = localWX.Cond
    .lblTemp = localWX.Temp
    .lblflTemp = localWX.flTemp
    .lblUdated = localWX.Udated
    .lblUV = localWX.UVIndex
    .lblWind = Replace(localWX.Wind(1), "<BR>", "")
    .lblHumidity = localWX.Humidity
    .lblPressure = localWX.Pressure
    .imgPressure = .imlPressure.ListImages(localWX.PressureIMG).Picture
    .lblDew = localWX.Dew
    .lblVisibility = localWX.Visibility
    .lblSR0 = localWX.SR0
    .lblSR1 = localWX.SR1
    .lblSS0 = localWX.SS0
    .lblSS1 = localWX.SS1
    .lblDaylight = localWX.Daylight
    .fraF1.Caption = localWX.f1Title
    .fraF2.Caption = localWX.f2Title
    .fraF3.Caption = localWX.f3Title
    .imgF1 = .imlWX31.ListImages(localWX.f1IMG).Picture
    .imgF2 = .imlWX31.ListImages(localWX.f2IMG).Picture
    .imgF3 = .imlWX31.ListImages(localWX.f3IMG).Picture
    .lblF1 = localWX.f1
    .lblF2 = localWX.f2
    .lblF3 = localWX.f3
  End With
  'add the zip to combo
  cboADD Left(cboZip.Text, 5) & " " & Left(localWX.Name, InStr(1, localWX.Name, "(") - 2)
  If localWX.Udated <> LastUpdate Then
    LastUpdate = localWX.Udated
'    With localWX
'      WriteUpdateLOG .Udated & ", " & .Temp & ", " & .flTemp & ", " & .Cond
'    End With
  End If
NO_WEATHER:
  'change back to default mouse
  cmdGo.MousePointer = vbDefault
  Me.MousePointer = vbDefault
  lblRefresh.Visible = False
End Sub

Private Sub cboADD(mItem As String)
  Dim bFound As Boolean
  'first look for the item
  For J = 0 To cboZip.ListCount - 1
    If cboZip.List(J) = mItem Then bFound = True
  Next J
  'if the item is not found, then add it to the list
  If Not bFound Then
    cboZip.AddItem mItem
  End If
End Sub

Private Function ReadLOG() As Boolean
  Dim isDefault As Boolean
  Dim tmp As String
  
  On Error GoTo ErrHandler
  Set FSO = New FileSystemObject
  'open the log file for reading, we don't want to create it yet if not there
  Set T = FSO.OpenTextFile(App.Path & "\wx.log", ForReading, False)
  'cycle through the log file and add each zip to the list
  While Not T.AtEndOfStream
    tmp = T.ReadLine
    If IsNumeric(Left(tmp, 5)) Then
      cboZip.AddItem tmp
      If cboZip.List(cboZip.NewIndex) Like "*[*]" Then
        cboZip.List(cboZip.NewIndex) = Mid(cboZip.List(cboZip.NewIndex), 1, Len(cboZip.List(cboZip.NewIndex)) - 1)
        cboZip.Text = cboZip.List(cboZip.NewIndex)
        isDefault = True
      End If
    End If
  Wend
  T.Close
  Set T = Nothing
  Set FSO = Nothing
  ReadLOG = True
  If isDefault Then
    cboZip.SelStart = 0
    cboZip.SelLength = Len(cboZip.Text)
    cmdGo_Click
  End If
  Exit Function
ErrHandler:
  ReadLOG = False
End Function

Private Function WriteLOG() As Boolean
  On Error GoTo ErrHandler
  Set FSO = New FileSystemObject
  'open the log file, create it if needed
  Set T = FSO.OpenTextFile(App.Path & "\wx.log", ForWriting, True)
  'cycle through the combo and write all zips in the list
  For J = 0 To cboZip.ListCount - 1
    If cboZip.List(J) = cboZip.Text Then
      T.WriteLine cboZip.List(J) & "*"
    Else
      T.WriteLine cboZip.List(J)
    End If
  Next J
  T.Close
  Set T = Nothing
  Set FSO = Nothing
  WriteLOG = True
  Exit Function
ErrHandler:
  WriteLOG = False
End Function

Private Sub cmdRemove_Click()
  'remove the selected item from the combo list
  On Error Resume Next
  cboZip.RemoveItem cboZip.ListIndex
  cboZip.SetFocus
End Sub

Private Sub Form_Load()
  StartFadeIn Me
  
  'get the list if zips from log file, if there
  ReadLOG
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    If localWX.Temp = "" Then
      Me.Caption = "No Data"
    Else
      Me.Caption = localWX.Temp & " in " & localWX.Name
    End If
  ElseIf Me.WindowState = vbNormal Then
    If localWX.Temp = "" Then
      Me.Caption = "WXLite - using www.Weather.com"
    Else
      Me.Caption = "WXLite - " & localWX.Temp & " - " & localWX.Name
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'write the list of zips to the log file
  If WriteLOG = False Then
    MsgBox "There was an error saving to the log file." & vbCrLf & vbCrLf & "Your settings were not saved.", vbCritical, "WXLite"
  End If
  StartFadeOut Me
End Sub

Private Sub tmrCheck_Timer()
  cmdGo_Click
End Sub

Private Sub tmrWind_Timer()
  If lblWind.Caption = localWX.Wind(1) Then
    lblWind.Caption = localWX.Wind(2)
  Else
    lblWind.Caption = localWX.Wind(1)
  End If
End Sub
