VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExMerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AsciiEXM"
   ClientHeight    =   12510
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   19845
   Icon            =   "frmExMerge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12510
   ScaleWidth      =   19845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_Intro 
      Caption         =   "Introduction"
      Height          =   3855
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   7755
      Begin VB.CommandButton cmdIntro_Exit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6480
         TabIndex        =   128
         Top             =   3180
         Width           =   915
      End
      Begin VB.CommandButton cmdIntro_Next 
         Caption         =   "Next"
         Height          =   375
         Left            =   4440
         TabIndex        =   123
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkDiscoverDiskdrives 
         Caption         =   "Discover disk drive volumes"
         Height          =   195
         Left            =   6300
         TabIndex        =   59
         Top             =   2160
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkDiscoverOnlineDBsOnly 
         Caption         =   "Discover online Databases only"
         Height          =   195
         Left            =   1620
         TabIndex        =   58
         Top             =   1500
         Width           =   2775
      End
      Begin VB.CommandButton cmdIntro_Discovery 
         Caption         =   "Discover"
         Height          =   375
         Left            =   2940
         TabIndex        =   50
         Top             =   1980
         Width           =   1515
      End
      Begin VB.TextBox txtLocalExchangeServer 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   49
         Top             =   1560
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ComboBox cboInitExchangeServers 
         Height          =   315
         Left            =   1620
         TabIndex        =   47
         Top             =   1080
         Width           =   2835
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Network Servers:"
         Height          =   195
         Left            =   6300
         TabIndex        =   46
         Top             =   1860
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Local Server"
         Height          =   195
         Left            =   6300
         TabIndex        =   45
         Top             =   1620
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label68 
         Caption         =   "Servers available:"
         Height          =   255
         Left            =   300
         TabIndex        =   180
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label55 
         Caption         =   "Please choose a server and press the ""Discover"" button."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   51
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label Label37 
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   1140
         Width           =   2475
      End
   End
   Begin VB.Frame Frame_Source 
      Caption         =   "Source"
      Height          =   4275
      Left            =   660
      TabIndex        =   22
      Top             =   4380
      Width           =   7095
      Begin ComctlLib.TreeView TreeView1 
         Height          =   2835
         Left            =   120
         TabIndex        =   100
         Top             =   600
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   5001
         _Version        =   327682
         Indentation     =   176
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdSource_Exit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6900
         TabIndex        =   127
         Top             =   3600
         Width           =   915
      End
      Begin VB.CommandButton cmdSource_Back 
         Caption         =   "Back"
         Height          =   435
         Left            =   120
         TabIndex        =   112
         Top             =   3600
         Width           =   915
      End
      Begin VB.CommandButton cmdSource_Next 
         Caption         =   "Next"
         Height          =   435
         Left            =   5340
         TabIndex        =   111
         Top             =   3600
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2835
         Left            =   2520
         TabIndex        =   37
         Top             =   600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   5001
         _Version        =   393216
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label Label41 
         Caption         =   "Exchange Structure:"
         Height          =   195
         Left            =   120
         TabIndex        =   163
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label59 
         Height          =   195
         Left            =   1380
         TabIndex        =   55
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label58 
         Height          =   195
         Left            =   4860
         TabIndex        =   54
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label57 
         Height          =   195
         Left            =   4860
         TabIndex        =   53
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label56 
         Height          =   195
         Left            =   4860
         TabIndex        =   52
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label47 
         Caption         =   "Email Accounts:"
         Height          =   195
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label48 
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4800
         Width           =   3255
      End
   End
   Begin VB.Frame Frame_INI 
      Caption         =   "Settings"
      Height          =   975
      Left            =   10440
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
      Begin VB.CommandButton cmdSettings_Exit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6120
         TabIndex        =   125
         Top             =   3780
         Width           =   915
      End
      Begin VB.CommandButton cmdSettings_Back 
         Caption         =   "Back"
         Height          =   375
         Left            =   1860
         TabIndex        =   116
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSettings_Next 
         Caption         =   "Next"
         Height          =   375
         Left            =   4080
         TabIndex        =   115
         Top             =   3840
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3015
         Left            =   180
         TabIndex        =   62
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5318
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Paths"
         TabPicture(0)   =   "frmExMerge.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label15"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label17"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label18"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label25"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label14"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label16"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Command3"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtExMergePath"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtPSTPath"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Command2"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Command12"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtLogPath"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Command14"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txtINIPath"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtINIFileName"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtMailBoxFileName"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtMailBoxPath"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Command11"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "chkUseRoot"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Command17"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtRootPath"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).ControlCount=   23
         TabCaption(1)   =   "MailBoxes.txt"
         TabPicture(1)   =   "frmExMerge.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtOrganization"
         Tab(1).Control(1)=   "txtCN1"
         Tab(1).Control(2)=   "Command8"
         Tab(1).Control(3)=   "Command9"
         Tab(1).Control(4)=   "Command10"
         Tab(1).Control(5)=   "txtGroup"
         Tab(1).Control(6)=   "Label6"
         Tab(1).Control(7)=   "Label7"
         Tab(1).Control(8)=   "Label8"
         Tab(1).Control(9)=   "Label12"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "INI"
         TabPicture(2)   =   "frmExMerge.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmbLogLevel"
         Tab(2).Control(1)=   "chkUserData"
         Tab(2).Control(2)=   "chkFolderData"
         Tab(2).Control(3)=   "chkFolderRules"
         Tab(2).Control(4)=   "chkDumpster"
         Tab(2).Control(5)=   "txtDR_EndTime"
         Tab(2).Control(6)=   "txtDR_EndDate"
         Tab(2).Control(7)=   "txtDR_StartTime"
         Tab(2).Control(8)=   "txtDR_StartDate"
         Tab(2).Control(9)=   "Picture1"
         Tab(2).Control(10)=   "Picture4"
         Tab(2).Control(11)=   "Picture6"
         Tab(2).Control(12)=   "Picture2"
         Tab(2).Control(13)=   "chkDateRange"
         Tab(2).Control(14)=   "Label20"
         Tab(2).Control(15)=   "Label19"
         Tab(2).Control(16)=   "Label33"
         Tab(2).Control(17)=   "Label34"
         Tab(2).ControlCount=   18
         Begin VB.TextBox txtRootPath 
            Height          =   315
            Left            =   1380
            TabIndex        =   119
            Top             =   480
            Width           =   2595
         End
         Begin VB.CommandButton Command17 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   118
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox chkUseRoot 
            Caption         =   "Use"
            Height          =   195
            Left            =   4620
            TabIndex        =   117
            Top             =   540
            Width           =   675
         End
         Begin VB.ComboBox cmbLogLevel 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmExMerge.frx":0496
            Left            =   -69240
            List            =   "frmExMerge.frx":04A6
            TabIndex        =   106
            Text            =   "Medium"
            Top             =   2340
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkUserData 
            Caption         =   "User Data"
            Height          =   195
            Left            =   -70380
            TabIndex        =   105
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkFolderData 
            Caption         =   "Folder Data"
            Height          =   195
            Left            =   -70380
            TabIndex        =   104
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox chkFolderRules 
            Caption         =   "Folder Rules"
            Height          =   195
            Left            =   -70380
            TabIndex        =   103
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox chkDumpster 
            Caption         =   "Dumpster"
            Height          =   195
            Left            =   -70380
            TabIndex        =   102
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtDR_EndTime 
            Height          =   315
            Left            =   -72480
            TabIndex        =   101
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   1740
            Width           =   1095
         End
         Begin VB.TextBox txtDR_EndDate 
            Height          =   315
            Left            =   -73980
            TabIndex        =   99
            Top             =   1740
            Width           =   975
         End
         Begin VB.TextBox txtDR_StartTime 
            Height          =   315
            Left            =   -72480
            TabIndex        =   98
            ToolTipText     =   "Ex. '4:00 PM'"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtDR_StartDate 
            Height          =   315
            Left            =   -73980
            TabIndex        =   97
            Top             =   1320
            Width           =   975
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -72960
            Picture         =   "frmExMerge.frx":04CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   96
            Top             =   1740
            Width           =   270
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -72960
            Picture         =   "frmExMerge.frx":090C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   95
            Top             =   1320
            Width           =   270
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -71280
            Picture         =   "frmExMerge.frx":0D4E
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   94
            Top             =   1740
            Width           =   390
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -71280
            Picture         =   "frmExMerge.frx":0ED8
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   93
            Top             =   1320
            Width           =   390
         End
         Begin VB.CheckBox chkDateRange 
            Caption         =   "Only use messages with this date and time range"
            Height          =   195
            Left            =   -74820
            TabIndex        =   92
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtOrganization 
            Height          =   315
            Left            =   -73320
            TabIndex        =   87
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtCN1 
            Height          =   315
            Left            =   -73320
            TabIndex        =   86
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton Command8 
            Caption         =   "?"
            Height          =   315
            Left            =   -71040
            TabIndex        =   85
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton Command9 
            Caption         =   "?"
            Height          =   315
            Left            =   -71040
            TabIndex        =   84
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton Command10 
            Caption         =   "?"
            Height          =   315
            Left            =   -71040
            TabIndex        =   83
            Top             =   1920
            Width           =   375
         End
         Begin VB.TextBox txtGroup 
            Height          =   315
            Left            =   -73320
            TabIndex        =   82
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton Command11 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   79
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtMailBoxPath 
            Height          =   315
            Left            =   1380
            TabIndex        =   78
            Top             =   1320
            Width           =   2595
         End
         Begin VB.TextBox txtMailBoxFileName 
            Height          =   315
            Left            =   4740
            TabIndex        =   77
            Text            =   "mailboxes.txt"
            Top             =   1320
            Width           =   1275
         End
         Begin VB.TextBox txtINIFileName 
            Height          =   315
            Left            =   4740
            TabIndex        =   71
            Text            =   "ExMerge.ini"
            Top             =   900
            Width           =   1275
         End
         Begin VB.TextBox txtINIPath 
            Height          =   315
            Left            =   1380
            TabIndex        =   70
            Top             =   900
            Width           =   2595
         End
         Begin VB.CommandButton Command14 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   69
            Top             =   900
            Width           =   375
         End
         Begin VB.TextBox txtLogPath 
            Height          =   315
            Left            =   1380
            TabIndex        =   68
            Top             =   2580
            Width           =   2595
         End
         Begin VB.CommandButton Command12 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   67
            Top             =   2580
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   66
            Top             =   1740
            Width           =   375
         End
         Begin VB.TextBox txtPSTPath 
            Height          =   315
            Left            =   1380
            TabIndex        =   65
            Top             =   2160
            Width           =   2595
         End
         Begin VB.TextBox txtExMergePath 
            Height          =   315
            Left            =   1380
            TabIndex        =   64
            Top             =   1740
            Width           =   2595
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   315
            Left            =   4080
            TabIndex        =   63
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Root Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   120
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label Label20 
            Caption         =   "Log Level:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -70440
            TabIndex        =   110
            Top             =   2400
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label19 
            Caption         =   "Export the following:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   109
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label33 
            Caption         =   "End Date:"
            Height          =   255
            Left            =   -74820
            TabIndex        =   108
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label34 
            Caption         =   "Start Date:"
            Height          =   255
            Left            =   -74820
            TabIndex        =   107
            Top             =   1380
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "Organization (/O=):"
            Height          =   195
            Left            =   -74820
            TabIndex        =   91
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label Label7 
            Caption         =   "CN1 (/CN=):"
            Height          =   195
            Left            =   -74820
            TabIndex        =   90
            Top             =   1980
            Width           =   1395
         End
         Begin VB.Label Label8 
            Caption         =   "Group (/OU):"
            Height          =   195
            Left            =   -74820
            TabIndex        =   89
            Top             =   1500
            Width           =   1395
         End
         Begin VB.Label Label12 
            Caption         =   "MailBox.txt Template Settings:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   88
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label14 
            Caption         =   "MailBox Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   81
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Label Label25 
            Caption         =   "\"
            Height          =   255
            Left            =   4560
            TabIndex        =   80
            Top             =   1380
            Width           =   135
         End
         Begin VB.Label Label18 
            Caption         =   "\"
            Height          =   255
            Left            =   4560
            TabIndex        =   76
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label17 
            Caption         =   "INI Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   75
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Log Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   74
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "PST Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   2220
            Width           =   1155
         End
         Begin VB.Label Label13 
            Caption         =   "Exmerge Path:"
            Height          =   195
            Left            =   180
            TabIndex        =   72
            Top             =   1800
            Width           =   1155
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "?"
         Height          =   315
         Left            =   5340
         TabIndex        =   43
         Top             =   7140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton optProcess_AsGroup 
         Caption         =   "as a group"
         Height          =   195
         Left            =   3600
         TabIndex        =   41
         Top             =   7200
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optProcess_individually 
         Caption         =   "Individually"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1800
         TabIndex        =   40
         Top             =   7200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Use Default"
         Height          =   315
         Left            =   5880
         TabIndex        =   28
         Top             =   7140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Process accounts:"
         Height          =   195
         Left            =   300
         TabIndex        =   42
         Top             =   7200
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame F_Header 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   12600
      TabIndex        =   177
      Top             =   660
      Width           =   8115
      Begin VB.Label Label67 
         BackColor       =   &H80000009&
         Caption         =   "EXM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   179
         Top             =   120
         Width           =   3315
      End
      Begin VB.Label Label53 
         BackColor       =   &H80000009&
         Caption         =   "By CPS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   178
         Top             =   660
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   15840
      TabIndex        =   168
      Top             =   9960
      Width           =   1995
      Begin VB.CheckBox Check11 
         Caption         =   "Starting Backup."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   172
         Top             =   240
         Width           =   1755
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Backup Completed."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   171
         Top             =   540
         Width           =   1755
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Compiling Details."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   170
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Processing Options"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   169
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label Label31 
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   660
         TabIndex        =   174
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Errors:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   173
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Frame Frame_Process 
      Caption         =   "Process"
      Height          =   4635
      Left            =   11640
      TabIndex        =   0
      Top             =   5040
      Width           =   7455
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   240
         TabIndex        =   167
         Top             =   3000
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdProcess_Exit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5100
         TabIndex        =   124
         Top             =   3360
         Width           =   915
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   2460
         TabIndex        =   5
         Top             =   540
         Width           =   3555
      End
      Begin VB.CommandButton cmdProcess_Start 
         Caption         =   "Start"
         Height          =   435
         Left            =   3600
         TabIndex        =   4
         Top             =   3360
         Width           =   915
      End
      Begin VB.CommandButton cmdProcess_Halt 
         Caption         =   "Halt"
         Height          =   435
         Left            =   2460
         TabIndex        =   3
         Top             =   3360
         Width           =   915
      End
      Begin VB.ListBox List5 
         Height          =   2400
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton cmdProcess_Back 
         Caption         =   "Back"
         Height          =   435
         Left            =   420
         TabIndex        =   1
         Top             =   3660
         Width           =   975
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   195
         Left            =   2460
         TabIndex        =   175
         Top             =   3000
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Progress:"
         Height          =   195
         Left            =   2460
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Accounts:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame_AdvancedOptions 
      Caption         =   "Options"
      Height          =   4095
      Left            =   10740
      TabIndex        =   8
      Top             =   5880
      Width           =   6975
      Begin VB.CommandButton cmdOptions_Exit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   5400
         TabIndex        =   126
         Top             =   3480
         Width           =   915
      End
      Begin VB.CommandButton cmdOptions_Back 
         Caption         =   "Back"
         Height          =   375
         Left            =   1140
         TabIndex        =   114
         Top             =   3540
         Width           =   1095
      End
      Begin VB.CommandButton cmdOptions_Next 
         Caption         =   "Next"
         Height          =   375
         Left            =   3360
         TabIndex        =   113
         Top             =   3540
         Width           =   1095
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2895
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Compression"
         TabPicture(0)   =   "frmExMerge.frx":1062
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label36"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label21"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command32"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtDotZipFilePath"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtDotZipFileName"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cboCompressionOption"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Send Files"
         TabPicture(1)   =   "frmExMerge.frx":107E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label26"
         Tab(1).Control(1)=   "Label23"
         Tab(1).Control(2)=   "Command22"
         Tab(1).Control(3)=   "txtMoveToFolder"
         Tab(1).Control(4)=   "cboSendType"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Clean Up"
         TabPicture(2)   =   "frmExMerge.frx":109A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkDeleteAfterZip"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Oversize"
         TabPicture(3)   =   "frmExMerge.frx":10B6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label28"
         Tab(3).Control(1)=   "Label27"
         Tab(3).Control(2)=   "Label65"
         Tab(3).Control(3)=   "Label66"
         Tab(3).Control(4)=   "Label64"
         Tab(3).Control(5)=   "Label63"
         Tab(3).Control(6)=   "Label32"
         Tab(3).Control(7)=   "cboPSTLimitOptions"
         Tab(3).Control(8)=   "cboLimit_Logical_Months"
         Tab(3).Control(9)=   "optLimit_Logical_ByMonths"
         Tab(3).Control(10)=   "optLimit_Logical_Mailbox_Size"
         Tab(3).Control(11)=   "txtLimit_Logical_GreaterThan"
         Tab(3).Control(12)=   "chkPSTLimit_Logical"
         Tab(3).ControlCount=   13
         TabCaption(4)   =   "Threads"
         TabPicture(4)   =   "frmExMerge.frx":10D2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "chkShowInterface"
         Tab(4).Control(1)=   "chkDefaultThreads"
         Tab(4).Control(2)=   "cboThreadCount"
         Tab(4).Control(3)=   "Label35"
         Tab(4).ControlCount=   4
         Begin VB.CheckBox chkPSTLimit_Logical 
            Caption         =   "Logical mode"
            Height          =   195
            Left            =   -74700
            TabIndex        =   176
            Top             =   1980
            Width           =   1335
         End
         Begin VB.TextBox txtLimit_Logical_GreaterThan 
            Height          =   315
            Left            =   -69300
            TabIndex        =   164
            Top             =   3000
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.CheckBox chkShowInterface 
            Caption         =   "Show ExMerge interface while processing"
            Height          =   195
            Left            =   -74640
            TabIndex        =   152
            Top             =   1200
            Width           =   3315
         End
         Begin VB.CheckBox chkDefaultThreads 
            Caption         =   "Use Default"
            Height          =   195
            Left            =   -70860
            TabIndex        =   151
            Top             =   660
            Width           =   1215
         End
         Begin VB.ComboBox cboThreadCount 
            Height          =   315
            ItemData        =   "frmExMerge.frx":10EE
            Left            =   -73860
            List            =   "frmExMerge.frx":1101
            TabIndex        =   150
            Top             =   600
            Width           =   2835
         End
         Begin VB.OptionButton optLimit_Logical_Mailbox_Size 
            Caption         =   "By mailBox size"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -70020
            TabIndex        =   145
            Top             =   2700
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.OptionButton optLimit_Logical_ByMonths 
            Caption         =   "By Months"
            Height          =   195
            Left            =   -70020
            TabIndex        =   144
            Top             =   2400
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cboLimit_Logical_Months 
            Height          =   315
            ItemData        =   "frmExMerge.frx":11AC
            Left            =   -72480
            List            =   "frmExMerge.frx":11D4
            TabIndex        =   143
            Text            =   "12"
            Top             =   1920
            Width           =   915
         End
         Begin VB.ComboBox cboPSTLimitOptions 
            Height          =   315
            ItemData        =   "frmExMerge.frx":11FF
            Left            =   -69600
            List            =   "frmExMerge.frx":1209
            TabIndex        =   142
            Text            =   "None"
            Top             =   2100
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.CheckBox chkDeleteAfterZip 
            Caption         =   "Delete PST if left over"
            Height          =   195
            Left            =   -74640
            TabIndex        =   141
            Top             =   660
            Width           =   2295
         End
         Begin VB.ComboBox cboSendType 
            Height          =   315
            ItemData        =   "frmExMerge.frx":121C
            Left            =   -73740
            List            =   "frmExMerge.frx":1229
            TabIndex        =   138
            Text            =   "None"
            Top             =   540
            Width           =   1995
         End
         Begin VB.TextBox txtMoveToFolder 
            Height          =   315
            Left            =   -73740
            TabIndex        =   137
            Top             =   1020
            Width           =   3795
         End
         Begin VB.CommandButton Command22 
            Caption         =   "..."
            Height          =   315
            Left            =   -69840
            TabIndex        =   136
            Top             =   1020
            Width           =   375
         End
         Begin VB.ComboBox cboCompressionOption 
            Height          =   315
            ItemData        =   "frmExMerge.frx":123F
            Left            =   1260
            List            =   "frmExMerge.frx":1241
            TabIndex        =   132
            Text            =   "None"
            Top             =   540
            Width           =   1995
         End
         Begin VB.TextBox txtDotZipFileName 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            TabIndex        =   131
            Text            =   "MailBoxes.zip"
            Top             =   1500
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.TextBox txtDotZipFilePath 
            Height          =   315
            Left            =   1260
            TabIndex        =   130
            Top             =   1020
            Width           =   3795
         End
         Begin VB.CommandButton Command32 
            Caption         =   "..."
            Height          =   315
            Left            =   5160
            TabIndex        =   129
            Top             =   1020
            Width           =   375
         End
         Begin VB.Label Label32 
            Caption         =   $"frmExMerge.frx":1243
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -74700
            TabIndex        =   182
            Top             =   840
            Width           =   5535
         End
         Begin VB.Label Label63 
            Caption         =   " if box > "
            Height          =   195
            Left            =   -70020
            TabIndex        =   166
            Top             =   3060
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label64 
            Caption         =   "MBs"
            Height          =   195
            Left            =   -68340
            TabIndex        =   165
            Top             =   3060
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label35 
            Caption         =   "Threads:"
            Height          =   195
            Left            =   -74640
            TabIndex        =   153
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label66 
            Caption         =   "Months"
            Height          =   195
            Left            =   -71520
            TabIndex        =   149
            Top             =   1980
            Width           =   675
         End
         Begin VB.Label Label65 
            Caption         =   "Per:"
            Height          =   195
            Left            =   -73020
            TabIndex        =   148
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label27 
            Caption         =   "Mode:"
            Height          =   195
            Left            =   -70140
            TabIndex        =   147
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Oversized Mailbox Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74700
            TabIndex        =   146
            Top             =   540
            Width           =   1995
         End
         Begin VB.Label Label23 
            Caption         =   "Method:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   140
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Location:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   139
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Type:"
            Height          =   195
            Left            =   180
            TabIndex        =   135
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Location:"
            Height          =   195
            Left            =   180
            TabIndex        =   134
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label36 
            Caption         =   "Output File:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   180
            TabIndex        =   133
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Label Label62 
         Height          =   195
         Left            =   3000
         TabIndex        =   60
         Top             =   1500
         Width           =   2775
      End
   End
   Begin VB.ComboBox cboEmailServers 
      Height          =   315
      Left            =   9900
      TabIndex        =   159
      Top             =   7800
      Width           =   1635
   End
   Begin VB.ComboBox cboStorageGroups 
      Height          =   315
      Left            =   9900
      TabIndex        =   158
      Top             =   8220
      Width           =   1635
   End
   Begin VB.ComboBox cboMailBoxStores 
      Height          =   315
      Left            =   9900
      TabIndex        =   157
      Top             =   8640
      Width           =   1635
   End
   Begin VB.OptionButton optProcessIndividually 
      Caption         =   "individually"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6300
      TabIndex        =   155
      Top             =   11400
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.OptionButton optProcessAsGroup 
      Caption         =   "as a group"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4620
      TabIndex        =   154
      Top             =   11400
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ComboBox cboEncryptionOption 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmExMerge.frx":12EE
      Left            =   1380
      List            =   "frmExMerge.frx":12FB
      TabIndex        =   121
      Text            =   "None"
      Top             =   11460
      Width           =   1575
   End
   Begin VB.Frame Frame_Menu 
      Height          =   3495
      Left            =   12300
      TabIndex        =   11
      Top             =   9060
      Width           =   2595
      Begin VB.CommandButton Command34 
         Caption         =   "5"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   435
      End
      Begin VB.CommandButton Command35 
         Caption         =   "3"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   435
      End
      Begin VB.CommandButton Command36 
         Caption         =   "4"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton Command38 
         Caption         =   "1"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   435
      End
      Begin VB.CommandButton Command39 
         Caption         =   "0"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label38 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   21
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label Label39 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   20
         Top             =   1980
         Width           =   1635
      End
      Begin VB.Label Label40 
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   19
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label42 
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   18
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label Label43 
         Caption         =   "Introduction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   17
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.ListBox List6 
      Height          =   1425
      ItemData        =   "frmExMerge.frx":1317
      Left            =   17040
      List            =   "frmExMerge.frx":132D
      TabIndex        =   39
      Top             =   7260
      Width           =   2055
   End
   Begin VB.ListBox List7 
      Height          =   1230
      Left            =   16620
      TabIndex        =   38
      Top             =   5820
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   32
      Top             =   9120
      Width           =   9975
      Begin VB.CommandButton Command1 
         Caption         =   "Messages:"
         Height          =   375
         Left            =   6240
         TabIndex        =   185
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtTotalCombinedCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   3900
         TabIndex        =   35
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtTotalCombinedSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   720
         TabIndex        =   33
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label61 
         Caption         =   "MBs"
         Height          =   195
         Left            =   2160
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label54 
         Caption         =   "Mail Boxes"
         Height          =   195
         Left            =   3000
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label49 
         Caption         =   "Size:"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   60
      TabIndex        =   25
      Top             =   10200
      Width           =   9975
   End
   Begin MSMask.MaskEdBox mskStartDate 
      Height          =   375
      Left            =   2640
      TabIndex        =   44
      Top             =   11820
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "  /  /    "
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   975
      Left            =   15720
      TabIndex        =   183
      Top             =   8940
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label Label3 
      Caption         =   "Drives:"
      Height          =   195
      Left            =   15720
      TabIndex        =   184
      Top             =   8700
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "AsciiEXM is a flexible application that allows you backup many aspects of Exchange 2003 via the MS ExMerge tool."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12780
      TabIndex        =   181
      Top             =   2340
      Width           =   6315
   End
   Begin VB.Label Label44 
      Caption         =   "Email Server:"
      Height          =   195
      Left            =   8640
      TabIndex        =   162
      Top             =   7860
      Width           =   1155
   End
   Begin VB.Label Label45 
      Caption         =   "Storage Groups:"
      Height          =   195
      Left            =   8640
      TabIndex        =   161
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label46 
      Caption         =   "MailBox Stores:"
      Height          =   195
      Left            =   8640
      TabIndex        =   160
      Top             =   8700
      Width           =   1275
   End
   Begin VB.Label Label24 
      Caption         =   "Process backup"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3060
      TabIndex        =   156
      Top             =   11400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label22 
      Caption         =   "Encryption:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   300
      TabIndex        =   122
      Top             =   11520
      Width           =   1035
   End
   Begin VB.Label Label60 
      Height          =   195
      Left            =   960
      TabIndex        =   56
      Top             =   9780
      Width           =   915
   End
   Begin VB.Label Label52 
      Caption         =   "Backup Email, put forward into other user, delete email account, save AD user profile, delete AD user."
      Height          =   555
      Left            =   4080
      TabIndex        =   31
      Top             =   11700
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label51 
      Caption         =   "Recipients, security, Distribution, ad groups  .... search for string in a specific spot/s"
      Height          =   435
      Left            =   7560
      TabIndex        =   30
      Top             =   11820
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label50 
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   9780
      Width           =   795
   End
   Begin VB.Label Label10 
      Height          =   195
      Left            =   2580
      TabIndex        =   27
      Top             =   9780
      Width           =   7635
   End
   Begin VB.Label Label4 
      Caption         =   "Process:"
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   9780
      Width           =   615
   End
End
Attribute VB_Name = "frmExMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fTop  As Integer
Private fLeft  As Integer
Private fHeight  As Integer
Private fWidth  As Integer

Dim ltotalsize As Double
Dim ltotalcount As Long
Dim CurrentFile As String

Dim bInitStage As Boolean

Dim DBs() As String

Dim ndSrv As Node
Dim ndSG As Node
Dim ndMBX As Node



'Check Winzip
Function WinZipExists() As Boolean
    Dim sPath As String
    Dim sVal As String
    Dim iCheckSum As Integer
    
    WinZipExists = False

On Error Resume Next
    
    'check if winzip exists
    sPath = "Software\Nico Mak Computing\WinZip"
    If CheckRegistryKey(HKEY_CURRENT_USER, sPath) Then
        iCheckSum = 1
    End If
    
    'check if winzip path exists
    sPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winzip.exe"
    sVal = ""
    sVal = Replace(QueryValue(HKEY_LOCAL_MACHINE, sPath, sVal), vbNullChar, "")
    If sVal = "C:\PROGRA~1\WINZIP\winzip32.exe" Then
        If FileExists(sVal) Then iCheckSum = iCheckSum + 3
    End If
    
    'check if winzip is registered
    sPath = "Software\Nico Mak Computing\WinZip\WinIni"
    sVal = "SN"
    sVal = Replace(QueryValue(HKEY_CURRENT_USER, sPath, sVal), vbNullChar, "")
    If sVal <> Empty Then
        iCheckSum = iCheckSum + 5
    End If
    
    If iCheckSum = 9 Then WinZipExists = True
    
End Function

Function CompressionExists() As Boolean

    cboCompressionOption.Enabled = False
    
    cboCompressionOption.Text = "none"
    cboCompressionOption.AddItem "none"
            
    If WinZipExists Then
        cboCompressionOption.AddItem "Winzip"
        cboCompressionOption.Enabled = True
    End If
    
    'winrar
    
    'cab
            
End Function


Private Sub cboCompressionOption_Change()
    If bInitStage = False And ltotalsize <> 0 Then
        If LCase(Trim(cboCompressionOption.Text)) = "none" Then
            Label62.Caption = ""
            txtTotalCombinedSize.Text = FormatNumber(ltotalsize / 2, 2, vbTrue, vbTrue, vbTrue)
        Else
            Label62.Caption = "Please allow for double the size, see below"
            txtTotalCombinedSize.Text = FormatNumber(ltotalsize * 2, 2, vbTrue, vbTrue, vbTrue)
        End If
    End If
End Sub

Private Sub cboEmailServers_Change()
'    GetStorageGroups Trim(cboEmailServers.Text), cboStorageGroups
    ''''''Main.Current.Server = Trim(cboEmailServers.Text)
    Show_StorageGroups Trim(cboEmailServers.Text)
End Sub

Private Sub cboEmailServers_Click()
'    GetStorageGroups Trim(cboEmailServers.Text), cboStorageGroups
    '''''Main.Current.Server = Trim(cboEmailServers.Text)
    Show_StorageGroups Trim(cboEmailServers.Text)
End Sub

Private Sub cboMailBoxStores_Change()
    'GetMailBoxes Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), Trim(cboMailBoxStores.Text), cboMailBoxStores
    Show_MailBoxes Trim(cboMailBoxStores.Text)

End Sub

Private Sub cboMailBoxStores_Click()
    'GetMailBoxes Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), Trim(cboMailBoxStores.Text), cboMailBoxStores
    Show_MailBoxes Trim(cboMailBoxStores.Text)
End Sub

Private Sub cboPSTLimitOptions_Change()
    If LCase(Trim(cboPSTLimitOptions.Text)) <> "none" Then
        chkDateRange.Enabled = False
        txtDR_StartDate.Enabled = False
        txtDR_EndDate.Enabled = False
        txtDR_StartTime.Enabled = False
        txtDR_EndTime.Enabled = False
        chkDateRange.Value = 0
    Else
        chkDateRange.Enabled = True
        txtDR_StartDate.Enabled = True
        txtDR_EndDate.Enabled = True
        txtDR_StartTime.Enabled = True
        txtDR_EndTime.Enabled = True
    End If
End Sub

Private Sub cboSendType_Change()
    If LCase(Trim(cboSendType.Text)) = "move" Then
        chkDeleteAfterZip.Value = 0
        chkDeleteAfterZip.Enabled = False
    Else
        chkDeleteAfterZip.Enabled = True
    End If
End Sub

Private Sub cboStorageGroups_Change()
    'GetMailBoxStores Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), cboMailBoxStores
    Show_MailBoxStoreDBs Trim(cboStorageGroups.Text)
End Sub

Private Sub cboStorageGroups_Click()
    'GetMailBoxStores Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), cboMailBoxStores
    Show_MailBoxStoreDBs Trim(cboStorageGroups.Text)
End Sub

Private Sub chkDefaultThreads_Click()
    ThreadCount CInt(Trim(txtTotalCombinedCount.Text))
End Sub

Private Sub chkDefaultThreads_KeyUp(KeyCode As Integer, Shift As Integer)
    ThreadCount CInt(Trim(txtTotalCombinedCount.Text))
End Sub



'search for exchange servers -> databases -> mailboxes/groups
'possibly use CSVDE -f list.csv to find the servers -> then exchange and Domain Controllers.
'if not apart of the domain search for computers with ldap traffic.
'if DC found then run "LDAP://Server/RootDSE" to start a dig for info.

'verify all tools, components, apps, and exmerge exist

'create fields ini/txt/log/etc files

'create variable structure

'add zip feature with/verifying that winzip32 is installed - zipping available

'drive discovery with size - functionality available

'calculations for drive vs accounts being backed up plus zip file

'a move to function, if the drive will get too full

'mailbox prediscovery - size

'2gb file limit solution
    'get total size of mailbox
    'get create date of mailbox
        'option 1:
            'take the total amount of time and divide by the size of mailbox in GBs
            'then do a back up of each segment of time seperately
            'this will attempt to average a 1 GB file per segment
        'option 2:
            'start with start date and do (for example) 3 month backups to the same pst
            'file and check size after each pass, when becomes greater than 1GB then
            'create then next file to continue the backup process.  keep repeating until mailbox



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub chkPSTLimit_Logical_Click()
    If chkPSTLimit_Logical.Value = 0 Then
        cboLimit_Logical_Months.Enabled = False
        chkDateRange.Enabled = False
        txtDR_StartDate.Enabled = False
        txtDR_EndDate.Enabled = False
        txtDR_StartTime.Enabled = False
        txtDR_EndTime.Enabled = False
        chkDateRange.Value = 0
    Else
        cboLimit_Logical_Months.Enabled = True
        chkDateRange.Enabled = True
        txtDR_StartDate.Enabled = True
        txtDR_EndDate.Enabled = True
        txtDR_StartTime.Enabled = True
        txtDR_EndTime.Enabled = True
    End If
End Sub

Private Sub cmdIntro_Discovery_Click()
    cmdIntro_Next.Enabled = False
    cmdIntro_Exit.Enabled = False
    prcDiscovery
    cmdIntro_Next.Enabled = True
    cmdIntro_Exit.Enabled = True
End Sub

Private Sub cmdIntro_Exit_Click()
    Unload Me
End Sub

Private Sub cmdIntro_Next_Click()
    ShowFrame 1
End Sub

Private Sub cmdOptions_Back_Click()
    ShowFrame 4
End Sub

Private Sub cmdOptions_Exit_Click()
    Unload Me
End Sub

Private Sub cmdOptions_Next_Click()
    ShowFrame 5
End Sub

Private Sub cmdProcess_Back_Click()
    ShowFrame 3
End Sub

Private Sub cmdProcess_Exit_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Halt_Click()
    bHaltProcess = True
End Sub

Private Sub cmdProcess_Start_Click()
    Dim v As Long
    cmdProcess_Back.Enabled = False
    cmdProcess_Exit.Enabled = False
    bHaltProcess = False
    
    'per email progress
    ProgressBar1.Min = 1
    ProgressBar1.Max = 5
    
    'overall progress
    ProgressBar2.Min = 1
    v = CLng(Trim(Me.txtTotalCombinedCount.Text) + 1)
    Debug.Print v
    ProgressBar2.Max = v
    
    
    StartProcess
    
    cmdProcess_Back.Enabled = True
    cmdProcess_Exit.Enabled = True
    bHaltProcess = True
End Sub

Private Sub cmdSettings_Back_Click()
    ShowFrame 1
End Sub

Private Sub cmdSettings_Exit_Click()
    Unload Me
End Sub



Sub Output_System_Messages(sMsg As String)
    'Output_2_List sMsg
End Sub

Sub Output_2_List(lst As ListBox, sMsg As String)
    lst.AddItem sMsg
    lst.ListIndex = lst.ListCount - 1
End Sub

Sub Output_System_Messages_Caption(lbl As Label, sMsg As String)
    lbl.Caption = sMsg
End Sub

Sub Output_System_Messages_MSSQL(sMsg As String)

End Sub

Function StartProcess() As Boolean
    'starting backup
    Check11.Value = 1
    Check11.Refresh
    
    ProgressBar2.Value = 1
    
    'copy all flagged accounts into a current array
    CopyOnlyFlaggedAccounts
    
    
    'if one at a time or group
    If optProcessIndividually.Value Then
        
        
        ProcessEachAccount
            
        
    ElseIf optProcessAsGroup.Value Then
        
        ProgressBar1.Value = 1
        ProgressBar2.Value = 1
            
        'Create the ini file
        If funCreateExMergeINI Then
            Output_2_List List2, "Processing Group"
            ProgressBar1.Value = 2
            ProgressBar2.Value = 2
        
            If funCreate_Mailboxes_txt(MSFlexGrid1) Then
                ProgressBar1.Value = 3
                ProgressBar2.Value = 3
                    
                ProcessBackup
                Check12.Value = 1
                Check12.Refresh
                                
                'apply options
                ProgressBar1.Value = 4
                ProgressBar2.Value = 4
                Check15.Value = 1
                Check15.Refresh
                CycleThroughFiles
                
                ProgressBar1.Value = 5
                ProgressBar2.Value = 5
            Else
                List2.AddItem "An error has occured while creating the MailBoxes.txt file."
            End If
        Else
            List2.AddItem "An error has occured while creating the ExMergeX.ini file."
        End If
        
    End If
    
    
End Function


Function ProcessOptions_PerUser(sAccountIndex As Integer)
    Dim Pst As String
    Dim tmpMsg As String
    
    
    With CurrentMailBoxes(sAccountIndex)
    
        If .FileName = "" Then
            Pst = Trim(txtPSTPath.Text) & "\" & .Alias & ".pst"
        Else
            Pst = .FileName & ".pst"
        End If
        
        CurrentFile = Pst
            
        If Exists_fso(Pst) Then
            
            ' encrypt
            '''''''''''''''''''
                
            ' compress
            If LCase(Trim(cboCompressionOption.Text)) = "winzip" Then
                            
                'Compression .Alias
                If Compress_File2File(.FileName & ".pst", .FileName & ".zip") Then
                    'list.add "successful
                Else
                    'error compressioning
                End If
                
            End If
            
            'move/copy/etc a file
            SendingFile_Account CurrentFile, Build_Dest_SendFile(CurrentFile)
            
            
            'clean up any left over psts if specified
            CleanUpFiles Pst
                               
                
        Else
            Output_2_List List2, vbTab & vbTab & "-> File: '" & Pst & "' does not exist"
            
        End If
        
    End With
        
End Function

Function Build_Dest_SendFile(sCurrentFile As String) As String
    Dim tmp() As String
    
On Error GoTo Err:
    
    Build_Dest_SendFile = ""
    tmp = Split(sCurrentFile, "\")
    
    If UBound(tmp) > 0 Then
        Build_Dest_SendFile = Trim(txtMoveToFolder.Text) & "\" & tmp(UBound(tmp))
    End If
    Exit Function
    
Err:
    Debug.Print Err.Number & " " & Err.Description
End Function

Function Compression(sAccount As String)
    Dim tmpMsg As String

    tmpMsg = "Options->Compress->WinZip(" & sAccount & ".pst)"
    Output_System_Messages_Caption Label10, tmpMsg
                
    'if winzip, cab, etc...
    CurrentFile = ZipFileSettings_Account(sAccount)
                
    tmpMsg = "Options->Compress->WinZip(" & sAccount & ".pst)"
    Output_System_Messages_Caption Label10, tmpMsg
End Function



Public Function funCreate_Mailboxes_PerUser(sAccount As String) As Boolean

    funCreate_Mailboxes_PerUser = funCreatefile(CreateMailBoxTxTPath, SetupMailBoxTxt_PerUser(sAccount))
    
    If funCreate_Mailboxes_PerUser Then
        Output_2_List List2, "MailBox file was created."
    Else
        Output_2_List List2, "Error creating MailBox file"
    End If
    
End Function



Function SetupMailBoxTxt_PerUser(sAccount As String) As String
    Dim sTemplate As String, sContent As String
    
    sTemplate = "/o=" & Trim(txtOrganization.Text) & _
                "/ou=" & Trim(txtGroup.Text) & _
                "/cn=" & Trim(txtCN1.Text) & _
                "/cn="
    
    sContent = "##~This file was generated by AsciiEXM for use with ExMerge.exe" & vbNewLine
    
    sContent = sContent & sTemplate & sAccount & vbNewLine
    
    Output_2_List List2, sAccount & " account was added to mailbox.txt."
    
    SetupMailBoxTxt_PerUser = sContent
End Function

Function funCreate_Exmerge_ini_PerUser() As Boolean

    funCreate_Exmerge_ini_PerUser = funCreatefile(CreateExMergeINIPath, SetupExMergeInI)
        
    If funCreate_Exmerge_ini_PerUser Then
        Output_2_List List2, "INI file was created."
    Else
        Output_2_List List2, "An Error occured while attempting to create the INI file"
    End If
End Function


Sub ProcessEachAccount()
    Dim j As Integer, i As Integer
    Dim Months As Integer
    Dim MonthsDiff As Integer
    Dim MonthSplit() As String
    Dim dInterval As Double
    Dim Interval As Integer
    Dim StartDate As Date
    Dim EndDate As Date
    Dim sStartDate As String
    Dim sEndDate As String
    
    Dim Pst As String, dpst As String
    Dim PstRename As String
    Dim FileName As String
    
    
    For j = 0 To UBound(CurrentMailBoxes) - 1
        ProgressBar2.Value = ProgressBar2.Value + 1
        ProgressBar1.Value = 1
            
        With CurrentMailBoxes(j)
        
            Pst = Trim(txtPSTPath.Text) & "\" & .Alias
                    
            Output_2_List List5, .Alias
            
            'mailbox file was created with the current mailbox
            If funCreate_Mailboxes_PerUser(.Alias) Then
            
                If chkPSTLimit_Logical.Value = 0 Then
                                
                    Pst = Pst & ".pst"
                    
                    'no limit control
                    ProgressBar1.Value = 2
                    If funCreate_Exmerge_ini_PerUser Then
                        
                            
                        ProgressBar1.Value = 3
                        If ProcessBackup Then
                            'pst file was created
                            ProgressBar1.Value = 4
                            ProcessOptions_PerUser j
                        Else
                            'there was an error with creating the pst file
                        End If
                    Else
                    
                    End If
                    ProgressBar1.Value = 5
                    
                ElseIf chkPSTLimit_Logical.Value = 1 Then
                'ElseIf LCase(Trim(cboPSTLimitOptions.Text)) = "logical" And .Size > Trim(txtLimit_Logical_GreaterThan.Text) Then
                    
                    'If optLimit_Logical_ByMonths.value = True Then
                        'Months = CInt(Trim(cboLimit_Logical_Months.Text))
                    'ElseIf optLimit_Logical_Mailbox_Size.value = True Then
                    
                    '''work on
                        'Months = CInt(.Size / 1000) 'turns it into approx GBs
                    'End If
                    
                    Months = CInt(Trim(cboLimit_Logical_Months.Text))
                    MonthsDiff = DateDiff("m", Date, .DateCreated)
                    
                    dInterval = MonthsDiff / Months
                    
                    MonthSplit = Split(dInterval, ".")
                    If UBound(MonthSplit) > 0 Then
                        Interval = Abs(CInt(MonthSplit(0)))
                    Else
                    'Interval = CInt(MonthsDiff / Months)
                        Interval = Abs(CInt(dInterval))
                    End If
                    
                    StartDate = .DateCreated
                    
                    ProgressBar1.Value = 2
                    For i = 0 To Interval
                        If i > 0 Then
                            StartDate = DateAdd("m", CDbl(Months), StartDate)
                        End If
                        
                        If i < Interval Then
                            EndDate = DateAdd("m", CDbl(Months), StartDate)
                        Else
                            EndDate = Format(Now, "mm/dd/yyyy")
                        End If
                        
                        Me.chkDateRange.Value = 1
                        txtDR_StartDate.Text = CStr(StartDate)
                        txtDR_StartTime.Text = "01:00:00"
                        txtDR_EndDate.Text = CStr(EndDate)
                        txtDR_EndTime.Text = "23:59:59"
                        
                        
                        If funCreate_Exmerge_ini_PerUser Then
                        
                            If ProcessBackup Then
                            
                                'pst file was created
                                If FileExists(Pst & ".pst") Then
                                                                   
                                    sStartDate = Replace(CStr(StartDate), "/", "-")
                                    sEndDate = Replace(CStr(EndDate), "/", "-")
                                    PstRename = Pst & "_" & sStartDate & "_to_" & sEndDate
                                    
                                    'rename file with daterange and session
                                    If Rename_File(Pst & ".pst", PstRename & ".pst") Then
                                        .FileName = PstRename
                                    
                                        ProcessOptions_PerUser j
                                    Else
                                        .FileName = Pst
                                        'error renaming
                                        i = Interval + 1
                                    End If
                                    
                                End If
                                
                            Else
                                'there was an error with creating the pst file
                            End If
                            
                        Else
                        
                        End If
                        
                    Next i
                    ProgressBar1.Value = 5
                           
                End If
                
                
            Else
                'mailbox file was not created properly
            End If
            
        End With
        
        
        'halt everything
        If bHaltProcess = True Then
            j = UBound(CurrentMailBoxes)
            Exit Sub
        End If
        
        
    Next j
    
End Sub

Function Compress_File2File(sFrom As String, sTo As String) As Boolean

    Compress_File2File = False

    If FileExists(sFrom) Then
        If ZipAFile(sTo, sFrom) Then
            Compress_File2File = True
        End If
    End If

End Function

Sub ClearFlaggedAccounts()
    Dim i As Integer, j As Integer
        
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
        
        ReDim CurrentMailBoxes(ltotalcount)
        j = 0
        For i = 0 To .MailBoxCount - 1
            .MBX(i).Flagged_2B_Processed = False
        Next i
        
    End With

End Sub

Sub CopyOnlyFlaggedAccounts()
    Dim i As Integer, j As Integer
        
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
        
        ReDim CurrentMailBoxes(ltotalcount)
        j = 0
        For i = 0 To .MailBoxCount - 1
            
            If .MBX(i).Flagged_2B_Processed = True Then
                CurrentMailBoxes(j).Alias = .MBX(i).Alias
                CurrentMailBoxes(j).Size = .MBX(i).Size
                CurrentMailBoxes(j).DateCreated = .MBX(i).DateCreated
                j = j + 1
            End If
            
        Next i
        
    End With

End Sub

Sub CycleThroughTargetAccounts()
    Dim i As Integer
        
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
            
        For i = 0 To .MailBoxCount - 1
        
            DoEvents
            If bHaltProcess = True Then Exit Sub
            
            
            If .MBX(i).Flagged_2B_Processed = True Then
                ProcessOptions i
            End If
            
        Next i
    
        If LCase(Trim(cboCompressionOption.Text)) <> "none" Then
        
            SendingFile
            
        End If
        
        
        
    End With


End Sub

Sub CycleThroughFiles()
    Dim i As Integer
        
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
        
        Output_2_List List2, "Cycling through options."
    
        For i = 0 To .MailBoxCount - 1
        
            DoEvents
            If bHaltProcess = True Then Exit Sub
            
            
            If .MBX(i).Flagged_2B_Processed = True Then
                ProcessOptions i
            End If
            
        Next i
    
        If LCase(Trim(cboCompressionOption.Text)) <> "none" Then
        
            SendingFile
            
        End If
        
        
        
    End With
    
End Sub

Function SendingFile_Account(sFrom As String, sTo As String)
    Dim tmpMsg As String
    
    ' sending
    If LCase(Trim(cboSendType.Text)) = "copy" Then
                
        tmpMsg = "Options->Send->CopyFile(" & sFrom & ")"
        Output_System_Messages_Caption Label10, tmpMsg
                    
        If Copy_File(sFrom, sTo) Then
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
        Else
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
        End If
                
    ElseIf LCase(Trim(cboSendType.Text)) = "move" Then
                
        tmpMsg = "Options->Send->MoveFile(" & sFrom & ")"
        Output_System_Messages_Caption Label10, tmpMsg
                    
        If Copy_File(sFrom, sTo) Then
            DelFile sFrom
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
        Else
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
        End If
                
    End If
End Function

Function SendingFile()

    Dim tmpMsg As String
            ' sending
            If LCase(Trim(cboSendType.Text)) = "copy" Then
                
                tmpMsg = "Options->Send->CopyFile(" & CurrentFile & ")"
                Output_System_Messages_Caption Label10, tmpMsg
                    
                If Copy_File(CurrentFile, Trim(txtMoveToFolder.Text)) Then
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
                Else
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
                End If
                
            ElseIf LCase(Trim(cboSendType.Text)) = "move" Then
                
                tmpMsg = "Options->Send->MoveFile(" & CurrentFile & ")"
                Output_System_Messages_Caption Label10, tmpMsg
                    
                If Copy_File(CurrentFile, Trim(txtMoveToFolder.Text) & "\" & Trim(txtDotZipFileName.Text)) Then
                    DelFile CurrentFile
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
                Else
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
                End If
                
            End If
End Function

Function CleanUpFiles(Pst As String)

    Dim tmpMsg As String
    ' clean up
    If chkDeleteAfterZip.Value = 1 Then
                    
        tmpMsg = "Options->Delete_PST_File(" & Pst & ")"
        Output_System_Messages_Caption Label10, tmpMsg
                        
        ' delete pst file
        If DelFile(Pst) Then
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
        Else
            Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
        End If
                            
    End If
End Function

Function ProcessOptions(iMailBoxIndex As Integer)
    Dim Pst As String
    Dim tmpMsg As String
    
    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs).MBX(iMailBoxIndex)
    
        'Output_2_List List5, vbTab & "->" & .Alias
        Output_2_List List2, vbTab & "->" & .Alias
    
        Pst = Trim(txtPSTPath.Text) & "\" & .Alias & ".pst"
        CurrentFile = Pst
            
        If Exists_fso(Pst) Then
            
            ' encrypt
            '''''''''''''''''''
                
            ' compress
            If LCase(Trim(cboCompressionOption.Text)) = "winzip" Then
            
                tmpMsg = "Options->Compress->WinZip(" & Pst & ")"
                Output_System_Messages_Caption Label10, tmpMsg
                
                CurrentFile = ZipFileSettings(Pst)
                
                If CurrentFile = "" Then
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = Error!"
                Else
                    Output_2_List List2, vbTab & vbTab & "->" & tmpMsg & " = successful!"
                End If
                
            ElseIf LCase(Trim(cboCompressionOption.Text)) = "none" Then
            
                SendingFile
                
            End If
            
            
            CleanUpFiles Pst
                               
                
        Else
            Output_2_List List2, vbTab & vbTab & "-> File: '" & Pst & "' does not exist"
            
        End If
        
    End With
        
End Function

Function Build_Src_PST(sAccount As String) As String
    Build_Src_PST = ""
    
    Build_Src_PST = Trim(txtPSTPath.Text) & "\" & sAccount & ".pst"
End Function

Function Build_Dest_WinZip(sAccount As String) As String
    Build_Dest_WinZip = ""
    
    If optProcessIndividually.Value = True Then
        Build_Dest_WinZip = Trim(txtDotZipFilePath.Text) & "\" & sAccount & ".zip"
    ElseIf optProcessAsGroup.Value = True Then
        Build_Dest_WinZip = Trim(txtDotZipFilePath.Text) & "\" & Trim(txtDotZipFileName.Text)
    End If
End Function

Function ZipFileSettings_Account(sAccount As String) As String
    Dim zip As String
    Dim Pst As String
    
    ZipFileSettings_Account = ""
    
    Pst = Build_Src_PST(sAccount)
    zip = Build_Dest_WinZip(sAccount)
    
    'If FileExists(zip) Then
        If ZipAFile(zip, Pst) Then ZipFileSettings_Account = zip
    'End If
End Function

Function ZipFileSettings(Pst As String) As String
    Dim zip As String
    
    ZipFileSettings = ""
    
    zip = Trim(txtDotZipFilePath.Text) & "\" & Trim(txtDotZipFileName.Text)
    
    'If FileExists(zip) Then
        If ZipAFile(zip, Pst) Then ZipFileSettings = zip
    'End If
End Function

'Sub ZipAllFiles()
'    Dim i As Integer
'    Dim zip As String, pst As String
    
'    zip = Trim(txtDotZipFilePath.Text) & "\" & Trim(txtDotZipFileName.Text)
    
'    For i = 0 To List1.ListCount
'        pst = Trim(txtPSTPath.Text) & "\" & Trim(List1.List(i)) & ".pst"
'        ZipAFile zip, pst
'        If chkDeleteAfterZip.value = 1 Then DelFile pst 'delete pst file
'    Next i
'End Sub


Sub ThreadCount(iItems As Integer)
    If chkDefaultThreads.Value = 1 Then
        If iItems < 5 Then
            cboThreadCount.ListIndex = 0
        ElseIf iItems < 25 Then
            cboThreadCount.ListIndex = 1
        ElseIf iItems < 50 Then
            cboThreadCount.ListIndex = 2
        ElseIf iItems < 100 Then
            cboThreadCount.ListIndex = 3
        ElseIf iItems >= 100 Then
            cboThreadCount.ListIndex = 4
        End If
    End If
End Sub





Private Sub cmdSettings_Next_Click()
    ShowFrame 3
End Sub

Private Sub cmdSource_Back_Click()
    ShowFrame 0
End Sub

Private Sub cmdSource_Exit_Click()
    Unload Me
End Sub

Private Sub cmdSource_Next_Click()
    ShowFrame 4
End Sub

Private Sub Command1_Click()
    If Me.Height = fHeight + 1747 + 880 Then
        Me.Height = fHeight + 1747 + 880 + 1100
    Else
        Me.Height = fHeight + 1747 + 880
    End If
    
End Sub

'Private Sub cmdResults_Click()
'    Call Shell("notepad " & Trim(frmExMerge.txtPSTPath.Text) & "\AdminEx.Log", vbNormalFocus)
'End Sub


Private Sub Command11_Click()
    txtMailBoxPath.Text = SHFolder
End Sub

Private Sub Command12_Click()
    txtLogPath.Text = SHFolder
End Sub

Private Sub Command13_Click()
    INI_Default
    MailBox_Default
End Sub

Private Sub Command14_Click()
    txtINIPath.Text = SHFolder
End Sub













Private Sub Command17_Click()
    txtRootPath.Text = SHFolder
End Sub





Private Sub Command2_Click()
    txtExMergePath.Text = SHFolder
End Sub


Private Sub Command22_Click()
    txtMoveToFolder.Text = SHFolder
End Sub


Private Sub Command3_Click()
    txtPSTPath.Text = SHFolder
End Sub






Private Sub Command32_Click()
    txtDotZipFilePath.Text = SHFolder
End Sub



Private Sub Command34_Click()
    ShowFrame 5
End Sub

Private Sub Command35_Click()
    ShowFrame 4
End Sub

Private Sub Command36_Click()
    ShowFrame 3
End Sub



Private Sub Command38_Click()
    ShowFrame 1
End Sub

Private Sub Command39_Click()
    ShowFrame 0
End Sub












Sub prcDiscovery()

    Option3.Enabled = False
    Option4.Enabled = False
    txtLocalExchangeServer.Enabled = False
    cboInitExchangeServers.Enabled = False
    
    
    ' Grab Disk Drive Information
    If chkDiscoverDiskdrives.Value = 1 Then
        Label10.Caption = "Disk Drive Discovery ..."
        Pause 0.5
        mod_Drives.Drives2Struct
        Output_2_List List4, "Disk Drives Discovered."
        Label50.Caption = ""
    End If
    
    
    If Option3.Value = True Then
        Label10.Caption = "Discovering " & Trim(txtLocalExchangeServer.Text)
        GetExchangeServer_Info Trim(txtLocalExchangeServer.Text), 0
        Pause 1
        
        Option3.Enabled = True
        txtLocalExchangeServer.Enabled = True
    Else
        
        If LCase(Trim(cboInitExchangeServers.Text)) = "all" Then
            Label10.Caption = "Discovering all exchange servers on the domain."
            GetAllExchangeInfo
        Else
            Label10.Caption = "Discovering " & Trim(cboInitExchangeServers.Text)
            GetExchangeServer_Info LCase(Trim(cboInitExchangeServers.Text)), 0
            Pause 1
        End If
        
    End If
    'Show_Servers
        
    Label10.Caption = "Discovery complete."
    List4.AddItem "Complete."
    List4.ListIndex = List4.ListCount - 1
    Option4.Enabled = True
    cboInitExchangeServers.Enabled = True
    Label10.Caption = ""
    Label50.Caption = ""
    Label60.Caption = ""
    cmdIntro_Discovery.Caption = "Re-Discover"
    
    Show_TreeView
    
    ShowFrame 1
End Sub





Private Sub Form_Load()
    
    Init
    Me.Show
    
    Label10.Caption = "Loading Settings..."
    Pause 1
    
    LoadAll
    
    bInitStage = False
    
End Sub

Sub LoadAll()
    
    'do these later
    ' Check Resources (cdo, systools, wmi, winzip, etc)
    CompressionExists
    
    Main.Sys.AD_Exists = ADExists
    
    
    ' normally used
    Label10.Caption = "Loading Defaults..."
    LoadDefaultSettings
    'LoadResources
    
    '''' Grab All domain and ad user Information
    ''''Label10.Caption = "Discovering Domain objects ..."
    ''''Pause 1
    ''''ListDomains
    
    ''''GetADUserInfo "charty"
    
    ' Grab All Exchange Information
    ReDim Servers(0)
    If Main.Sys.AD_Exists Then
        Label10.Caption = "Discovering Exchange Servers ..."
        Pause 0.5
        InitExchangeServers_Only
        Output_2_List List4, "Exchange Servers Discovered."
        
        Pause 1
        Label10.Caption = "Discovery and Loading are Complete."
        Output_2_List List4, "Discovery and Loading are Complete."
        
        cmdIntro_Discovery.Enabled = True
    Else
        
        ' Fake Exchange Information for testing
        Fake_ExchangeInfo
        Show_TreeView
    
        Label10.Caption = "Active Directory was not detected.  Please restart on a computer that is on a domain."
        Pause 4
    End If
    
    'Label10.Caption = "Loading Exchange objects ..."
    'Pause 1
    
    
    
    'Label10.Caption = "Discovery and Loading are Complete."
    
        
    cboInitExchangeServers.Enabled = True
    cmdIntro_Next.Enabled = True
    cmdIntro_Exit.Enabled = True
    Frame_Intro.Enabled = True
End Sub

Sub InitExchangeServers_Only()

    GetExchangeServers_Only
    
    If LocalExchangeServerAvailable = True Then
        Option3.Enabled = True
        Option3.Value = True
        Option4.Value = False
    Else
        Option3.Enabled = False
        Option3.Value = False
        Option4.Value = True
    End If
    
    DisplayExchangeServer_Only
    
End Sub

Function LocalExchangeServerAvailable() As Boolean

    LocalExchangeServerAvailable = False
    
    If UBound(Servers) > 0 Then
        Dim i As Integer
        
        For i = 0 To UBound(Servers) - 1
            If Servers(i).name = Main.Sys.ComputerName Then
                txtLocalExchangeServer.Text = Servers(i).name
                LocalExchangeServerAvailable = True
                Exit Function
            End If
        Next i
        
    End If
    
End Function

Function DisplayExchangeServer_Only()
    Dim i As Integer
    
    Label37.Caption = UBound(Servers) & " Exchange servers were found."
    cboInitExchangeServers.Text = "All"
    cboInitExchangeServers.AddItem "All"
    
    For i = 0 To UBound(Servers) - 1
        cboInitExchangeServers.AddItem Servers(i).name
    Next i

End Function



Sub Reset_MailBox_MSFlexGrid()

   
    With MSFlexGrid1
        .Cols = 6
        .AllowBigSelection = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(0) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 900
        .ColWidth(5) = 600
    
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "User"
        .TextMatrix(0, 2) = "Alias"
        .TextMatrix(0, 3) = "DateCreated"
        .TextMatrix(0, 4) = "Msg Count"
        .TextMatrix(0, 5) = "MBs"
       
    End With
    
End Sub
Sub Reset_Drives_MSFlexGrid()

    With MSFlexGrid2
      .Cols = 6
    
      .ColWidth(0) = 0
      .ColWidth(1) = 1000
      .ColWidth(2) = 2000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
    
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "DriveLetter"
      .TextMatrix(0, 2) = "Path"
      .TextMatrix(0, 3) = "Type"
      .TextMatrix(0, 4) = "Free Space"
      .TextMatrix(0, 5) = "Total Space"
       
    End With
    
End Sub



Sub LoadDefaultSettings()
    MailBox_Default
    ExMergeOptions_Default
    'Account_Default
    INI_Default
End Sub

Sub Init()
    
    'Option3.Enabled = False
    'txtLocalExchangeServer.Enabled = False
    'Option4.Enabled = False
    'cboInitExchangeServers.Enabled = False
    
    SysInfo
    txtLocalExchangeServer.Text = Main.Sys.ComputerName
    txtTotalCombinedCount.Text = 0
    chkDiscoverOnlineDBsOnly.Value = 1
    cboLimit_Logical_Months.Enabled = False
    bInitStage = True
    optProcessIndividually.Value = True
    txtDotZipFileName.Enabled = False
    optLimit_Logical_ByMonths.Value = True
    txtLimit_Logical_GreaterThan.Text = "1.5"
    
    Frame_Intro.Enabled = False
    cboInitExchangeServers.Enabled = False
    cmdIntro_Next.Enabled = False
    cmdIntro_Exit.Enabled = False
    cmdIntro_Discovery.Enabled = False
    
    cmdProcess_Halt.Enabled = False
    cmdProcess_Start.Enabled = False
    
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
    F_Header.Top = 0
    F_Header.Left = 0
    F_Header.Width = 10155
    F_Header.Height = 1155
        
    fTop = 1080
    fLeft = -60
    fWidth = 8115 '10155 + 240
    fHeight = 4215 '5535
    
    Frame2.Top = fHeight + fTop + 20
    Frame2.Left = fLeft
    Frame2.Width = fWidth
    
    Label4.Top = Frame2.Top + 700
    Label4.Left = 1920
    Label10.Top = Frame2.Top + 700
    Label10.Left = 2580
    Label50.Top = Frame2.Top + 700
    Label50.Left = 60
    Label60.Top = Frame2.Top + 700
    Label60.Left = 960
    
    List4.Top = Frame2.Top + 700 + 360
    List4.Left = 60
    List4.Width = fWidth - 300
    
    Frame_Menu.Height = fHeight
    Frame_Menu.Width = 2655
    
    Me.Width = fWidth - 100
    Me.Height = fHeight + 1747 + 880
    
    Reset_MailBox_MSFlexGrid
    Reset_Drives_MSFlexGrid
    
    SetFrame Frame_Intro
    SetFrame Frame_Source
    'SetFrame Frame_Destination
    SetFrame Frame_INI
    SetFrame Frame_AdvancedOptions
    SetFrame Frame_Process
    
    
    setCmd_Exit cmdIntro_Exit
    setCmd_Next cmdIntro_Next
        
    setCmd_Exit cmdSource_Exit
    setCmd_Next cmdSource_Next
    setCmd_Back cmdSource_Back
        
    setCmd_Exit cmdSettings_Exit
    setCmd_Next cmdSettings_Next
    setCmd_Back cmdSettings_Back
        
    setCmd_Exit cmdOptions_Exit
    setCmd_Next cmdOptions_Next
    setCmd_Back cmdOptions_Back
        
    setCmd_Exit cmdProcess_Exit
    setCmd_Back cmdProcess_Back

    ShowFrame 0
    
End Sub

Sub setCmd_Exit(c As CommandButton)
    c.Width = 915
    c.Height = 435
    c.Left = 6900
    c.Top = 3600
End Sub

Sub setCmd_Next(c As CommandButton)
    c.Width = 915
    c.Height = 435
    c.Left = 5340
    c.Top = 3600
End Sub

Sub setCmd_Back(c As CommandButton)
    c.Width = 915
    c.Height = 435
    c.Left = 180
    c.Top = 3600
End Sub

Sub SetFrame(f As Frame)
    f.Top = fTop
    f.Left = fLeft
    f.Width = fWidth
    f.Height = fHeight
End Sub

Sub ShowFrame(iFrame)
    
    If iFrame = 0 Then
        Frame_Intro.ZOrder vbBringToFront
    ElseIf iFrame = 1 Then
        Frame_Source.ZOrder vbBringToFront
    'ElseIf iFrame = 2 Then
        'Frame_Destination.ZOrder vbBringToFront
    ElseIf iFrame = 3 Then
        Frame_AdvancedOptions.ZOrder vbBringToFront
    ElseIf iFrame = 4 Then
        Frame_INI.ZOrder vbBringToFront
    ElseIf iFrame = 5 Then
        Frame_Process.ZOrder vbBringToFront
    End If
    
End Sub

Sub Fake_ExchangeInfo()

    'Main.Current.StorageGroup

    With Main.Exch
    
    .ServerCount = 3
    ReDim .Svrs(.ServerCount)
    .Svrs(0).name = "SUSAMAIL"
    .Svrs(1).name = "NASDATA"
    .Svrs(2).name = "SNYEDGE"
        
        
        .Svrs(0).StorageGroupCount = 2
        ReDim .Svrs(0).SG(.Svrs(0).StorageGroupCount)
        
        .Svrs(0).SG(0).name = "First"
        
            .Svrs(0).SG(0).MailBoxStoreDBCount = 3
            ReDim .Svrs(0).SG(0).MBSDB(.Svrs(0).SG(0).MailBoxStoreDBCount)
            
                .Svrs(0).SG(0).MBSDB(0).name = "acc"
                
                    .Svrs(0).SG(0).MBSDB(0).MailBoxCount = 2
                    ReDim .Svrs(0).SG(0).MBSDB(0).MBX(.Svrs(0).SG(0).MBSDB(0).MailBoxCount)
                        .Svrs(0).SG(0).MBSDB(0).MBX(0).Alias = "alon"
                        .Svrs(0).SG(0).MBSDB(0).MBX(1).Alias = "bvon"
                        
                .Svrs(0).SG(0).MBSDB(1).name = "exec"
                .Svrs(0).SG(0).MBSDB(2).name = "MBX"
            
        .Svrs(0).SG(1).name = "Second"
        
            .Svrs(0).SG(1).MailBoxStoreDBCount = 3
            ReDim .Svrs(0).SG(1).MBSDB(.Svrs(0).SG(1).MailBoxStoreDBCount)
            
                .Svrs(0).SG(1).MBSDB(0).name = "meat"
                .Svrs(0).SG(1).MBSDB(1).name = "meat2"
                .Svrs(0).SG(1).MBSDB(2).name = "meat3"
                
                    .Svrs(0).SG(1).MBSDB(1).MailBoxCount = 2
                    ReDim .Svrs(0).SG(1).MBSDB(1).MBX(.Svrs(0).SG(1).MBSDB(1).MailBoxCount)
                        .Svrs(0).SG(1).MBSDB(1).MBX(0).Alias = "2alon"
                        .Svrs(0).SG(1).MBSDB(1).MBX(1).Alias = "2bvon"
            
            
        .Svrs(1).StorageGroupCount = 3
        ReDim .Svrs(1).SG(.Svrs(1).StorageGroupCount)
        .Svrs(1).SG(0).name = "nas-SG1"
        .Svrs(1).SG(1).name = "nas-SG2"
        .Svrs(1).SG(2).name = "nas-SG3"
            
            
        .Svrs(2).StorageGroupCount = 3
        ReDim .Svrs(2).SG(.Svrs(2).StorageGroupCount)
        .Svrs(2).SG(0).name = "edge-SG1"
        .Svrs(2).SG(1).name = "edge-SG1"
        .Svrs(2).SG(2).name = "edge-SG1"
        
        
    End With
        
End Sub


Sub Show_Servers()
    Dim iObj As Integer
    
    cboEmailServers.Clear
    cboStorageGroups.Clear
    cboMailBoxStores.Clear
    List7.Clear
    Label56.Caption = " (" & Main.Exch.ServerCount & ") "
    
    For iObj = 0 To Main.Exch.ServerCount - 1
    
        If iObj = 0 Then
            Main.Current.Server = Main.Exch.Svrs(iObj).name
            cboEmailServers.Text = Main.Exch.Svrs(iObj).name
        End If
            
        cboEmailServers.AddItem Main.Exch.Svrs(iObj).name
    
    Next iObj
End Sub

Function GetServerIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetServerIndex = -i
    
    For i = 0 To Main.Exch.ServerCount - 1
        If Main.Exch.Svrs(i).name = sTarget Then
            GetServerIndex = i
            Exit Function
        End If
    Next i
End Function

Sub Show_StorageGroups(sTarget As String)
    Dim iObj As Integer, iSvr As Integer
    
    Main.Current.Server = GetServerIndex(sTarget)
    
    cboStorageGroups.Clear
    cboMailBoxStores.Clear
    List7.Clear
    Label57.Caption = " (" & Main.Exch.Svrs(Main.Current.Server).StorageGroupCount & ") "
    
    For iObj = 0 To Main.Exch.Svrs(Main.Current.Server).StorageGroupCount - 1
                
        If iObj = 0 Then
            cboStorageGroups.Text = Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
            'Main.Current.StorageGroup = Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
        End If
            
        cboStorageGroups.AddItem Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
    
    Next iObj
End Sub


Function GetStorageGroupIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetStorageGroupIndex = -i
    
    For i = 0 To Main.Exch.Svrs(Main.Current.Server).StorageGroupCount - 1
        If Main.Exch.Svrs(Main.Current.Server).SG(i).name = sTarget Then
            GetStorageGroupIndex = i
            Exit Function
        End If
    Next i
End Function

Sub Show_MailBoxStoreDBs(sTarget As String)
    Dim iObj As Integer
    
    Main.Current.StorageGroup = GetStorageGroupIndex(sTarget)
    
On Error GoTo Err:

    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup)
        
        cboMailBoxStores.Clear
        List7.Clear
        Label58.Caption = " (" & .MailBoxStoreDBCount & ") "
        
        For iObj = 0 To .MailBoxStoreDBCount - 1
        
            If iObj = 0 Then
                cboMailBoxStores.Text = .MBSDB(iObj).name
                'Main.Current.MailBoxStoreDBs = .MBSDB(iObj).name
            End If
                
            cboMailBoxStores.AddItem .MBSDB(iObj).name
        
        Next iObj
    End With
    
    Exit Sub
Err:
    Exit Sub
End Sub

Sub Show_TreeView()
    Dim iSrv As Integer, iSG As Integer, iMbx As Integer
    Dim tmp_MBSDB As String
    Dim iMBSDBCount As Integer
    Dim sKey As String
    
    '''''''''''
    ''''''''''work on
    '''''''''''
    
    iMBSDBCount = 0
    For iSrv = 0 To Main.Exch.ServerCount - 1
    
        With Main.Exch.Svrs(iSrv)
        
            'add server
            If .name <> "" And .StorageGroupCount > 0 Then
            Set ndSrv = TreeView1.Nodes.Add(, , , .name & "(" & .StorageGroupCount & ")")
            ndSrv.Expanded = True
            
            'inspect SGs
            For iSG = 0 To .StorageGroupCount - 1
                'add SG
                If .SG(iSG).name <> "" Then
                Set ndSG = TreeView1.Nodes.Add(ndSrv, tvwChild, , .SG(iSG).name & "(" & .SG(iSG).MailBoxStoreDBCount & ")")
                ndSG.Expanded = True
            
                'inspect MBXs
                With .SG(iSG)
                    For iMbx = 0 To .MailBoxStoreDBCount - 1
                        'add MBX
                        
                        If .MBSDB(iMbx).name <> "" Then
                        
                        'create key
                        sKey = iSrv & "-" & iSG & "-" & iMbx
                        
                        tmp_MBSDB = tmp_MBSDB & .MBSDB(iMbx).name & ","
                        
                        
                        Set ndMBX = TreeView1.Nodes.Add(ndSG, tvwChild, sKey, .MBSDB(iMbx).name & "(" & .MBSDB(iMbx).MailBoxCount & ")")
                        ndMBX.Expanded = True
                        End If
                    
                    Next iMbx
                End With
                End If
                
            Next iSG
            End If
            
        End With
    
    Next iSrv
    
    If tmp_MBSDB = "" Then
        ReDim DBs(0)
    Else
        DBs = Split(tmp_MBSDB, ",")
    End If
    
End Sub

Function GetMailBoxStoreDBIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetMailBoxStoreDBIndex = -i
    
On Error GoTo Err:
    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup)
    
        For i = 0 To .MailBoxStoreDBCount - 1
            If .MBSDB(i).name = sTarget Then
                GetMailBoxStoreDBIndex = i
                Exit Function
            End If
        Next i
    
    End With
    
Err:
    
End Function

Sub Show_MailBoxes(sTarget As String)
    Dim iObj As Integer
    
    'Main.Current.MailBoxStoreDBs = GetMailBoxStoreDBIndex(sTarget)
    
On Error GoTo Err:
    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
    
        MSFlexGrid1.Rows = .MailBoxCount + 1
        List7.Clear
        Label58.Caption = " (" & .MailBoxCount & ") "
        
        For iObj = 0 To .MailBoxCount - 1
                    
            List7.AddItem .MBX(iObj).Alias & " | " & .MBX(iObj).MessageCount & " | " & .MBX(iObj).Size
            
            MSFlexGrid1.TextMatrix(iObj + 1, 0) = iObj
            MSFlexGrid1.TextMatrix(iObj + 1, 1) = .MBX(iObj).ADInfo.FullName
            MSFlexGrid1.TextMatrix(iObj + 1, 2) = .MBX(iObj).Alias
            MSFlexGrid1.TextMatrix(iObj + 1, 3) = .MBX(iObj).DateCreated
            MSFlexGrid1.TextMatrix(iObj + 1, 4) = .MBX(iObj).MessageCount
            MSFlexGrid1.TextMatrix(iObj + 1, 5) = .MBX(iObj).DisplaySize
                        
        Next iObj
    End With
    
    Exit Sub
Err:
    Exit Sub
End Sub

Sub LoadResources()
    ''' get Exchange Server list
    'GetExchangeServers
    'GetExchangeServers2 cboEmailServers
    
End Sub

Sub Check_Winzip()
    'chkUseWinZip.Enabled = False
    'chkUseWinZip.Value = 0
    'chkDeleteAfterZip.Enabled = False
    'chkDeleteAfterZip.Value = 0
    'txtDotZipFilePath.Enabled = False
    'Command32.Enabled = False
    'txtDotZipFilePath.Enabled = False
    'If Exists_fso("") Then
    '    chkUseWinZip.Enabled = True
    '    chkDeleteAfterZip.Enabled = True
    '    txtDotZipFilePath.Enabled = True
    '    Command32.Enabled = True
    '    txtDotZipFilePath.Enabled = True
    'End If
End Sub


'example: sFilter = "All Files .* |*.*"
Function funFileChoice(cdg As CommonDialog, sFilter As String) As String
    cdg.Filter = sFilter
    cdg.ShowOpen
    
    If cdg.FileName = "" Then Exit Function
    
    funFileChoice = cdg.FileName
End Function


'example: sFilter = "All Files .* |*.*"
Function funFolderChoice(cdg As CommonDialog, sFilter As String) As String
    cdg.Filter = sFilter
    cdg.ShowOpen
    
    If cdg.FileName = "" Then Exit Function
    
    funFolderChoice = cdg.FileName
End Function


Sub LoadTestVariables()
    With Me
        '.cboEmailServers.Text = "susamail"
        .txtExMergePath.Text = "c:\program files\exchsrvr\bin"
        .txtPSTPath.Text = "c:\chris"
        .txtOrganization.Text = "SIBONEYUSA CORP"
        .txtGroup.Text = "EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)"
        .txtCN1.Text = "RECIPIENTS"
        '.List1.AddItem "NGUPTA"
    End With
End Sub

'Sub Account_Default()
 '   With Me
 '       .List1.AddItem "NGUPTA"
 '       .List1.AddItem "CCANTO"
 '       .List1.AddItem "BFORD"
 '       .List1.AddItem "AMINOGUE"
 '   End With
'End Sub

'Sub Account_ClearForm()
'    With Me
'        .List1.Clear
'        .txtEmailAccount.Text = ""
'    End With
'End Sub


Sub INI_ClearForm()
    With Me
        .optProcess_AsGroup = True
        .txtRootPath.Text = ""
        .chkUseRoot.Value = 0
        .txtINIPath.Text = ""
        .txtINIFileName.Text = ""
        '.cboEmailServers.Text = ""
        .txtExMergePath.Text = ""
        .txtPSTPath.Text = ""
        .txtLogPath.Text = ""
        .cmbLogLevel.ListIndex = 0
        .chkUserData.Value = 0
        .chkDumpster.Value = 0
        .chkFolderRules.Value = 0
        .chkFolderData.Value = 0
    End With
End Sub

Sub ExMergeOptions_Default()
    With Me
        .txtDotZipFilePath.Text = "C:\backup\exchange\PST"
        .txtMoveToFolder.Text = "C:\backup\exchange\TEMP"
        .cboThreadCount.ListIndex = 0
        .chkShowInterface.Value = 1
        .chkDefaultThreads = 1
    End With
End Sub

Sub INI_Default()
    With Me
        .txtINIPath.Text = "C:\backup\exchange"
        .txtINIFileName.Text = "ExMerge.ini"
        .txtExMergePath.Text = "c:\program files\exchsrvr\bin"
        .txtPSTPath.Text = "C:\backup\exchange\PST"
        .txtLogPath.Text = "C:\backup\exchange\Log"
        .cmbLogLevel.ListIndex = 2
        .chkUserData.Value = 1
        .chkDumpster.Value = 1
        .chkFolderRules.Value = 1
        .chkFolderData.Value = 1
        .chkUseRoot.Value = 1
        .txtRootPath.Text = "C:\backup\exchange"
    End With
End Sub
Sub MailBox_ClearForm()
    With Me
        .txtMailBoxPath.Text = ""
        .txtMailBoxFileName.Text = ""
        .txtOrganization.Text = ""
        .txtGroup.Text = ""
        .txtCN1.Text = ""
    End With
End Sub
Sub MailBox_Default()
On Error Resume Next
    With Me
        .optProcess_AsGroup.Value = True
        .txtMailBoxPath.Text = "C:\backup\exchange"
        .txtMailBoxFileName.Text = "mailboxes.txt"
        
        If Main.Sys.AD_Exists Then
            .txtOrganization.Text = Get_OrganizationName '"SIBONEYUSA CORP"
            .txtGroup.Text = Get_AdministrativeGroup '"EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)"
        Else
            .txtOrganization.Text = ""
            .txtGroup.Text = ""
        End If
        
        .txtCN1.Text = "RECIPIENTS"
    End With
End Sub
















Private Sub MSFlexGrid1_Click()
    HighlightRow MSFlexGrid1.Row
End Sub

Public Function HighlightRow(iRow As Integer)
    Dim iCpt As Integer
    Dim ipass As Integer

    ipass = 0
    For iCpt = 0 To MSFlexGrid1.Cols - 1
        With MSFlexGrid1
            .Row = iRow
            .Col = iCpt
            If .CellForeColor = 0 Then
                If ipass = 0 Then
                    GetAccountsSelectedTotalSize iRow - 1, True
                    GetAccountsSelectedCount True
                End If
                .CellForeColor = &H80000003
            Else
                If ipass = 0 Then
                    GetAccountsSelectedTotalSize iRow - 1, False
                    GetAccountsSelectedCount False
                End If
                .CellForeColor = 0
            End If
        End With
        ipass = ipass + 1
    Next
    
End Function

Function GetAccountsSelectedTotalSize(iIndex As Integer, bAdd As Boolean) As String

    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
        If bAdd = True Then
            ltotalsize = ltotalsize + .MBX(iIndex).Size
            .MBX(iIndex).Flagged_2B_Processed = True
        Else
            ltotalsize = ltotalsize - .MBX(iIndex).Size
            .MBX(iIndex).Flagged_2B_Processed = False
        End If
    End With
    
    txtTotalCombinedSize.Text = FormatNumber(ltotalsize, 2, vbTrue, vbTrue, vbTrue)
    
End Function

Function GetAccountsSelectedCount(bAdd As Boolean) As String

    If bAdd = True Then
        ltotalcount = ltotalcount + 1
    Else
        ltotalcount = ltotalcount - 1
    End If
    
    ThreadCount CInt(ltotalcount)
    
    If ltotalcount > 0 Then
        cmdProcess_Start.Enabled = True
        cmdProcess_Halt.Enabled = True
    Else
        cmdProcess_Start.Enabled = False
        cmdProcess_Halt.Enabled = False
    End If
    txtTotalCombinedCount.Text = ltotalcount
    
End Function

Private Sub optProcessAsGroup_Click()
    txtDotZipFileName.Enabled = True
End Sub

Private Sub optProcessIndividually_Click()
    txtDotZipFileName.Enabled = False
End Sub

Private Sub Picture1_Click()
    Main.DateRange.InputDate = Format(Date, "mm/dd/yyyy")
    Main.DateRange.Requestor = "EndDate"
    frmCalendar.Show
End Sub

Private Sub Picture2_Click()
    txtDR_StartTime.Text = Format(Time, "HH:NN:SS")
End Sub

Private Sub Picture4_Click()
    Main.DateRange.InputDate = Format(Date, "mm/dd/yyyy")
    Main.DateRange.Requestor = "StartDate"
    frmCalendar.Show
End Sub

Private Sub Picture6_Click()
    txtDR_EndTime.Text = Format(Time, "HH:NN:SS")
End Sub






Public Function funCreate_Mailboxes_txt(mg As MSFlexGrid) As Boolean

    funCreate_Mailboxes_txt = funCreatefile(CreateMailBoxTxTPath, SetupMailBoxTxt_MSFlexGrid(mg))
    
    If funCreate_Mailboxes_txt Then
        Output_2_List List2, "MailBox file was created."
    Else
        Output_2_List List2, "An Error occured while attempting to create the MailBox file"
    End If
    
End Function

Function CreateMailBoxTxTPath() As String

    Output_2_List List2, "Attempting to create MailBox file."
    CreateMailBoxTxTPath = Trim(txtMailBoxPath.Text) & "\" & Trim(txtMailBoxFileName.Text)
    
End Function

Function SetupMailBoxTxt_MSFlexGrid(mg As MSFlexGrid) As String
    Dim i As Integer
    Dim sTemplate As String, sContent As String
    Dim iTrack As Integer, iRpt As Integer
    
    sTemplate = "/o=" & Trim(txtOrganization.Text) & _
                "/ou=" & Trim(txtGroup.Text) & _
                "/cn=" & Trim(txtCN1.Text) & _
                "/cn="
    
    sContent = "##~This file was generated by AsciiEXM for use with ExMerge.exe" & vbNewLine
    
    iTrack = 0
    For iRpt = 0 To MSFlexGrid1.Rows - 1
        With MSFlexGrid1
            .Row = iRpt
            .Col = 2
            If .CellForeColor = &H80000003 Then
                With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
                    Output_2_List List5, .MBX(iRpt - 1).Alias
                    sContent = sContent & sTemplate & .MBX(iRpt - 1).Alias & vbNewLine
                    iTrack = iTrack + 1
                End With
            End If
        End With
    Next
    
    Output_2_List List2, iTrack & " mailbox accounts found."
    
    SetupMailBoxTxt_MSFlexGrid = sContent
End Function

Function SetupMailBoxTxt_listbox(lst As ListBox) As String
    Dim i As Integer
    Dim sTemplate As String, sContent As String
    
    sTemplate = "/o=" & Trim(txtOrganization.Text) & _
                "/ou=" & Trim(txtGroup.Text) & _
                "/cn=" & Trim(txtCN1.Text) & _
                "/cn="
    
    sContent = "##~This file was generated by AsciiEXM for use with ExMerge.exe" & vbNewLine
    
    For i = 0 To lst.ListCount - 1
        sContent = sContent & sTemplate & lst.List(i) & vbNewLine
    Next i
    
    SetupMailBoxTxt_listbox = sContent
End Function

Function funCreateExMergeINI() As Boolean

    funCreateExMergeINI = funCreatefile(CreateExMergeINIPath, SetupExMergeInI)
        
    If funCreateExMergeINI Then
        Output_2_List List2, "INI file was created."
    Else
        Output_2_List List2, "An Error occured while attempting to create the INI file"
    End If
End Function

Function CreateExMergeINIPath() As String

    Output_2_List List2, "Attempting to create INI file."
    CreateExMergeINIPath = Trim(txtINIPath.Text) & "\" & Trim(txtINIFileName.Text)

End Function

Function SetupExMergeInI() As String
    Dim sBody As String
    
    sBody = "; EXEMERGE.INI; This file is for use with the EXMERGE.EXE program, version 4.00 or higher." & vbNewLine
    sBody = sBody & "[EXMERGE]" & vbNewLine
    sBody = sBody & "MergeAction=0" & vbNewLine
    sBody = sBody & "SourceServerName=" & Main.Exch.Svrs(Main.Current.Server).name & vbNewLine
    sBody = sBody & "DataDirectoryName=" & Trim(txtPSTPath.Text) & vbNewLine
    sBody = sBody & "FileContainingListOfMailboxes=" & Trim(txtMailBoxPath.Text) & "\" & Trim(Me.txtMailBoxFileName.Text) & vbNewLine
    'sBody = sBody & "FileContainingListOfDatabases=C:\Program Files\Exchsrvr\bin\PRIVMDBS.TXT" & vbNewLine
    sBody = sBody & "LogFileName=" & Trim(txtLogPath.Text) & "\ExMergeX.Log" & vbNewLine
    sBody = sBody & "LoggingLevel=" & cmbLogLevel.ListIndex & vbNewLine
    If chkDumpster.Value = 1 Then sBody = sBody & "CopyDeletedItemsFromDumpster = 1" & vbNewLine
    If chkFolderRules.Value = 1 Then sBody = sBody & "CopyFolderPermissions = 1" & vbNewLine
    If chkFolderData.Value = 1 Then sBody = sBody & "CopyAssociatedFolderData = 1" & vbNewLine
    If chkUserData.Value = 1 Then sBody = sBody & "CopyUserData = 1" & vbNewLine
    
    If chkDateRange.Value = 1 Then
        If Trim(txtDR_StartDate.Text) <> "" And Trim(txtDR_EndDate.Text) <> "" And Trim(txtDR_StartTime.Text) <> "" And Trim(txtDR_EndTime.Text) <> "" Then
            sBody = sBody & "SelectMessageStartDate = " & Trim(txtDR_StartDate.Text) & " " & Trim(txtDR_StartTime.Text) & vbNewLine
            sBody = sBody & "SelectMessageEndDate = " & Trim(txtDR_EndDate.Text) & " " & Trim(txtDR_EndTime.Text) & vbNewLine
        End If
    End If
    
    Output_2_List List2, "INI contents finished compiling."
    
    SetupExMergeInI = sBody

End Function

Sub INIParameters()
'   MergeAction
'       0 = specify SourceServerName
'       1 = specify DestServerName
'       2 = specify both
'
'   SourceServername =
'
'   DestServerName =
'
'   DomainControllerForSourceServer =
'
'   SrcServerLDAP-Port = 389
'
'   DomainControllerForDestServer =
'
'   DestServerLDAP-Port =
'
'   LogFileName = file.log
'
'   LoggingLevel
'       0 = None
'       1 = minimum
'       2 = medium
'       3 = maximum
'
'   DataDirectoryName
'       default = c:\EXMERGEDATA
'       where the .pst files are written or exported to.
'
'   FileContainingListOfMailboxes
'       any file, default = mailboxes.txt
'
'   FileContainingListOfDatabases
'       if listed will override FileContainingListOfMailboxes
'
'   SelectMessageStartDate = 1/1/98 00:00:00
'
'   SelectMessageEndDate = 12/31/01 23:59:59
'
'       when using SelectMessageStartDate and SelectMessageEndDate
'       the PR_DELIVERY_TIME MAPI attribute is used.
'
'       to use the modified time (aka PR_MODIFIED_TIME MAPI attribute)
'       use the DateAttribute setting in the ini file to 1
'
'   DataImportMethod = 3
'       0 = copy all messages from the source store to the target store
'       1 = merge messages into the target store.  copy only those messages
'           that do not exist in the target store.
'       2 = replace existing message in the target store. (if a message
'           in the source store exists in the target store, delete that
'           message in the target store and then copy the message from
'           the target store.)
'       3 = archive existing messages from the source store into the
'           target store.  if this option is selected, the program copies data from the source store to the target store and then deletes the data from the source store.  this option is valid only if the mergeaction is extract.
'       use this setting to archive data from the source
'       WARNING - this will delete all messages/items that are exported to the destination
'
'   ReplaceDataOnlyIfSourceItemIsMoreRecent
'       0 = replace all in the target store
'       1 = replace anly items in the target store, if the source store has more recent version
'       this is applicable only if DataImportMethod = 3(replace data)
'       This setting uses PR_LAST_MODIFICATION_TIME when checking datetime
'
'   CopyFolderPermissions = 1
'       using this will copy all folder permissions from each mailbox
'
'   CopyAssociatedFolderData = 1
'
'   CopyUserData
'       0 = do not copy user data (msg, folders, cal, contacts)
'       1 = copy all
'
'   CopyDeletedItemsFromDumpster
'       0 = do not copy items from the dumpster
'       1 = copy items from the dumpster
'       default = 0
'       all dumpster data will be retrieved regardless of date selections
'       all dumpster data recovered will put placed in the Deleted Items folder
'
'
'   -NUMTHREADS #
'       max 10 threads
'
'       boxes < 5 = 1 thread
'       boxes < 25 = 2 threads
'       boxes < 50 = 3 threads
'       boxes < 100 = 4 threads
'       boxes >= 100 = 5 threads
'
'       use of threads is cautioned and is urged to use on a dedicated workstation
'       resources can increase exponentially as the thread count increases.
'
'
'

End Sub

Function ProcessBackup() As Boolean
    Dim sPath As String
    ProcessBackup = False
    
    ChDrive "c:"
    ChDir (Trim(txtExMergePath.Text))
    
    sPath = CreateExMergeString
    
    Output_2_List List2, "Attempting to process items with ExMerge, Please wait."
    ProcessBackup = ShellAndWait(sPath, vbNormalFocus)
End Function

Function CreateExMergeString() As String
    Dim sPath As String
    
    'Create InI Full Path
    sPath = Trim(txtINIPath.Text) & "\" & Trim(txtINIFileName.Text)
    
    'Create Exmerge Shell cmd
    sPath = "exmerge -F " & sPath & " -B "
    
    'Show ExMerge interface
    If chkShowInterface.Value = 1 Then
        sPath = sPath & "-D "
    End If
    
    'Add custom number of threads
    If chkDefaultThreads.Value = 0 Then
        sPath = sPath & "-NUMTHREADS " & cboThreadCount.ListIndex + 1
    End If
    
    'Debug.Print "----"
    'Debug.Print sPath
    'MsgBox sPath
    CreateExMergeString = sPath
    
End Function


Private Sub TreeView1_Click()
    ltotalsize = "0.00"
    Me.txtTotalCombinedSize.Text = ltotalsize
    ltotalcount = 0
    Me.txtTotalCombinedCount.Text = ltotalcount
    prcResetCurrentVars TreeView1.SelectedItem.Key
    Show_MailBoxes TreeView1.SelectedItem.Key
    ClearFlaggedAccounts
End Sub

Sub prcResetCurrentVars(sKey As String)
    Dim tmp() As String
    tmp = Split(sKey, "-")
    
    If UBound(tmp) > 0 Then
        Main.Current.Server = tmp(0)
        Main.Current.StorageGroup = tmp(1)
        Main.Current.MailBoxStoreDBs = tmp(2)
    End If
End Sub

Private Sub txtRootPath_Change()
    If chkUseRoot.Value = 1 Then prcUpdateAllForms
End Sub

Sub prcUpdateAllForms()
    With Me
        .txtMailBoxPath.Text = Trim(txtRootPath.Text)
        .txtINIPath.Text = Trim(txtRootPath.Text)
        .txtPSTPath.Text = Trim(txtRootPath.Text) & "\PST"
        .txtLogPath.Text = Trim(txtRootPath.Text) & "\LOG"
        .txtDotZipFilePath.Text = Trim(txtRootPath.Text) & "\PST"
        .txtMoveToFolder.Text = Trim(txtRootPath.Text) & "\TEMP"
    End With
End Sub


