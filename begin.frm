VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "regTweaker®"
   ClientHeight    =   7455
   ClientLeft      =   2340
   ClientTop       =   825
   ClientWidth     =   7215
   Icon            =   "begin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7215
   Begin VB.Timer tmrTaskTimer 
      Left            =   4560
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pbTaskProgress 
      Height          =   235
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Windows 1"
      TabPicture(0)   =   "begin.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command3(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ass"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Windows 2"
      TabPicture(1)   =   "begin.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command3(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(6)=   "Frame6"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Windows 3"
      TabPicture(2)   =   "begin.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command3(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame15"
      Tab(2).Control(3)=   "Frame13"
      Tab(2).Control(4)=   "Frame12"
      Tab(2).Control(5)=   "Frame11"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Internet Explorer"
      TabPicture(3)   =   "begin.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command3(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Dialog"
      Tab(3).Control(3)=   "Frame18"
      Tab(3).Control(4)=   "Frame17"
      Tab(3).Control(5)=   "Frame16"
      Tab(3).Control(6)=   "Frame14"
      Tab(3).ControlCount=   7
      Begin VB.Timer Timer26 
         Left            =   1560
         Top             =   6600
      End
      Begin VB.TextBox ass 
         Height          =   285
         Left            =   4200
         TabIndex        =   92
         Text            =   """%1"""
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Finished"
         Height          =   375
         Index           =   3
         Left            =   -69240
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Saves all changes."
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Restore All"
         Height          =   375
         Index           =   3
         Left            =   -74880
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Restores all changes made by regTweaker."
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Finished"
         Height          =   375
         Index           =   2
         Left            =   -69240
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Saves all changes."
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Restore All"
         Height          =   375
         Index           =   2
         Left            =   -74880
         TabIndex        =   82
         TabStop         =   0   'False
         ToolTipText     =   "Restores all changes made by regTweaker."
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Finished"
         Height          =   375
         Index           =   1
         Left            =   -69240
         TabIndex        =   81
         TabStop         =   0   'False
         ToolTipText     =   "Saves all changes."
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Restore All"
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   80
         TabStop         =   0   'False
         ToolTipText     =   "Restores all changes made by regTweaker."
         Top             =   6600
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   -69120
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame18 
         Caption         =   "Toolbar Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74760
         TabIndex        =   55
         Top             =   4680
         Width           =   6735
         Begin VB.Timer Timer32 
            Left            =   1560
            Top             =   240
         End
         Begin VB.CommandButton Command38 
            Caption         =   "Customi&ze"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command37 
            Caption         =   "&Browse"
            Height          =   375
            Left            =   5160
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1275
            Width           =   1095
         End
         Begin VB.TextBox bitmap 
            Height          =   315
            Left            =   1920
            TabIndex        =   62
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label Label25 
            Caption         =   "hint:  if you leave this box blank, it will set the background to default"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1560
            TabIndex        =   93
            Top             =   960
            Width           =   4935
         End
         Begin VB.Label Label24 
            Caption         =   "This tweak allows you to choose a a bitmap image to display as the background for the Explorer toolbars. "
            Height          =   495
            Left            =   1560
            TabIndex        =   89
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label6 
            Caption         =   "Choose a Bitmap:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Customized Toolbar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   54
         Top             =   3000
         Width           =   6735
         Begin VB.Timer Timer31 
            Left            =   1680
            Top             =   480
         End
         Begin VB.CommandButton Command36 
            Caption         =   "C&ustomize"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   2640
            TabIndex        =   60
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label Label23 
            Caption         =   $"begin.frx":037A
            Height          =   615
            Left            =   1560
            TabIndex        =   88
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Custom Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1200
            TabIndex        =   59
            Top             =   960
            Width           =   1320
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "UnCache URLs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   53
         Top             =   1800
         Width           =   6735
         Begin VB.Timer Timer30 
            Left            =   1680
            Top             =   360
         End
         Begin VB.CommandButton Command35 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   $"begin.frx":0420
            Height          =   615
            Left            =   1560
            TabIndex        =   87
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Remove Branding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   52
         Top             =   600
         Width           =   6735
         Begin VB.Timer Timer29 
            Left            =   1680
            Top             =   360
         End
         Begin VB.CommandButton Command34 
            Caption         =   "ReBrand"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5400
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Feature not yet functional"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command33 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   $"begin.frx":04CF
            Height          =   615
            Left            =   1560
            TabIndex        =   86
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Customize MouseOvers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   36
         Top             =   4200
         Width           =   6735
         Begin VB.CommandButton Command32 
            Caption         =   "save"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1680
            Width           =   495
         End
         Begin VB.CommandButton Command31 
            Caption         =   "save"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton Command30 
            Caption         =   "save"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox MDIT 
            Height          =   285
            Left            =   1920
            TabIndex        =   48
            Text            =   "Stores your documents, graphics, and other files."
            Top             =   1680
            Width           =   4695
         End
         Begin VB.TextBox RBIT 
            Height          =   285
            Left            =   1920
            TabIndex        =   47
            Text            =   "Contains deleted items you can permanently remove or restore."
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox MCIT 
            Height          =   285
            Left            =   1920
            TabIndex        =   46
            Text            =   "Displays the contents of your computer."
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label Label4 
            Caption         =   "My Documents"
            Height          =   255
            Left            =   720
            TabIndex        =   45
            Top             =   1690
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Recycle Bin"
            Height          =   255
            Left            =   720
            TabIndex        =   44
            Top             =   1090
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "My Computer"
            Height          =   255
            Left            =   720
            TabIndex        =   43
            Top             =   490
            Width           =   1335
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Turn Off Autorun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   35
         Top             =   3000
         Width           =   6735
         Begin VB.Timer Timer25 
            Left            =   4920
            Top             =   360
         End
         Begin VB.Timer Timer24 
            Left            =   1320
            Top             =   360
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Turn O&n"
            Height          =   375
            Left            =   5400
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Turn O&ff"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   $"begin.frx":055C
            Height          =   615
            Left            =   1560
            TabIndex        =   79
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Edit with Notepad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   34
         Top             =   1800
         Width           =   6735
         Begin VB.Timer Timer23 
            Left            =   4680
            Top             =   480
         End
         Begin VB.Timer Timer22 
            Left            =   1560
            Top             =   360
         End
         Begin VB.CommandButton Command28 
            Caption         =   "&Disable"
            Height          =   375
            Left            =   5400
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command25 
            Caption         =   "&Enable"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   $"begin.frx":05F1
            Height          =   615
            Left            =   1560
            TabIndex        =   78
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Command Prompt from Anywhere"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   33
         Top             =   600
         Width           =   6735
         Begin VB.Timer Timer21 
            Left            =   4800
            Top             =   120
         End
         Begin VB.Timer Timer20 
            Left            =   1680
            Top             =   480
         End
         Begin VB.CommandButton Command27 
            Caption         =   "&Disable"
            Height          =   375
            Left            =   5400
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command24 
            Caption         =   "&Enable"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "This tweak creates a new right-click option to open a command prompt from the specified directory."
            Height          =   615
            Left            =   1560
            TabIndex        =   91
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Show Windows Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   23
         Top             =   5400
         Width           =   6735
         Begin VB.Timer Timer19 
            Left            =   4680
            Top             =   240
         End
         Begin VB.Timer Timer18 
            Left            =   1560
            Top             =   240
         End
         Begin VB.CommandButton Command23 
            Caption         =   "&Hide"
            Height          =   375
            Left            =   5400
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command18 
            Caption         =   "&Show"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   $"begin.frx":068C
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   77
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Remove Favorites from Start Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   22
         Top             =   4200
         Width           =   6735
         Begin VB.Timer Timer17 
            Left            =   4920
            Top             =   240
         End
         Begin VB.Timer Timer16 
            Left            =   1560
            Top             =   360
         End
         Begin VB.CommandButton Command22 
            Caption         =   "&Re-add"
            Height          =   375
            Left            =   5400
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Re&move"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "This tweak allows you to remove the Favorites folder from the Start Menu. "
            Height          =   615
            Left            =   1560
            TabIndex        =   76
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "High Color Icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   21
         Top             =   3000
         Width           =   6735
         Begin VB.Timer Timer15 
            Left            =   4560
            Top             =   480
         End
         Begin VB.Timer Timer14 
            Left            =   1440
            Top             =   360
         End
         Begin VB.CommandButton Command21 
            Caption         =   "&Default"
            Height          =   375
            Left            =   5400
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command16 
            Caption         =   "&High Color"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   $"begin.frx":071D
            Height          =   615
            Left            =   1560
            TabIndex        =   75
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Start Menu Delay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   20
         Top             =   1800
         Width           =   6735
         Begin VB.Timer Timer13 
            Left            =   1560
            Top             =   240
         End
         Begin VB.TextBox Text6 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   5400
            TabIndex        =   73
            Text            =   "500"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command15 
            Caption         =   "&Custom Speed"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "         Delay            (in milliseconds)"
            Height          =   375
            Left            =   5280
            TabIndex        =   74
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Windows normally delays menus before they are displayed.With this tweak you to change the delay time or remove it altogether."
            Height          =   615
            Left            =   1560
            TabIndex        =   72
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Remove ""Find"" Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   6735
         Begin VB.Timer Timer12 
            Left            =   4800
            Top             =   360
         End
         Begin VB.Timer Timer11 
            Left            =   1320
            Top             =   360
         End
         Begin VB.CommandButton Command19 
            Caption         =   "&Re-add"
            Height          =   375
            Left            =   5400
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Re&move"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "This tweak will allow you to remove the ""Find: On the Internet..."" and/or ""Find: People..."" items from the Start menu."
            Height          =   615
            Left            =   1560
            TabIndex        =   71
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Restore All"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Restores all changes made by regTweaker."
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Finished"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Saves all changes."
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Allow All Caps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   5400
         Width           =   6735
         Begin VB.Timer Timer10 
            Left            =   4920
            Top             =   240
         End
         Begin VB.Timer Timer9 
            Left            =   1320
            Top             =   240
         End
         Begin VB.CommandButton Command13 
            Caption         =   "&Correct Case"
            Height          =   375
            Left            =   5400
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command12 
            Caption         =   "All&ow"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   $"begin.frx":07BE
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   70
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Remove Network Neighborhood"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   4200
         Width           =   6735
         Begin VB.Timer Timer8 
            Left            =   4800
            Top             =   360
         End
         Begin VB.Timer Timer7 
            Left            =   1560
            Top             =   360
         End
         Begin VB.CommandButton Command11 
            Caption         =   "&Re-add"
            Height          =   375
            Left            =   5400
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Re&move"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   $"begin.frx":0854
            Height          =   615
            Left            =   1560
            TabIndex        =   69
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Eliminate Shortcut Arrows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   6735
         Begin VB.Timer Timer6 
            Left            =   4920
            Top             =   360
         End
         Begin VB.Timer Timer5 
            Left            =   1320
            Top             =   360
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&Put it Back"
            Height          =   375
            Left            =   5400
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command8 
            Caption         =   "&Eliminate"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "This tweak will get rid of the stupid little arrow that is put on the corner of all created shortcuts."
            Height          =   615
            Left            =   1560
            TabIndex        =   68
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Eliminate ""Shortcut to"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   6735
         Begin VB.Timer Timer4 
            Left            =   5040
            Top             =   360
         End
         Begin VB.Timer Timer3 
            Left            =   1440
            Top             =   360
         End
         Begin VB.CommandButton Command7 
            Caption         =   "&Put it Back"
            Height          =   375
            Left            =   5400
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Eliminate"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "This tweak will get rid of the annoying 'Shortcut to' placed in front of every shortcut."
            Height          =   615
            Left            =   1560
            TabIndex        =   67
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rename the Recycle Bin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   6735
         Begin VB.Timer Timer2 
            Left            =   4800
            Top             =   240
         End
         Begin VB.Timer Timer1 
            Left            =   1440
            Top             =   360
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Hide Rename"
            Height          =   375
            Left            =   5400
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Allow Rename"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "This tweak will allow the function to rename the recycle bin.  After applied, just right click and rename."
            Height          =   615
            Left            =   1560
            TabIndex        =   66
            Top             =   240
            Width           =   3615
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "&About regTweaker"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If Win32 Then
    Private Declare Function ShutdownWindows Lib "user32" Alias _
      "ExitWindowsEx" (ByVal uFlags As Long, _
      ByVal dwReserved As Long) As Long
#Else
    Private Declare Function ShutdownWindows Lib "user" Alias _
      "ExitWindows" (ByVal wReturnCode As Integer, _
      ByVal dwReserved As Integer) As Integer
#End If
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4





Private Sub TabStrip1_Click()

End Sub

Private Sub TabStrip2_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub SysInfo1_ConfigChangeCancelled()

End Sub



Private Sub Command1_Click(Index As Integer)
Unload Me

End Sub



Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1(0).BackColor = &HFF00&
Command1(1).BackColor = &HFF00&
Command1(2).BackColor = &HFF00&
Command1(3).BackColor = &HFF00&
End Sub

Private Sub Command10_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood", "1"
    
    Command10.Enabled = False
    MousePointer = vbHourglass
    Command11.Enabled = True
    Command11.Caption = "&Re-add"
    Command10.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer7.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command11_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood", "0"
    
    Command11.Enabled = False
    MousePointer = vbHourglass
    Command10.Enabled = True
    Command10.Caption = "Re&move"
    Command11.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer8.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command12_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "DontPrettyPath", "1"
    
    Command12.Enabled = False
    MousePointer = vbHourglass
    Command13.Enabled = True
    Command13.Caption = "&Correct Case"
    Command12.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer9.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command13_Click()
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "DontPrettyPath", "0"
    
    Command13.Enabled = False
    MousePointer = vbHourglass
    Command12.Enabled = True
    Command12.Caption = "All&ow"
    Command13.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer10.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command14_Click()
    
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\InetFind"
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind"
DeleteKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch"

 
    Command14.Enabled = False
    MousePointer = vbHourglass
    Command19.Enabled = False
    Command19.Caption = "&Readd"
    Command14.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer11.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command15_Click()
SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "MenuShowDelay", " " + Text6.Text + " "
    
    Command15.Enabled = False
    MousePointer = vbHourglass
    Command15.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer13.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command16_Click()
    
SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", "16"
     
    Command16.Enabled = False
    MousePointer = vbHourglass
    Command21.Enabled = True
    Command21.Caption = "&Defualt"
    Command16.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer14.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command17_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", "1"
    
    Command17.Enabled = False
    MousePointer = vbHourglass
    Command22.Enabled = True
    Command22.Caption = "&Re-add"
    Command17.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer17.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command18_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Control Panel\Desktop", "PaintDesktopVersion", "1"

    
    Command18.Enabled = False
    MousePointer = vbHourglass
    Command23.Enabled = True
    Command23.Caption = "&Hide"
    Command18.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer18.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command19_Click()
    
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\InetFind"
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\HelpText"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\DefaultIcon"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0\DefaultIcon"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch", "", "{07798131-AF23-11d1-9111-00A0C98BA67D}"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0", "", "On the &Internet..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\DefaultIcon", "", "C:\WINDOWS\SYSTEM\shdocvw.dll,-111"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\HelpText", "", "Search the web"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind", "", "{32714800-2E5F-11d0-8B85-00AA0044F941}"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0", "", "&People..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0\DefaultIcon", "", "C:\PROGRAM FILES\OUTLOOK EXPRESS\WABFIND.DLL, 0"

    
    Command19.Enabled = False
    MousePointer = vbHourglass
    Command14.Enabled = True
    Command14.Caption = "Re&move"
    Command19.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer12.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub


Private Sub Command21_Click()


SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", "8"
    
    Command21.Enabled = False
    MousePointer = vbHourglass
    Command16.Enabled = True
    Command16.Caption = "&High Color"
    Command21.Caption = "Done"
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer15.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command22_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", "0"
    
    Command22.Enabled = False
    MousePointer = vbHourglass
    Command17.Enabled = True
    Command17.Caption = "Re&move"
    Command22.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer16.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command23_Click()
    
SetDWORDValue "HKEY_CURRENT_USER\Control Panel\Desktop", "PaintDesktopVersion", "0"

    
    Command23.Enabled = False
    MousePointer = vbHourglass
    Command18.Enabled = True
    Command18.Caption = "&Show"
    Command23.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer19.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command24_Click()
    
CreateKey "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt"

CreateKey "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt\command"

SetStringValue "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt", "", "Command Prompt Here"

SetStringValue "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt\command", "", "command.com /k cd " + ass.Text + " "
    
    Command24.Enabled = False
    MousePointer = vbHourglass
    Command27.Enabled = True
    Command27.Caption = "&Disable"
    Command24.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer20.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command25_Click()
    
CreateKey "HKEY_CLASSES_ROOT\*\shell\EditText"

CreateKey "HKEY_CLASSES_ROOT\*\shell\EditText\command"


SetStringValue "HKEY_CLASSES_ROOT\*\shell\EditText", "", "Edit as Text"

SetStringValue "HKEY_CLASSES_ROOT\*\shell\EditText\command", "", "notepad.exe %1"
    
    
    Command25.Enabled = False
    MousePointer = vbHourglass
    Command28.Enabled = True
    Command28.Caption = "&Disable"
    Command25.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer22.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command26_Click()
    
CreateKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom"
SetDWORDValue "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom", "Autorun", "0"
    
    
    Command26.Enabled = False
    MousePointer = vbHourglass
    Command29.Enabled = True
    Command29.Caption = "Turn O&n"
    Command26.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer24.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command27_Click()
    
DeleteKey "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt"


    
    Command27.Enabled = False
    MousePointer = vbHourglass
    Command24.Enabled = True
    Command24.Caption = "&Enable"
    Command27.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer21.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command28_Click()
    
DeleteKey "HKEY_CLASSES_ROOT\*\shell\EditText"

DeleteKey "HKEY_CLASSES_ROOT\*\shell\EditText\command"
    
    Command28.Enabled = False
    MousePointer = vbHourglass
    Command25.Enabled = True
    Command25.Caption = "&Enable"
    Command28.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer23.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command29_Click()
    
CreateKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom"
SetDWORDValue "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom", "Autorun", "1"
    
    
    Command29.Enabled = False
    MousePointer = vbHourglass
    Command26.Enabled = True
    Command26.Caption = "Turn O&ff"
    Command29.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer25.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub



Private Sub Command3_Click(Index As Integer)

      Dim X As Integer
    Dim Response As Variant
    
    Response = MsgBox("This option will restore all registy settings that regTweaker has made." & _
      vbCrLf & "Are you sure you would like to do this?", vbYesNo + vbQuestion, _
      "Restore Registry?")
    If Response = vbYes Then
    
        X = Command3(0).Enabled = False
    Command3(2).Enabled = False
    Command3(3).Enabled = False
    Command3(1).Enabled = False
    Command3(0).Enabled = False
    MousePointer = vbHourglass
    Command3(1).Caption = "Restoring.."
      Command3(2).Caption = "Restoring.."
        Command3(3).Caption = "Restoring.."
          Command3(0).Caption = "Restoring.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer26.Enabled = True
    

'undo rename
SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", Chr$(&H40) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20)

'put back shortcut to prefix
SetBinaryValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer", "link", Chr$(&H15) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)

'name Recycle Bin
SetStringValue "HKEY_CURRENT_USER\Software\Classes\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", "Recycle Bin"

'put back shortcut arrows

SetStringValue "HKEY_CLASSES_ROOT\piffile", "IsShortCut", ""
SetStringValue "HKEY_CLASSES_ROOT\lnkfile", "IsShortCut", ""

'show network neighborhood icon
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetHood", "0"

'correct CaSe
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "DontPrettyPath", "0"

'put original find items back
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\InetFind"
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\HelpText"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\DefaultIcon"
CreateKey "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0\DefaultIcon"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch", "", "{07798131-AF23-11d1-9111-00A0C98BA67D}"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0", "", "On the &Internet..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\DefaultIcon", "", "C:\WINDOWS\SYSTEM\shdocvw.dll,-111"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WebSearch\0\HelpText", "", "Search the web"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind", "", "{32714800-2E5F-11d0-8B85-00AA0044F941}"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0", "", "&People..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions\Static\WabFind\0\DefaultIcon", "", "C:\PROGRAM FILES\OUTLOOK EXPRESS\WABFIND.DLL, 0"

'readd favorites
SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", "0"

'hide windows version
SetDWORDValue "HKEY_CURRENT_USER\Control Panel\Desktop", "PaintDesktopVersion", "0"

'hide COMMAND PROMPT FROM HERE
DeleteKey "HKEY_CLASSES_ROOT\Directory\shell\Command Prompt"

'hide EDIT as TEXT
DeleteKey "HKEY_CLASSES_ROOT\*\shell\EditText"

DeleteKey "HKEY_CLASSES_ROOT\*\shell\EditText\command"

'turn on AutoRun
CreateKey "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom"
SetDWORDValue "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom", "Autorun", "1"

'default IE titlebar
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Window Title", "Microsoft Internet Explorer"

'clear toolbar background
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", ""
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmapIE5", ""

'default mouseovers
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", "Displays the contents of your computer."
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", "Contains deleted items you can permanently remove or restore."
SetStringValue "HKEY_LOCAL_MACHINE\Software\CLASSES\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}", "InfoTip", "Stores your documents, graphics, and other files."

 If Not X Then
            Response = MsgBox("regTweaker is ready to restore your registry, click OK to continue", _
          vbExclamation, "About to restore....")
        End If
    End If


End Sub

Private Sub Command3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Command3(0).BackColor = &HC0&
Command3(1).BackColor = &HC0&
Command3(2).BackColor = &HC0&
Command3(3).BackColor = &HC0&
End Sub

Private Sub Command30_Click()
    Command30.Enabled = False
    MousePointer = vbHourglass
   
    Command30.Caption = "Done"
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    tmrTaskTimer.Enabled = True
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", " " + MCIT.Text + " "
End Sub

Private Sub Command31_Click()
    Command31.Enabled = False
    MousePointer = vbHourglass

    Command31.Caption = "Done"
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    tmrTaskTimer.Enabled = True
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", " " + RBIT.Text + " "
End Sub

Private Sub Command32_Click()
    Command32.Enabled = False
    MousePointer = vbHourglass
    Command32.Caption = "Done"
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    tmrTaskTimer.Enabled = True
 SetStringValue "HKEY_LOCAL_MACHINE\Software\CLASSES\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}", "InfoTip", " " + MDIT.Text + " "
End Sub


Private Sub Command33_Click()
    
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "SmBrandBitmap", ""
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BrandBitmap", ""
    
    Command33.Enabled = False
    MousePointer = vbHourglass

    Command33.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer29.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command35_Click()
    
DeleteKey "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs"

CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs"
    
    Command35.Enabled = False
    MousePointer = vbHourglass

    Command35.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer30.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command36_Click()
    
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "Window Title", " " + Text4.Text + " "
    
    Command36.Enabled = False
    MousePointer = vbHourglass

    Command36.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer31.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command37_Click()
Dialog.DialogTitle = "Select the Bitmap you would like to use..." 'set the dialog title
Dialog.Filter = "BITMAP Files (*.BMP)|*.BMP"
Dialog.ShowOpen
bitmap.Text = Dialog.FileName

End Sub

Private Sub Command38_Click()
    
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", "" + bitmap.Text + ""
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmapIE5", "" + bitmap.Text + ""
    Command38.Enabled = False
    MousePointer = vbHourglass

    Command38.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer32.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command4_Click()
    
SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", Chr$(&H50) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20)
    
    Command4.Enabled = False
    MousePointer = vbHourglass
    Command5.Enabled = True
    Command5.Caption = "&Hide Rename"
    Command4.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer1.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command5_Click()
    
SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", Chr$(&H40) + Chr$(&H1) + Chr$(&H0) + Chr$(&H20)
    
    Command5.Enabled = False
    MousePointer = vbHourglass
    Command4.Enabled = True
    Command4.Caption = "&Allow Rename"
    Command5.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer2.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command6_Click()
    
SetBinaryValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer", "link", Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    
    Command6.Enabled = False
    MousePointer = vbHourglass
    Command7.Enabled = True
    Command7.Caption = "&Put it Back"
    Command6.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer3.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command7_Click()
SetBinaryValue "HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer", "link", Chr$(&H15) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
    


    Command7.Enabled = False
    MousePointer = vbHourglass
    Command6.Enabled = True
    Command6.Caption = "&Eliminate"
    Command7.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer4.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command8_Click()

On Error Resume Next
  modRegistry.DeleteRegKeyValue HKEY_CLASSES_ROOT, "lnkfile", "IsShortCut"
  
  
  modRegistry.DeleteRegKeyValue HKEY_CLASSES_ROOT, "piffile", "IsShortCut"

    
    Command8.Enabled = False
    MousePointer = vbHourglass
    Command9.Enabled = True
    Command9.Caption = "&Put it Back"
    Command8.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer5.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Command9_Click()

 On Error Resume Next
   SetStringValue "HKEY_CLASSES_ROOT\piffile", "IsShortCut", ""
   SetStringValue "HKEY_CLASSES_ROOT\lnkfile", "IsShortCut", ""
    
    Command9.Enabled = False
    MousePointer = vbHourglass
    Command8.Enabled = True
    Command8.Caption = "&Eliminate"
    Command9.Caption = "Doing it.."
    pbTaskProgress.Value = pbTaskProgress.Min
    pbTaskProgress.Visible = True
    Timer6.Enabled = True
    tmrTaskTimer.Enabled = True
End Sub

Private Sub Form_Load()
CreateKey "HKEY_LOCAL_MACHINE\Software\Johneboy Inc."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Johneboy Inc.", "First Run", "no"


    pbTaskProgress.Visible = True
    pbTaskProgress.Align = vbAlignNone
    pbTaskProgress.Min = 0
    pbTaskProgress.Max = 100

    tmrTaskTimer.Interval = 100
    tmrTaskTimer.Enabled = False
    Timer1.Interval = 1000
    Timer1.Enabled = False
        Timer2.Interval = 1000
    Timer2.Enabled = False
        Timer3.Interval = 1000
    Timer3.Enabled = False
        Timer4.Interval = 1000
    Timer4.Enabled = False
        Timer5.Interval = 1000
    Timer5.Enabled = False
        Timer6.Interval = 1000
    Timer6.Enabled = False
        Timer7.Interval = 1000
    Timer7.Enabled = False
        Timer8.Interval = 1000
    Timer8.Enabled = False
        Timer9.Interval = 1000
    Timer9.Enabled = False
        Timer10.Interval = 1000
    Timer10.Enabled = False
        Timer11.Interval = 1000
    Timer11.Enabled = False
        Timer12.Interval = 1000
    Timer12.Enabled = False
        Timer13.Interval = 1000
    Timer13.Enabled = False
        Timer14.Interval = 1000
    Timer14.Enabled = False
        Timer15.Interval = 1000
    Timer15.Enabled = False
        Timer16.Interval = 1000
    Timer16.Enabled = False
        Timer17.Interval = 1000
    Timer17.Enabled = False
        Timer18.Interval = 1000
    Timer18.Enabled = False
        Timer19.Interval = 1000
    Timer19.Enabled = False
        Timer20.Interval = 1000
    Timer20.Enabled = False
        Timer21.Interval = 1000
    Timer21.Enabled = False
        Timer22.Interval = 1000
    Timer22.Enabled = False
        Timer23.Interval = 1000
    Timer23.Enabled = False
        Timer24.Interval = 1000
    Timer24.Enabled = False
        Timer25.Interval = 1000
    Timer25.Enabled = False
        Timer29.Interval = 1000
    Timer29.Enabled = False
        Timer30.Interval = 1000
    Timer30.Enabled = False
        Timer31.Interval = 1000
    Timer31.Enabled = False
        Timer32.Interval = 1000
    Timer32.Enabled = False
    
    Timer26.Interval = 1000
        Timer26.Enabled = False
  

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80000007
End Sub

Private Sub Form_Unload(Cancel As Integer)

If Command14.Caption = "Done" Or Command17.Caption = "Done" Or Command13.Caption = "Done" Or Command12.Caption = "Done" Or Command11.Caption = "Done" Or Command10.Caption = "Done" Or Command16.Caption = "Done" Or Command18.Caption = "Done" Or Command19.Caption = "Done" Or Command21.Caption = "Done" Or Command22.Caption = "Done" Or Command23.Caption = "Done" Or Command6.Caption = "Done" Or Command7.Caption = "Done" Or Command8.Caption = "Done" Or Command9.Caption = "Done" Or Command17.Caption = "Done" Or Command17.Caption = "Done" Or Command3(0).Caption = "Restored" Then


    Dim X As Integer
    Dim Response As Variant
    
    Response = MsgBox("In order for changes to take effect, regTweaker must restart Windows." & _
      vbCrLf & "Would you like to restart now?", vbYesNo + vbQuestion, _
      "Restart Windows?")
    If Response = vbYes Then
       X = ShutdownWindows(EWX_REBOOT, 0)
        If Not X Then
            Response = MsgBox("Some program(s) refused to terminate", _
          vbExclamation, "Cannot Restart Windows")
        End If
    End If
   If Not Command14.Caption = "Done" Then
   
    
    Unload Me
End If
    End If
End Sub

Private Sub Label1_Click()
About.Show 1
End Sub

Private Sub restore1_Click(Index As Integer)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.ForeColor = &H80000007
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Command4.Caption = "Done"


End Sub

Private Sub Timer10_Timer()
Timer10.Enabled = False
Command13.Caption = "Done"
End Sub

Private Sub Timer11_Timer()
Timer11.Enabled = False
Command14.Caption = "Done"
End Sub

Private Sub Timer12_Timer()
Timer12.Enabled = False
Command19.Caption = "Done"
End Sub

Private Sub Timer13_Timer()
Timer13.Enabled = False
Command15.Caption = "&Customize"
Command15.Enabled = True
End Sub

Private Sub Timer14_Timer()
Timer14.Enabled = False
Command16.Caption = "Done"
End Sub

Private Sub Timer15_Timer()
Timer15.Enabled = False
Command21.Caption = "Done"
End Sub

Private Sub Timer16_Timer()
Timer16.Enabled = False
Command22.Caption = "Done"
End Sub

Private Sub Timer17_Timer()
Timer17.Enabled = False
Command17.Caption = "Done"
End Sub

Private Sub Timer18_Timer()
Timer18.Enabled = False
Command18.Caption = "Done"
End Sub

Private Sub Timer19_Timer()
Timer19.Enabled = False
Command23.Caption = "Done"
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Command5.Caption = "Done"
End Sub

Private Sub Timer20_Timer()
Timer20.Enabled = False
Command24.Caption = "Done"
End Sub

Private Sub Timer21_Timer()
Timer21.Enabled = False
Command27.Caption = "Done"
End Sub

Private Sub Timer22_Timer()
Timer22.Enabled = False
Command25.Caption = "Done"
End Sub

Private Sub Timer23_Timer()
Timer23.Enabled = False
Command28.Caption = "Done"
End Sub

Private Sub Timer24_Timer()
Timer24.Enabled = False
Command26.Caption = "Done"
End Sub

Private Sub Timer25_Timer()
Timer25.Enabled = False
Command29.Caption = "Done"
End Sub

Private Sub Timer26_Timer()
    pbTaskProgress.Value = _
        pbTaskProgress.Value + 10

    If pbTaskProgress.Value >= pbTaskProgress.Max _
    Then
        pbTaskProgress.Visible = False
Timer26.Enabled = False
   Command3(1).Caption = "Restored"
   Command3(2).Caption = "Restored"
   Command3(3).Caption = "Restored"
   Command3(0).Caption = "Restored"
    MousePointer = vbDefault
   End If
End Sub

Private Sub Timer27_Timer()
Timer1.Enabled = False
Command4.Caption = "Done"
End Sub

Private Sub Timer28_Timer()
Timer1.Enabled = False
Command4.Caption = "Done"
End Sub

Private Sub Timer29_Timer()
Timer29.Enabled = False
Command33.Caption = "Done"
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Command6.Caption = "Done"
End Sub

Private Sub Timer30_Timer()
Timer30.Enabled = False
Command35.Caption = "Done"
End Sub

Private Sub Timer31_Timer()
Timer31.Enabled = False
Command36.Caption = "Done"
End Sub

Private Sub Timer32_Timer()
Timer32.Enabled = False
Command38.Caption = "Done"
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Command7.Caption = "Done"
End Sub

Private Sub Timer5_Timer()
Timer5.Enabled = False
Command8.Caption = "Done"
End Sub

Private Sub Timer6_Timer()
Timer6.Enabled = False
Command9.Caption = "Done"
End Sub

Private Sub Timer7_Timer()
Timer7.Enabled = False
Command10.Caption = "Done"
End Sub

Private Sub Timer8_Timer()
Timer8.Enabled = False
Command11.Caption = "Done"
End Sub

Private Sub Timer9_Timer()
Timer9.Enabled = False
Command12.Caption = "Done"
End Sub

Private Sub tmrTaskTimer_Timer()
    pbTaskProgress.Value = _
        pbTaskProgress.Value + 10

    If pbTaskProgress.Value >= pbTaskProgress.Max _
    Then
        pbTaskProgress.Visible = False
       
       
        
        tmrTaskTimer.Enabled = False
        MousePointer = vbDefault
       
       
    End If
End Sub
