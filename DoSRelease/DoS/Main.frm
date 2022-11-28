VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DIA Batch Release"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8040
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   312
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   6360
   End
   Begin MSComDlg.CommonDialog dlgHelp 
      Left            =   2520
      Top             =   6930
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      HelpFile        =   "TextRel.hlp"
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "#Help"
      Height          =   312
      Left            =   7020
      TabIndex        =   67
      Top             =   7005
      Width           =   900
   End
   Begin TabDlg.SSTab tabText 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   600
      TabCaption(0)   =   "#Index Storage"
      TabPicture(0)   =   "Main.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "#Document Storage"
      TabPicture(1)   =   "Main.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "#Image Format"
      TabPicture(2)   =   "Main.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ODBC Connection"
      TabPicture(3)   =   "Main.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraTabODBC"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraTabODBC 
         Height          =   5175
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtPWD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   3400
            PasswordChar    =   "*"
            TabIndex        =   85
            Top             =   4200
            Width           =   2000
         End
         Begin VB.TextBox txtUID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   3400
            PasswordChar    =   "*"
            TabIndex        =   84
            Top             =   3200
            Width           =   2000
         End
         Begin VB.CheckBox chkWinAuth 
            Caption         =   "Use Windows Authentication"
            Height          =   495
            Left            =   2000
            TabIndex        =   83
            Top             =   2400
            Width           =   3000
         End
         Begin VB.TextBox txtDB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   3400
            PasswordChar    =   "*"
            TabIndex        =   82
            Top             =   1600
            Width           =   2000
         End
         Begin VB.TextBox txtDSN 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   3400
            PasswordChar    =   "*"
            TabIndex        =   81
            Top             =   675
            Width           =   2000
         End
         Begin VB.Label Label6 
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2000
            TabIndex        =   80
            Top             =   1700
            Width           =   1300
         End
         Begin VB.Label Label5 
            Caption         =   "User ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2000
            TabIndex        =   79
            Top             =   3300
            Width           =   1300
         End
         Begin VB.Label Label4 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2000
            TabIndex        =   78
            Top             =   4300
            Width           =   1300
         End
         Begin VB.Label Label3 
            Caption         =   "DSN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2000
            TabIndex        =   77
            Top             =   800
            Width           =   1300
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Enter ODBC Connection Values"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1200
            TabIndex        =   75
            Top             =   200
            Width           =   5055
         End
      End
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5340
         Index           =   2
         Left            =   -74950
         TabIndex        =   73
         Top             =   360
         Width           =   7695
         Begin VB.Frame fraReleaseImagesAs 
            Caption         =   "#Release Images As"
            Height          =   612
            Left            =   165
            TabIndex        =   60
            Top             =   120
            Width           =   7332
            Begin VB.CommandButton cmdSettings 
               Caption         =   "#Settings..."
               Height          =   312
               Left            =   6200
               TabIndex        =   63
               Top             =   200
               Width           =   940
            End
            Begin VB.ComboBox cboImageType 
               Height          =   315
               ItemData        =   "Main.frx":007C
               Left            =   1800
               List            =   "Main.frx":007E
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   62
               Top             =   200
               Width           =   4200
            End
            Begin VB.Label lblImageType 
               Caption         =   "#Image file type:"
               Height          =   192
               Left            =   240
               TabIndex        =   61
               Top             =   261
               Width           =   1332
            End
         End
      End
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5340
         Index           =   1
         Left            =   -74950
         TabIndex        =   68
         Top             =   360
         Width           =   7695
         Begin VB.Frame fraImageFiles 
            Caption         =   " #Image Files "
            Height          =   1575
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   7440
            Begin VB.CheckBox chkReleaseImageFiles 
               Caption         =   "#Release image files"
               Height          =   210
               Left            =   180
               TabIndex        =   45
               Top             =   360
               Width           =   6000
            End
            Begin VB.TextBox txtImageDir 
               Height          =   312
               Left            =   1680
               TabIndex        =   47
               Top             =   720
               Width           =   4560
            End
            Begin VB.CommandButton cmdImageBrowse 
               Caption         =   "#Browse..."
               Height          =   312
               Left            =   6360
               TabIndex        =   48
               Top             =   720
               Width           =   900
            End
            Begin VB.CheckBox chkSkipFirstPage 
               Caption         =   "#Skip first page of each document"
               Height          =   210
               Left            =   180
               TabIndex        =   49
               Top             =   1170
               Width           =   6000
            End
            Begin VB.Label lblImageDir 
               Caption         =   "#Release directory:"
               Height          =   195
               Left            =   180
               TabIndex        =   46
               Top             =   779
               Width           =   1455
            End
         End
         Begin VB.Frame fraOCRFiles 
            Caption         =   " #OCR Full Text File "
            Height          =   1185
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   7440
            Begin VB.CheckBox chkReleaseOCRFullText 
               Caption         =   "#Release OCR full text"
               Height          =   210
               Left            =   180
               TabIndex        =   51
               Top             =   360
               Width           =   6000
            End
            Begin VB.TextBox txtOCRDir 
               Height          =   312
               Left            =   1680
               TabIndex        =   53
               Top             =   690
               Width           =   4560
            End
            Begin VB.CommandButton cmdOCRBrowse 
               Caption         =   "#Browse..."
               Height          =   312
               Left            =   6360
               TabIndex        =   54
               Top             =   690
               Width           =   900
            End
            Begin VB.Label lblOCRDir 
               Caption         =   "#Release directory:"
               Height          =   255
               Left            =   180
               TabIndex        =   52
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame fraPDFFiles 
            Caption         =   "#Kofax PDF Files"
            Height          =   1185
            Left            =   120
            TabIndex        =   55
            Top             =   3120
            Width           =   7440
            Begin VB.CommandButton cmdPDFBrowse 
               Caption         =   "#Browse..."
               Height          =   312
               Left            =   6360
               TabIndex        =   59
               Top             =   690
               Width           =   900
            End
            Begin VB.TextBox txtKofaxPDFDir 
               Height          =   312
               Left            =   1680
               TabIndex        =   58
               Top             =   690
               Width           =   4560
            End
            Begin VB.CheckBox chkReleaseKofaxPDF 
               Caption         =   "#Release Kofax PDF files"
               Height          =   210
               Left            =   180
               TabIndex        =   56
               Top             =   330
               Width           =   6000
            End
            Begin VB.Label lblPDFDir 
               Caption         =   "#Release directory:"
               Height          =   255
               Left            =   180
               TabIndex        =   57
               Top             =   720
               Width           =   1575
            End
         End
      End
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'None
         Height          =   5340
         Index           =   0
         Left            =   -74950
         TabIndex        =   71
         Top             =   360
         Width           =   7695
         Begin VB.Frame fraIndexVals 
            Caption         =   "#Index Values "
            Height          =   3540
            Left            =   120
            TabIndex        =   8
            Top             =   885
            Width           =   7440
            Begin VB.CommandButton cmdDeleteAllIndex 
               Caption         =   "#Delete All"
               Height          =   312
               Left            =   6360
               TabIndex        =   41
               Top             =   1155
               Width           =   900
            End
            Begin VB.CommandButton cmdAddIndex 
               Caption         =   "#Add"
               Height          =   312
               Left            =   6360
               TabIndex        =   39
               Top             =   315
               Width           =   900
            End
            Begin VB.CommandButton cmdDeleteIndex 
               Caption         =   "#Delete"
               Enabled         =   0   'False
               Height          =   312
               Left            =   6360
               TabIndex        =   40
               Top             =   735
               Width           =   900
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   4
               Left            =   5598
               Picture         =   "Main.frx":0080
               Style           =   1  'Graphical
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1875
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   3
               Left            =   5598
               Picture         =   "Main.frx":038A
               Style           =   1  'Graphical
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1575
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   2
               Left            =   5598
               Picture         =   "Main.frx":0694
               Style           =   1  'Graphical
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1275
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   1
               Left            =   5598
               Picture         =   "Main.frx":099E
               Style           =   1  'Graphical
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   975
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   0
               Left            =   5598
               Picture         =   "Main.frx":0CA8
               Style           =   1  'Graphical
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   675
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   5
               Left            =   5598
               Picture         =   "Main.frx":0FB2
               Style           =   1  'Graphical
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   2175
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   6
               Left            =   5598
               Picture         =   "Main.frx":12BC
               Style           =   1  'Graphical
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   2475
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   7
               Left            =   5598
               Picture         =   "Main.frx":15C6
               Style           =   1  'Graphical
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   2775
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.CommandButton cmdMenu 
               Height          =   288
               Index           =   8
               Left            =   5598
               Picture         =   "Main.frx":18D0
               Style           =   1  'Graphical
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   3075
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   8
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   35
               TabStop         =   0   'False
               Text            =   "9"
               Top             =   3075
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   7
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   32
               TabStop         =   0   'False
               Text            =   "8"
               Top             =   2775
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   6
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "7"
               Top             =   2475
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   5
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   26
               TabStop         =   0   'False
               Text            =   "6"
               Top             =   2175
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   4
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "5"
               Top             =   1875
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.VScrollBar vsbIndex 
               Height          =   2688
               LargeChange     =   8
               Left            =   5940
               Max             =   1
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   675
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   3
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   20
               TabStop         =   0   'False
               Text            =   "4"
               Top             =   1575
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   2
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "3"
               Top             =   1275
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   1
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "2"
               Top             =   975
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.TextBox txtSequence 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000012&
               Height          =   288
               Index           =   0
               Left            =   180
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   11
               TabStop         =   0   'False
               Text            =   "1"
               Top             =   675
               Visible         =   0   'False
               Width           =   1620
            End
            Begin ComCtl2.UpDown updnIndex 
               Height          =   600
               Left            =   6360
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1575
               Width           =   195
               _ExtentX        =   423
               _ExtentY        =   1058
               _Version        =   327681
               Alignment       =   0
               OrigLeft        =   6330
               OrigTop         =   1575
               OrigRight       =   6525
               OrigBottom      =   2175
               Max             =   0
               Min             =   8
               Enabled         =   0   'False
            End
            Begin VB.PictureBox Picture1 
               Height          =   168
               Left            =   6696
               ScaleHeight     =   105
               ScaleWidth      =   120
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   360
               Width           =   180
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   8
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   3075
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   7
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   2775
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   6
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   2475
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   5
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   2175
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   0
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   12
               Top             =   684
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   1
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   975
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   2
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   1275
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   3
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1575
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.TextBox txtIndexData 
               Height          =   288
               Index           =   4
               Left            =   1860
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   1875
               Visible         =   0   'False
               Width           =   3990
            End
            Begin VB.Label lblMove 
               Caption         =   "#Move"
               Enabled         =   0   'False
               Height          =   240
               Left            =   6660
               TabIndex        =   43
               Top             =   1740
               Width           =   600
            End
            Begin VB.Label lblIndexLabel 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "#Index Value"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   1
               Left            =   1860
               TabIndex        =   10
               Top             =   315
               Width           =   3990
            End
            Begin VB.Label lblIndexLabel 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "#Sequence"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   0
               Left            =   180
               TabIndex        =   9
               Top             =   315
               Width           =   1620
            End
         End
         Begin VB.CommandButton cmdFileBrowse 
            Caption         =   "#Browse..."
            Height          =   312
            Left            =   6480
            TabIndex        =   7
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox txtFileName 
            Height          =   312
            Left            =   1560
            TabIndex        =   6
            Top             =   360
            Width           =   4740
         End
         Begin VB.Label lblFileName 
            Caption         =   "#File name:"
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   405
            Width           =   1245
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "#Apply"
      Enabled         =   0   'False
      Height          =   312
      Left            =   6000
      TabIndex        =   66
      Top             =   7005
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "#Cancel"
      Height          =   312
      Left            =   4980
      TabIndex        =   65
      Top             =   7005
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "#OK"
      Default         =   -1  'True
      Height          =   312
      Left            =   3960
      TabIndex        =   64
      Top             =   7005
      Width           =   900
   End
   Begin MSComDlg.CommonDialog dlgDialogs 
      Left            =   3240
      Top             =   6930
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "ODBC DSN"
      Height          =   495
      Left            =   1500
      TabIndex        =   76
      Top             =   2300
      Width           =   2000
   End
   Begin VB.Label lblDocClass 
      Caption         =   "#Document Class:"
      Height          =   200
      Left            =   120
      TabIndex        =   1
      Top             =   359
      Width           =   1380
   End
   Begin VB.Label lblDocClassName 
      Height          =   200
      Left            =   1560
      TabIndex        =   70
      Top             =   359
      Width           =   6360
   End
   Begin VB.Label lblName 
      Caption         =   "#Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   659
      Width           =   1380
   End
   Begin VB.Label lblBatchClassName 
      Height          =   200
      Left            =   1560
      TabIndex        =   69
      Top             =   60
      Width           =   6360
   End
   Begin VB.Label lblBatchClass 
      Caption         =   "#Batch Class:"
      Height          =   200
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1380
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This string is written to the error log to identify
' in which source module the error occurred
Private Const M_SETUPFORM = "pdfCD Release with variable index names"

Const IN_PROGRESS = "IP"
Const BUTTON_CLICK = "BC"
Const SKIP_EVENT = "SKIP"

Enum ReturnType
    rtFATAL_ERROR = -1
    rtLOADING = 0
    rtOK = 1
    rtDONE = 2
End Enum

Enum ProductCode
    pcAscentCapture = 1
    pcTitan = 2
End Enum

'=======================
' Object Declarations
'=======================
' Remember to release these in the Form_Unload event
Dim SetupData As ReleaseSetupData
Dim oTextFile As New ASCIITextFile
Dim oMenu As New frmMenu

'=======================
' Form Level Variables
'=======================
Dim fCurrStatus As ReturnType
Dim fDirty As Boolean
Dim fTabKeepsFocus As Boolean
Dim fIndexList() As T_Link
Dim fSavedLink As T_Link
Dim fIndexCount As Integer
Dim fSelectedIndex As Integer
Dim fVerified As Boolean

' Internationalized PDF strings for image type combo.
Private fstrPDFImageTypes(4) As String

'=======================
' Global Variables
'=======================
Public gNewIndexType As Integer
Public gNewIndexData As String


'*************************************************
' Dirty [Let Property]
'-------------------------------------------------
' Purpose:  The dirty property will set the
'           current status of the data.  If
'           the data is dirty, the Apply
'           button is enabled.
' Inputs:   NewStatus   Boolean indicating if
'                       data is dirty (TRUE)
'                       or clean (FALSE)
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Public Property Let Dirty(NewStatus As Boolean)
        fDirty = NewStatus
        If (fDirty) Then
            cmdApply.Enabled = True
        Else
            cmdApply.Enabled = False
        End If
End Property

'*************************************************
' Dirty [Get Property]
'-------------------------------------------------
' Purpose:  The dirty property will return
'           the current status of the data.
' Inputs:   None
' Outputs:  TRUE if the data is dirty
'           FALSE if the data is clean
' Returns:  None
' Notes:    None
'*************************************************
Public Property Get Dirty() As Boolean
        Dirty = fDirty
End Property

'*************************************************
' FormStatus [Let Property]
'-------------------------------------------------
' Purpose:  Set the current status of the form.
' Inputs:   NewStatus   ReurnType indicating the
'                       current state of the form
' Outputs:  None
' Returns:  None
' Notes:    ReturnType is an enum defined at
'           the top of this file.  These
'           values are used to determine if
'           the form is being loaded and to
'           define the result returned from
'           the release setup script.
'*************************************************
Public Property Let FormStatus(NewStatus As ReturnType)
        fCurrStatus = NewStatus
End Property

'*************************************************
' FormStatus [Get Property]
'-------------------------------------------------
' Purpose:  Get the current status of the form.
' Inputs:   None
' Outputs:  None
' Returns:  ReturnType - rtLOADING, rtOK
'           rtDONE, or rtFATAL_ERROR
' Notes:    ReturnType is an enum defined at
'           the top of this file.  These
'           values are used to determine if
'           the form is being loaded and to
'           define the result returned from
'           the release setup script.
'*************************************************
Public Property Get FormStatus() As ReturnType
        FormStatus = fCurrStatus
End Property

'*************************************************
' AddString
'-------------------------------------------------
' Purpose:  Formats a specified string to be
'           appended on a new line of another
'           string.  If more than 10 strings
'           have been appended, the function
'           substitutes the phrase "And More"
'           for the specified string.
' Inputs:   nCount  the number of strings that
'                   have been appended.
'           sField  the string to append
' Outputs:  None
' Returns:  The formatted string
' Notes:    This function is used solely by the
'           data verification routines to list
'           the Index Fields and Batch Fields
'           that were not used as document
'           Index Values.
'*************************************************
Function AddString(nCount As Long, sField As String) As String
        nCount = nCount + 1
        If (nCount < 10) Then
            AddString = vbCrLf + vbTab + "- " + sField
        ElseIf (nCount = 10) Then
100         AddString = vbCrLf + vbTab + LoadResString(MSG_ANDMORE)
        Else
            AddString = ""
        End If
End Function

'*************************************************
' BuildLinkingMenu
'-------------------------------------------------
' Purpose:  This routine will build the linking
'           popup menu used by the link box.
'           It will add entries for each of the
'           Index Fields, Batch Fields, and
'           Ascent Capture Values.
' Inputs:   oSetupData  ReleaseSetupData object
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub BuildLinkingMenu(oSetupData As ReleaseSetupData)

        On Error GoTo BLM_LogAndPropError
        
        ' Build the three variable linking menu lists.
        ' This only needs to be done once at the start.
300     Call oMenu.BuildAscentMenu(oSetupData)
310     Call oMenu.BuildBatchMenu(oSetupData)
320     Call oMenu.BuildIndexMenu(oSetupData)

        Exit Sub

BLM_LogAndPropError:

    Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' cboImageType_Click
'-------------------------------------------------
' Purpose:  The user selected a different image
'           format.  Mark the data dirty.  If
'           they chose PDF, enable the controls
'           to define the PDF settings.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cboImageType_Click()
        
    If (cboImageType.ItemData(cboImageType.ListIndex) <> SetupData.ImageType) Then
        Me.Dirty = True
    End If
  
    '*** CAP Tools
    frmAdobeAcrobatSetup.pdfImageFormat3.PDFInputImageType = cboImageType.ItemData(cboImageType.ListIndex)
    
    ' Enable PDF User interface if selection is PDF
    Call EnableAdobeAcrobatSettings
    
End Sub



'*************************************************
' chkReleaseImageFiles_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           release image files of each
'           document.  It enables/disables
'           various controls that are used to
'           to release images.
'           Mark the data dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkReleaseImageFiles_Click()
    lblImageDir.Enabled = CBool(chkReleaseImageFiles.Value)
    txtImageDir.Enabled = CBool(chkReleaseImageFiles.Value)
    cmdImageBrowse.Enabled = CBool(chkReleaseImageFiles.Value)
    chkSkipFirstPage.Enabled = CBool(chkReleaseImageFiles.Value)
    frmAdobeAcrobatSetup.chkWaitForStatus.Enabled = CBool(chkReleaseImageFiles.Value)
    frmAdobeAcrobatSetup.chkDeleteOnHung.Enabled = CBool(frmAdobeAcrobatSetup.chkWaitForStatus.Value) _
        And frmAdobeAcrobatSetup.chkWaitForStatus.Enabled And CBool(chkReleaseImageFiles.Value)
    fraReleaseImagesAs.Enabled = CBool(chkReleaseImageFiles.Value)
    lblImageType.Enabled = CBool(chkReleaseImageFiles.Value)
    cboImageType.Enabled = CBool(chkReleaseImageFiles.Value)
    Me.Dirty = True
End Sub



'*************************************************
' chkReleaseOCRFullText_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           release OCR Full Text of each
'           document.  It enables/disables
'           controls that are used for OCR Full
'           Text setup.  Mark the data dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkReleaseOCRFullText_Click()
    lblOCRDir.Enabled = CBool(chkReleaseOCRFullText.Value)
    txtOCRDir.Enabled = CBool(chkReleaseOCRFullText.Value)
    cmdOCRBrowse.Enabled = CBool(chkReleaseOCRFullText.Value)
    Me.Dirty = True
End Sub

'*************************************************
' chkReleaseKofaxPDF_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           release Kofax PDF files of each
'           document.  It enables/disables
'           various controls that are used to
'           to release images.
'           Mark the data dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkReleaseKofaxPDF_Click()
    lblPDFDir.Enabled = CBool(chkReleaseKofaxPDF.Value)
    txtKofaxPDFDir.Enabled = CBool(chkReleaseKofaxPDF.Value)
    cmdPDFBrowse.Enabled = CBool(chkReleaseKofaxPDF.Value)
    Me.Dirty = True
End Sub

'*************************************************
' chkSkipFirstPage_Click
'-------------------------------------------------
' Purpose:  The user toggled whether or not to
'           release the first page of each
'           document.  Mark the data dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub chkSkipFirstPage_Click()

        Me.Dirty = True

End Sub



Private Sub chkWinAuth_Click()
    If chkWinAuth.Value = vbChecked Then
        txtUID.Enabled = False
        txtPWD.Enabled = False
    Else
        txtUID.Enabled = True
        txtPWD.Enabled = True
    End If
End Sub

'*************************************************
' cmdAddIndex_Click
'-------------------------------------------------
' Purpose:  Add a blank (unlinked) Index Value
'           to the end of the list and place
'           focus on the control.  Mark the
'           data dirty.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdAddIndex_Click()
    
    On Error GoTo AddIndex_Failure
    
        Me.Dirty = True
    
        ' Add a blank Index Value at the end of the list
400     ReDim Preserve fIndexList(fIndexCount)
410     With fIndexList(fIndexCount)
            .Destination = fIndexCount
            .Source = ""
            .SourceType = NO_LINK
        End With
        fIndexCount = fIndexCount + 1
    
        ' Show the new Index Value and give it the focus
        If fIndexCount > LINK_BOX_SIZE Then
420         Call DisplayIndexValues(fIndexCount - LINK_BOX_SIZE)
430         txtIndexData(LINK_BOX_SIZE - 1).SetFocus
        Else
440         Call DisplayIndexValues(0)
450         txtIndexData(fIndexCount - 1).SetFocus
        End If
        
        Exit Sub

AddIndex_Failure:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)
    
End Sub

'*************************************************
' cmdApply_Click
'-------------------------------------------------
' Purpose:  Verify the settings.  If there are
'           no errors, save the changes and
'           allow the user to continue editting.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    If the settings are validated and
'           saved, the data is marked clean.
'*************************************************
Private Sub cmdApply_Click()
        
        On Error GoTo Unexpected_Error
        
460     If (VerifyReleaseSettings()) Then
470         Call SaveReleaseSettings
        End If
        
        Exit Sub

Unexpected_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdCancel_Click
'-------------------------------------------------
' Purpose:  Discard changes and unload the form.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdCancel_Click()

        ' When the user is in the middle of editing a
        ' a Text Constant, the ESC key triggers the
        ' Cancel button's click event.  We simply
        ' discard any changes and lock the text box.
        If ActiveControl.Name = "txtIndexData" Then
            ' Check if the Index Value text box is unlocked
            If ActiveControl.Locked = False Then
                ' Prompt the user whether to lose changes
                If MsgBox(LoadResString(MSG_DISCARDTEXTCONST), _
                          vbOKCancel + vbInformation, _
                          LoadResString(TITLE_TEXTRSETUP)) = vbOK Then
                    ' Lock the Text Constant text box and
                    ' restore the previous Index Value
480                 Call LockTextConstant(ActiveControl.Index, False)
                End If
                Exit Sub
            Else
                ' Set focus to the Cancel button.  The call to
                ' DoEvents is required to process the
                ' txtIndexData_LostFocus event which forces the
                ' selected Index Value to lose focus and clean
                ' up its highlighting and menu button.
                cmdCancel.SetFocus
                DoEvents
            End If
        End If
    
490     Unload Me
        
End Sub

'*************************************************
' cmdDeleteAllIndex_Click
'-------------------------------------------------
' Purpose:  Ask the user if it is OK and then
'           delete all defined Index Values.
'           Mark the data dirty if the Index
'           Values are deleted.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdDeleteAllIndex_Click()
    Dim Result As Integer

        On Error GoTo DeleteAll_Error
        
       ' Verify that the user REALLY
       ' wants to delete all Index Values
        Result = MsgBox(LoadResString(MSG_DELETEALLINDEX), _
                        vbExclamation + vbYesNo, _
                        LoadResString(TITLE_TEXTRSETUP))
                        
        If (Result = vbYes) Then
            ' Go ahead and delete them
500         Call DeleteAllIndex
            Me.Dirty = True
    
            ' Display the empty list
510         Call DisplayIndexValues(0)
        End If

        Exit Sub

DeleteAll_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdDeleteIndex_Click
'-------------------------------------------------
' Purpose:  Ask the user if it is OK and then
'           delete the selected Index Value.
'           Mark the data dirty if the Index
'           Value is deleted.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    VB bahavior causes this function to
'           get called recursively so we set a
'           local InProgress flag to skip it.
'*************************************************
Public Sub cmdDeleteIndex_Click()
    Static InProgress As Boolean
    Dim Result As Integer

        If InProgress Then Exit Sub
        
        InProgress = True
        DoEvents
        
        On Error GoTo DeleteIndex_Error
        
        ' Verify that the user REALLY wants
        ' to delete the selected Index Value
        Result = MsgBox(LoadResString(MSG_DELETEINDEX), _
                        vbExclamation + vbYesNo, _
                        LoadResString(TITLE_TEXTRSETUP))
                        
        If (Result = vbYes) Then
            ' Go ahead and delete it
550         Call DeleteIndex(fSelectedIndex)
            Me.Dirty = True
        
            ' If the last Index Value was deleted then
            ' fSelectedIndex is now invalid and the
            ' previous Index Value is selected
            If fSelectedIndex = fIndexCount Then
                fSelectedIndex = fSelectedIndex - 1
            End If
                
            ' Display the modified list
560         Call DisplayIndexValues(vsbIndex.Value)
        End If
    
        ' Set focus back to the textbox
        If fIndexCount > 0 Then
570         txtIndexData(fSelectedIndex - vsbIndex.Value).SetFocus
        End If
    
        InProgress = False
        Exit Sub

DeleteIndex_Error:
    
        InProgress = False
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdDeleteIndex_GotFocus
'-------------------------------------------------
' Purpose:  When the user clicks the Delete
'           button, the Index Value textbox
'           LostFocus event fires which in
'           turn disables the button. Therefore
'           the cmdDeleteIndex_Click event does
'           not occur and we must call it here.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    This is a side effect of the fact
'           that the Delete button is only
'           enabled when an Index Value is
'           selected (has focus).
'*************************************************
Private Sub cmdDeleteIndex_GotFocus()

        On Error GoTo Unexpected_Error
        
590     Call cmdDeleteIndex_Click
        
        Exit Sub

Unexpected_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdFileBrowse_Click
'-------------------------------------------------
' Purpose:  Initialize and display the dialog
'           allowing the user to browse for the
'           index data file.  Mark the data
'           dirty if the user selects a file.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdFileBrowse_Click()
    Dim Path As String
    Dim nTextLen As Integer
    Dim LastPos As Integer
    Dim TmpPos As Integer
    
        On Error GoTo TryAgain
    
600     dlgDialogs.Flags = dlgDialogs.Flags Or cdlOFNHideReadOnly
605     dlgDialogs.Filter = LoadResString(CAP_ALLFILES) & " (*.*)|*.*"
        
        ' Dialog should generate an error
        ' if the user presses Cancel
        dlgDialogs.CancelError = True
        
        ' Initialize the dialog to the file and path specified in the text box.
        ' If nothing is specified and no previous default directory has been
        ' set, default to C:\
        nTextLen = Len(txtFileName.Text)
        LastPos = 0
        If nTextLen <> 0 Then
            ' Find the last "\" character in the path
            Do
                TmpPos = InStr(LastPos + 1, txtFileName.Text, "\")
                If TmpPos <> 0 Then LastPos = TmpPos
            Loop Until TmpPos = 0
            
610         If InStr(LastPos + 1, txtFileName.Text, ".") Then
                ' The textbox includes a filename with an extension
                ' so we strip it off to initialize the dialog box
                ' to the correct subdirectory and file name
615             Path = RemoveLastPathSegment(txtFileName.Text)
620             dlgDialogs.FileName = txtFileName.Text
                If Path <> "" Then
625                 dlgDialogs.InitDir = Path
                End If
            Else
                ' We didn't find a file name so we assume the
                ' textbox contains a directory and we initialize
                ' the dialog box
630             dlgDialogs.FileName = ""
635             dlgDialogs.InitDir = txtFileName.Text
            End If
        Else
640         dlgDialogs.FileName = ""
            If dlgDialogs.InitDir = "" Then
645             dlgDialogs.InitDir = "C:\"
            End If
        End If
        
650     dlgDialogs.ShowOpen
        txtFileName = dlgDialogs.FileName
655     dlgDialogs.InitDir = RemoveLastPathSegment(dlgDialogs.FileName)
        Me.Dirty = True
        
        GoTo SkipIt
    
TryAgain:

        ' If the user didn't press the cancel key while
        ' selecting a text file, then clear the dialog
        ' initialization values and try again.  We do
        ' this in case an invalid filename or path is
        ' provided and the dialog fails to load the
        ' first time.  Clear the filename and path and
        ' try opening the dialog again.
        If Err <> cdlCancel Then
            dlgDialogs.FileName = ""
            dlgDialogs.InitDir = "C:\"
            On Error GoTo LogIt
            Resume
        End If
    

LogIt:
        If Err <> cdlCancel Then
            Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)
            txtFileName.Text = ""
        End If

SkipIt:
        ' Always reset the dialog for next time
        dlgDialogs.FileName = ""
        dlgDialogs.InitDir = ""
    
End Sub

'*************************************************
' cmdHelp_Click
'-------------------------------------------------
' Purpose:  Display the help topic for the tab
'           that is currently displayed.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    Each tab on the SSTab control has
'           a unique Help Context ID.  We add
'           the tab index to the first ID to
'           display the appropriate help info.
'           If additional tabs are added to the
'           SSTab control, their Help Context IDs
'           must be kept sequential.
'           This release script uses a proprietary
'           COM object for its help system.
'*************************************************
Private Sub cmdHelp_Click()

       Dim oKChmHlp As KChmHlp
       Dim bRetVal As Boolean
       Dim HelpFilePath As String

       Set oKChmHlp = New KChmHlp
       HelpFilePath = App.Path & "\" & App.HelpFile
       
       Call oKChmHlp.ShowHelp(ByVal HelpFilePath, CLng(tabText.Tab + TABS_FIRST_HELPID))
       'Call oKChmHlp.ShowHelp(ByVal HelpFilePath, "", CLng(tabText.Tab + TABS_FIRST_HELPID), 1, bRetVal)
End Sub

'*************************************************
' cmdImageBrowse_Click
'-------------------------------------------------
' Purpose:  Initialize and display the dialog
'           allowing the user to browse for the
'           directory where images will be
'           stored during Release.  Mark the
'           data dirty if the user selects a
'           directory.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    We store the Help Context ID in the
'           dialog's Tag property since it is
'           used for multiple purposes.
'*************************************************
Private Sub cmdImageBrowse_Click()
    Dim NewFolder As String
    
        On Error GoTo ImageBrowse_Error
        
660     NewFolder = BrowseFolders(Me.hwnd, LoadResString(TITLE_SELECTIMGDIR), BIF_RETURNONLYFSDIRS)
        
        If NewFolder <> "" Then
670         txtImageDir.Text = NewFolder
680         Me.Dirty = True
        End If
        
        Exit Sub

ImageBrowse_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdMenu_Click
'-------------------------------------------------
' Purpose:  This routine begins the sequence of
'           events that cause the popup menu to
'           appear just below the link box.
' Inputs:   Index   control array index
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_Click(Index As Integer)
    Static InProgress As Boolean
    
        On Error GoTo MenuClick_Error
        
        ' InProgress is used to keep this code from being
        ' re-entrant.  The buttons tag value can be "IP",
        ' "SKIP", or empty.  The value is set to skip if the
        ' user clicks the button while the menu is still up.
        ' This will keep the menu from popping back up.  The
        ' "IP" value is set if we are In Progress and is
        ' used by the MouseDown event.
700     If (InProgress = False And cmdMenu(Index).Tag <> SKIP_EVENT) Then
            ' Set our in progress flags.
            InProgress = True
710         cmdMenu(Index).Tag = IN_PROGRESS
            ' Move the focus back to the link box
720         txtIndexData(Index).SetFocus
            ' Start the indexing code.  This routine will
            ' not return until the popup menu goes away.
730         Call DoTheLink(txtIndexData(Index), Index + vsbIndex.Value)
            ' Allow events to fire.  If the user clicked on the
            ' menu button to drop the menu down, this will cause
            ' this routine to be fired re-entrantly.
            DoEvents
            ' Clear the in-progress flags
            InProgress = False
            ' If the tag is set to SKIP, let the other instance
            ' of this event clear the tag.
740         If (cmdMenu(Index).Tag <> SKIP_EVENT) Then
750             cmdMenu(Index).Tag = ""
            End If
        Else
            ' Clear the tag and set focus back on the link box.
760         cmdMenu(Index).Tag = ""
770         txtIndexData(Index).SetFocus
        End If
        
        Exit Sub

MenuClick_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdMenu_LostFocus
'-------------------------------------------------
' Purpose:  The menu button returns focus to
'           the link box after setting its own
'           Tag property to IN_PROGRESS in the
'           Click event. In that instance we do
'           nothing.  If the menu button loses
'           focus with a different Tag value,
'           then the Click event did not occur.
'           The user must have moved the mouse
'           off the menu button before doing a
'           Mouse Up and clicked a different
'           focus.  We therefore need to clean
'           up the link box so it no longer
'           looks like it has focus.
' Inputs:   Index
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_LostFocus(Index As Integer)
    
    If cmdMenu(Index).Tag <> IN_PROGRESS Then
        txtIndexData(Index).Tag = ""
        Call txtIndexData_LostFocus(Index)
    End If

End Sub

'*************************************************
' cmdMenu_MouseDown
'-------------------------------------------------
' Purpose:  This event handler has two purposes.
'           It checks to see if the popup menu
'           is up to keep the click event from
'           re-displaying the menu again.  It
'           also keeps the link boxes SetFocus
'           event from rerunning when the focus
'           is returning from the menu button.
' Inputs:   Index   control array index
'           Button  which mouse button was down
'           Shift   flag for Ctrl, Alt, Shift
'           x       horizontal position
'           y       vertical position
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        
        On Error GoTo Menu_MouseDown_Error
        
        ' If the menu button's Click event is in progress,
        ' tell the event that the next time it's fired
        ' to skip the event handling code
780     If cmdMenu(Index).Tag = IN_PROGRESS Then
790         cmdMenu(Index).Tag = SKIP_EVENT
        End If
        
        ' Tell the link textbox that the command
        ' button is returning the focus.
800     txtIndexData(Index).Tag = BUTTON_CLICK
        
        Exit Sub

Menu_MouseDown_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdOCRBrowse_Click
'-------------------------------------------------
' Purpose:  Initialize and display the dialog
'           allowing the user to browse for the
'           directory where OCR Full Text files
'           will be stored during Release. Mark
'           the data dirty if the user selects
'           a directory.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    We store the Help Context ID in the
'           dialog's Tag property since it is
'           used for multiple purposes.
'*************************************************
Private Sub cmdOCRBrowse_Click()
    Dim NewFolder As String
    
        On Error GoTo OCRBrowse_Error
        
810     NewFolder = BrowseFolders(Me.hwnd, LoadResString(TITLE_SELECTOCRDIR), BIF_RETURNONLYFSDIRS)
        
        If NewFolder <> "" Then
820         txtOCRDir.Text = NewFolder
830         Me.Dirty = True
        End If
        
        Exit Sub

OCRBrowse_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' cmdOK_Click
'-------------------------------------------------
' Purpose:  Validates the settings if they are
'           dirty and saves them before exiting
'           from the Release Setup script.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Note:     The first time the Release Setup
'           script is run for a Doc Class, it
'           is not marked Dirty initially.
'           Therefore the OK button still needs
'           to validate the settings because the
'           defaults are not complete.  However,
'           once the data has been changed and
'           validated, the fVerified flag is set
'           True.  This means that OK only needs
'           to verify settings when the form is
'           Dirty from that point forward.
'*************************************************
Private Sub cmdOK_Click()
        
        ' When the user is in the middle of editing a
        ' a Text Constant, the ENTER key triggers the
        ' OK button's click event.  We should simply
        ' accept the string and lock the text box.
        If ActiveControl.Name = "txtIndexData" Then
            ' Check if the Index Value text box is unlocked
            If ActiveControl.Locked = False Then
                ' Lock the Text Constant text box
900             Call LockTextConstant(ActiveControl.Index, True)
                Exit Sub
            Else
                ' Set focus to the OK button.  The call to
                ' DoEvents is required to process the
                ' txtIndexData_LostFocus event which forces the
                ' selected Index Value to lose focus and clean
                ' up its highlighting and menu button.
                cmdOK.SetFocus
                DoEvents
            End If
        End If
        
        On Error GoTo ClickOK_Error
        
        ' We always need to verify the settings if the
        ' form is dirty, or the setting is not checked
        ' yet (e.g., the imported batch class).
        If Me.Dirty Or (Not fVerified) Then
910         If Not VerifyReleaseSettings() Then Exit Sub
        End If
        
920     If (SaveReleaseSettings()) Then
930         Unload Me
        End If
    
        Exit Sub

ClickOK_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' DeleteAllIndex
'-------------------------------------------------
' Purpose:  This routine will delete all Index
'           Values from the list of links.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub DeleteAllIndex()
    
    ' Discard all Index Values
    fIndexCount = 0
    fSelectedIndex = 0
    ReDim fIndexList(0)

End Sub

'*************************************************
' DeleteIndex
'-------------------------------------------------
' Purpose:  This routine deletes the specified
'           Index Value from the list of links
' Inputs:   Index   index into the links array
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub DeleteIndex(Index As Integer)
    Dim I As Integer
    
        On Error GoTo DeleteIndex_Error
        
        ' If there is more than one Index Value in the list
        ' and this is not the last Index Value in the list,
        ' shift the remaining values up one position
        If (fIndexCount > 1) And (Index <> fIndexCount - 1) Then
            For I = Index To fIndexCount - 2
950             fIndexList(I).Source = fIndexList(I + 1).Source
960            fIndexList(I).SourceType = fIndexList(I + 1).SourceType
            Next
        End If
        
        ' Decrement the number of Index Values
        fIndexCount = fIndexCount - 1
        
        ' Resize the list of Index Values
        If fIndexCount > 0 Then
970         ReDim Preserve fIndexList(fIndexCount - 1)
        Else
            ReDim fIndexList(0)
        End If
        
        Exit Sub

DeleteIndex_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' DisableUnsupportedControl
'-------------------------------------------------
' Purpose:  This routine disables all controls
'           that are not supported by the running
'           product.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    This is necessary to prevent the user
'           confusion over features that are not
'           supported.
'*************************************************
Private Sub DisableUnsupportedControls()

    ' Get the running product and if it is
    ' Titan hide the OCR full text controls
    
    If GetProductCode = pcTitan Then
        txtOCRDir.Visible = False
        fraOCRFiles.Visible = False
        cmdOCRBrowse.Visible = False
        lblOCRDir.Visible = False
    End If

End Sub
 
'*************************************************
' DisplayIndexValues
'-------------------------------------------------
' Purpose:  This routine displays Index Values
'           and their output sequence in the link
'           box.  The caller specifies where in
'           the links array to start displaying.
' Inputs:   Index   specifies the starting point
'                   in the links array
' Outputs:  None
' Returns:  None
' Notes:    Text Constants are enclosed in quotes
'           when they are not edittable and all
'           Ascent Capture system-provided values
'           are displayed in bold.
'*************************************************
Sub DisplayIndexValues(Index As Integer)
    Dim I As Integer
    
        On Error GoTo DisplayIndexValues_Error
        
        ' Fill the link box display
        For I = 0 To LINK_BOX_SIZE - 1
            ' If our index is within the link array,
            ' display the array values
            If (I + Index) < fIndexCount Then
                ' Make the sequence column text box visible
                ' and fill it with the sequence number
1000            txtSequence(I).Visible = True
1010            txtSequence(I).Text = fIndexList(I + Index).Destination + 1
                                                       
                ' Display the link data box and fill it with
                ' the appropriate Index Value
1020            With txtIndexData(I)
                    .Visible = True
1030                Select Case fIndexList(I + Index).SourceType
                        Case NO_LINK
                            .FontBold = False
                            .Text = ""
                        Case KFX_REL_TEXTCONSTANT
                            .FontBold = False
1040                        .Text = """" & fIndexList(I + Index).Source & """"
                        Case KFX_REL_VARIABLE
                            .FontBold = True
1050                        .Text = "{" & fIndexList(I + Index).Source & "}"
                        Case KFX_REL_INDEXFIELD
                            .FontBold = True
1060                        .Text = fIndexList(I + Index).Source
                        Case KFX_REL_BATCHFIELD
                            .FontBold = True
1065                        .Text = "{$" & fIndexList(I + Index).Source & "}"
                        Case KFX_REL_DOCUMENTID
                            .FontBold = True
                            .Text = "{" & fIndexList(I + Index).Source & "}"
                    End Select
                End With
            Else
                ' Otherwise hide the sequence number and the data box
1070            txtSequence(I).Visible = False
1080            txtSequence(I).Text = ""
                
1090            txtIndexData(I).Visible = False
1100            txtIndexData(I).Text = ""
            End If
        Next I
        
        ' Enable or Disable the <Delete All> button
        cmdDeleteAllIndex.Enabled = (fIndexCount > 0)
        
        ' Update the scrollbar to represent the display
1110    SetScrollBar (Index)
        
        Exit Sub

DisplayIndexValues_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' DoTheLink
'-------------------------------------------------
' Purpose:  This routine performs the steps to
'           allow a user to select a link value.
' Inputs:   LBox    selected text box
'           Index   array index of the link
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub DoTheLink(LBox As TextBox, Index As Integer)
    Dim BoxLeft As Integer
    Dim BoxTop As Integer
    Dim Results As Integer
        
        On Error GoTo DoTheLink_Error
        
        ' Initialize the return value to use no selection.  Handles when
        ' user clicks away from the menu without making a menu selection.
        gNewIndexType = NO_SELECTION
        
        ' Pop up the menu just below the textbox
1150    Call GetBoxPosition(LBox, BoxLeft, BoxTop)
1160    Call PopupMenu(oMenu.mnuLinks, 0, BoxLeft, BoxTop)
        
        ' After the menu returns, establish the new link
        If gNewIndexType <> DELETE_LINK And _
           gNewIndexType <> NO_SELECTION Then
1170        Results = EstablishLink(Index)
            If Results = UNLOCK_TEXT_BOX Then
                ' The user is entering a Text Constant.  Unlock the text box,
                ' remove the highlighting, and use the default cursor.
                With LBox
1180                .Text = fIndexList(Index).Source
                    .Locked = False
                    .MousePointer = vbDefault
                    .BackColor = vbWindowBackground
                    .ForeColor = vbWindowText
                    .FontBold = False
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            Else
1190            Call DisplayIndexValues(vsbIndex.Value)
            End If
        End If
        
        Exit Sub

DoTheLink_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'************************************************
'*** Routine: EnableAdobeAcrobatSettings
'*** Purpose: Enable/disable the Settings button for
'***    the Adobe Acrobat PDF . The button is enabled
'***    whenever the cboImageType contains the
'***    Adobe Acrobat PDF selection.
'************************************************
Private Sub EnableAdobeAcrobatSettings()
    
    On Error Resume Next
    
    ' Only enable if the image type is one of the following Adobe Acrobat PDFs (3.0)
    cmdSettings.Enabled = (((cboImageType.ItemData(cboImageType.ListIndex) = CAP_FORMAT_PDF_SINGLE) Or _
        (cboImageType.ItemData(cboImageType.ListIndex) = CAP_FORMAT_PDF_MULTI) Or _
        (cboImageType.ItemData(cboImageType.ListIndex) = CAP_FORMAT_PDF_JPEG) Or _
        (cboImageType.ItemData(cboImageType.ListIndex) = CAP_FORMAT_PDF_PCX))) And CBool(chkReleaseImageFiles.Value)
        frmAdobeAcrobatSetup.pdfImageFormat3.Enabled = cmdSettings.Enabled
        frmAdobeAcrobatSetup.pdfImageFormat3.TabStop = frmAdobeAcrobatSetup.pdfImageFormat3.Enabled And (tabText.Tab = IMAGE_TAB)
End Sub

'*************************************************
' EstablishLink
'-------------------------------------------------
' Purpose:  This routine will build the link
'           between a destination value and
'           the selected index data
' Inputs:   Index   index to the links array
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Function EstablishLink(Index As Integer) As Integer

        ' Default to locking the text box on the
        ' conclusion of this routine.  Only the
        ' selection of a Text Constant link will
        ' unlock the text box.
        EstablishLink = LOCK_TEXT_BOX
    
        ' Depending upon the type, handle source data value.
        Select Case gNewIndexType
            Case KFX_REL_TEXTCONSTANT
                ' Save the previous link info in
                ' case the user presses <ESC> to
                ' discard their changes
                fSavedLink.SourceType = fIndexList(Index).SourceType
                fSavedLink.Source = fIndexList(Index).Source
                ' If the current source type is not
                ' Text Constant, initialize the text
1300            If fIndexList(Index).SourceType <> KFX_REL_TEXTCONSTANT Then
1310                fIndexList(Index).Source = ""
                End If
                ' Unlock the text box and wait until
                ' the user has entered the text string
                EstablishLink = UNLOCK_TEXT_BOX
            
            Case KFX_REL_BATCHFIELD, _
                 KFX_REL_INDEXFIELD, _
                 KFX_REL_VARIABLE, _
                 KFX_REL_DOCUMENTID
                ' For the above just store the field name as the source
1320            fIndexList(Index).Source = gNewIndexData
        End Select
        
        ' Set the new index type
1330    fIndexList(Index).SourceType = gNewIndexType
        
        Exit Function

EstablishLink_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Function

'*************************************************
' FatalErrorExit
'-------------------------------------------------
' Purpose:  This routine will unload the form
'           and set the status to FATAL_ERROR
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    This routine may be called from any
'           of the error handlers in the form
'           routines
'*************************************************
Public Function FatalErrorExit() As Integer

        Me.FormStatus = rtFATAL_ERROR
        Unload Me
        
End Function

'*************************************************
' FillUIWithCaptions
'-------------------------------------------------
' Purpose:  This routine will load the captions
'           for all of the controls on the UI
'           from the resource file
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub FillUIWithCaptions()

    On Error GoTo FillCaptions_Error
    
        ' -- Main Form --
            
        ' Form title
1400    'Caption = LoadResString(TITLE_FORM)
        
        ' Batch Class, Doc Class and Destination Name labels
1405    lblBatchClass.Caption = LoadResString(CAP_BATCH_CLASS)
1410    lblDocClass.Caption = LoadResString(CAP_DOC_CLASS)
        lblName.Caption = LoadResString(CAP_NAME)
            
        ' Bottom buttons
1415    cmdOK.Caption = LoadResString(CMD_OK)
1420    cmdCancel.Caption = LoadResString(CMD_CANCEL)
1425    cmdApply.Caption = LoadResString(CMD_APPLY)
1430    cmdHelp.Caption = LoadResString(CMD_HELP)
    
        ' Main Tab control
1440    tabText.Tab = INDEX_TAB
1442    tabText.Caption = LoadResString(CAP_TAB_INDEX)
1444    tabText.Tab = DOCUMENT_TAB
1446    tabText.Caption = LoadResString(CAP_TAB_DOCUMENT)
1448    tabText.Tab = IMAGE_TAB
1450    tabText.Caption = LoadResString(CAP_TAB_IMAGEFORMAT)
        
        ' -- Index Storage Tab --
        
1455    lblFileName.Caption = LoadResString(CAP_FILENAME)
1460    cmdFileBrowse.Caption = LoadResString(CMD_BROWSE)
            
        ' Index Value frame captions
1465    fraIndexVals.Caption = LoadResString(CAP_IDXFRAME)
1470    lblIndexLabel(0).Caption = LoadResString(CAP_SEQUENCE)
1475    lblIndexLabel(1).Caption = LoadResString(CAP_INDEXVALUE)
1480    cmdAddIndex.Caption = LoadResString(CMD_ADDINDEX)
1485    cmdDeleteIndex.Caption = LoadResString(CMD_DELETE)
1490    cmdDeleteAllIndex.Caption = LoadResString(CMD_DELETEALL)
1495    lblMove.Caption = LoadResString(CAP_MOVE)
        
        ' -- Document Storage Tab --
        
        ' Image File Frame Contents
1500    fraImageFiles.Caption = LoadResString(CAP_IMGFRAME)
        chkReleaseImageFiles.Caption = LoadResString(CAP_RELEASE_IMAGE_FILES)
1530    lblImageDir.Caption = LoadResString(CAP_RELDIRNAME)
1535    cmdImageBrowse.Caption = LoadResString(CMD_BROWSE)
1540    chkSkipFirstPage.Caption = LoadResString(CAP_SKIPFIRST)
    
        ' OCR File Frame Contents
1575    fraOCRFiles.Caption = LoadResString(CAP_OCRFRAME)
        chkReleaseOCRFullText.Caption = LoadResString(CAP_RELEASE_FULL_TEXT)
1580    lblOCRDir.Caption = LoadResString(CAP_OCRRELDIR)
1585    cmdOCRBrowse.Caption = LoadResString(CMD_BROWSE2)
        
        ' PDF File Frame Contents
        fraPDFFiles.Caption = LoadResString(CAP_PDFFRAME)
        chkReleaseKofaxPDF.Caption = LoadResString(CAP_RELEASE_KOFAX_PDF_FILES)
        lblPDFDir.Caption = LoadResString(CAP_PDFRELDIR)
        cmdPDFBrowse.Caption = LoadResString(CMD_BROWSE3)
        
        ' -- Image Format Tab --
            
1590    lblImageType.Caption = LoadResString(CAP_IMGTYPE)
        cmdSettings.Caption = LoadResString(CMD_SETTINGS)
1592    fraReleaseImagesAs = LoadResString(CAP_RELEASEFILESAS)
        fstrPDFImageTypes(0) = LoadResString(CAP_PDFFORMAT_JPEG)
        fstrPDFImageTypes(1) = LoadResString(CAP_PDFFORMAT_MTIFF)
        fstrPDFImageTypes(2) = LoadResString(CAP_PDFFORMAT_PCX)
        fstrPDFImageTypes(3) = LoadResString(CAP_PDFFORMAT_TIFF)

       ' -- Adobe Acrobat PDF Settings Dialog --
1595    frmAdobeAcrobatSetup.Caption = LoadResString(TITLE_ADOBEACROBAT)
        
        ' PDF Related Settings Contents
1596    frmAdobeAcrobatSetup.frmAdvancedPDFSettings.Caption = LoadResString(CAP_FRA_ADV_PDF_SETTINGS)
        frmAdobeAcrobatSetup.chkWaitForStatus.Caption = LoadResString(CAP_CHK_WAIT_STATUS)
        frmAdobeAcrobatSetup.chkDeleteOnHung.Caption = LoadResString(CAP_CHK_DEL_HUNG)
        frmAdobeAcrobatSetup.cmdHelp.Caption = LoadResString(CMD_HELP)
      
        Exit Sub

FillCaptions_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' FillUIWithDefaults
'-------------------------------------------------
' Purpose:  If the user is editing the release
'           settings for the first time on this
'           batch class/doc class instance,
'           fill the user interface with
'           default values rather than reading
'           the settings in the data object.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub FillUIWithDefaults()
    Dim Results As Integer
    
        On Error GoTo FillDefaults_Error
        
        '*** Init and Default to Adobe Acrobat Capture 3.0 control
        frmAdobeAcrobatSetup.pdfImageFormat3.Visible = True
        frmAdobeAcrobatSetup.pdfImageFormat3.Enabled = True
        frmAdobeAcrobatSetup.pdfImageFormat3.Init SetupData
        frmAdobeAcrobatSetup.pdfImageFormat3.RestoreDefaults
        
        ' --- Release destination name ---
        txtName.Text = ""
        
        ' --- Index Storage Tab ---
        txtFileName.Text = ""
1600    Call InitializeIndexValues(SetupData)
        
        ' --- Document Storage Tab ---
            
        ' The directory text boxes are cleared
        ' and skip first page is disabled.
        chkReleaseImageFiles.Value = vbChecked
        txtImageDir.Text = ""
       
       ' --- Disable OCR Full text ---
        chkReleaseOCRFullText.Value = vbUnchecked
        lblOCRDir.Enabled = False
        txtOCRDir.Text = ""
        txtOCRDir.Enabled = False
        cmdOCRBrowse.Enabled = False
        
        ' --- Disable Kofax PDF ---
        chkReleaseKofaxPDF.Value = vbUnchecked
        lblPDFDir.Enabled = False
        txtKofaxPDFDir.Text = ""
        txtKofaxPDFDir.Enabled = False
        cmdPDFBrowse.Enabled = False
       
        chkSkipFirstPage.Value = vbUnchecked
        frmAdobeAcrobatSetup.chkWaitForStatus.Value = IIf(frmAdobeAcrobatSetup.pdfImageFormat3.PDFWaitForStatus, vbChecked, vbUnchecked)
        frmAdobeAcrobatSetup.chkDeleteOnHung.Value = IIf(frmAdobeAcrobatSetup.pdfImageFormat3.PDFDeleteHungDoc, vbChecked, vbUnchecked)
       

        ' Disable controls that are not supported by the running product
        DisableUnsupportedControls
        
        ' --- Image Format Tab ---
        
        ' Defaults to multi-page TIFF release
        Call SetImageType(CAP_FORMAT_MTIFF_G4)
        
        Exit Sub

FillDefaults_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' FillUIWithImageType
'-------------------------------------------------
' Purpose:  This routine gets all image types
'           from the SetupData object and fills
'           the combo box with the description
'           and ID
' Inputs:   oSetupData      ReleaseSetupData object
'           bMultiPageOnly  flag to list only
'                           multipage image formats
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub FillUIWithImageType(oSetupData As ReleaseSetupData, bMultiPageOnly As Boolean)
    Dim oImageType As Object

        On Error GoTo FillImageType_Error
        
        ' The Ascent Titan Product does not support JPEG so we
        ' need to remove this option
        
        Dim eProductCode As ProductCode
        eProductCode = GetProductCode
        
        With cboImageType
            ' Start with an empty combo box
            .Clear
            ' Get each item and add it
            For Each oImageType In oSetupData.ImageTypes
            
                If (Not bMultiPageOnly Or oImageType.MultiplePage) Then
                
                    '  Do not put up color options for the Titan product
                    If Not (eProductCode = pcTitan And _
                           (oImageType.Type = CAP_FORMAT_MTIFF_G4 Or _
                            oImageType.Type = CAP_FORMAT_TIFF_JPEG Or _
                            oImageType.Type = CAP_FORMAT_JPG_JPEG)) Then

1650                    .AddItem oImageType.Description
1660                    .ItemData(.NewIndex) = oImageType.Type
                    End If
                End If
            Next
            
            ' Add PDF 3.0 image types
            .AddItem fstrPDFImageTypes(0)
            .ItemData(.NewIndex) = CAP_FORMAT_PDF_JPEG
                        
            .AddItem fstrPDFImageTypes(1)
            .ItemData(.NewIndex) = CAP_FORMAT_PDF_MULTI
            
            .AddItem fstrPDFImageTypes(2)
            .ItemData(.NewIndex) = CAP_FORMAT_PDF_PCX
            
            .AddItem fstrPDFImageTypes(3)
            .ItemData(.NewIndex) = CAP_FORMAT_PDF_SINGLE
        End With
        
        Exit Sub

FillImageType_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' FillUIWithSettings
'-------------------------------------------------
' Purpose:  This routine will fill the user
'           interface with the current release
'           settings for this batch class/doc
'           class combination from SetupData.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub FillUIWithSettings()
    Dim I As Integer
        
        On Error GoTo FUWS_SetImageFilesDefault
        chkReleaseImageFiles.Value = IIf(SetupData.CustomProperties.Item(KEY_DISABLE_IMAGE_EXPORT).Value, vbUnchecked, vbChecked)
        
        On Error GoTo FUWS_SetFullTextDefault
        chkReleaseOCRFullText.Value = IIf(SetupData.CustomProperties.Item(KEY_DISABLE_TEXT_EXPORT).Value, vbUnchecked, vbChecked)
    
        On Error GoTo FUWS_SetKofaxPDFDefault
        chkReleaseKofaxPDF.Value = IIf(SetupData.KofaxPDFReleaseScriptEnabled, vbChecked, vbUnchecked)
        On Error GoTo FUWS_LogAndPropError

        With SetupData
            ' -- Release Destination Name --
            txtName.Text = .Name

            ' -- Index Storage Tab --
1700        txtFileName.Text = .CustomProperties(KEY_ASCIIFILE).Value
1710        Call LoadIndexValues(.Links)
            
            ' -- Document Storage Tab --
1720        txtImageDir.Text = .ImageFilePath
1730        txtOCRDir.Text = .TextFilePath
1740        txtKofaxPDFDir.Text = .KofaxPDFPath

            ' Disable controls that are not supported by the running product
            DisableUnsupportedControls
            
            chkSkipFirstPage = IIf(.SkipFirstPage, vbChecked, vbUnchecked)
            
              ' Adobe Acrobat Capture Setup
            On Error GoTo FUWS_AdobeWaitForStatus
            frmAdobeAcrobatSetup.chkWaitForStatus.Value = IIf(CBool(.CustomProperties.Item(KEY_ADOBE_WAIT_FOR_STATUS).Value), vbChecked, vbUnchecked)
            On Error GoTo FUWS_AdobeDeleteOnHung
            frmAdobeAcrobatSetup.chkDeleteOnHung.Value = IIf(CBool(.CustomProperties.Item(KEY_ADOBE_DELETE_HUNG).Value), vbChecked, vbUnchecked)
            On Error GoTo FUWS_LogAndPropError
            
            
            ' Refresh controls to their proper states
            If chkReleaseImageFiles.Value = vbUnchecked Then
                Call chkReleaseImageFiles_Click
            End If
            
            If chkReleaseOCRFullText.Value = vbUnchecked Then
                Call chkReleaseOCRFullText_Click
            End If
            
            If chkReleaseKofaxPDF.Value = vbUnchecked Then
                Call chkReleaseKofaxPDF_Click
            End If

            
            ' -- Image Format Tab --

            ' Selected the saved image selection otherwise use default multi-page TIFF image format
            If CBool(chkReleaseImageFiles.Value) Then
                 ' Warn the users if PDF 2.01 is detected, otherwise use the saved image settings
                If .ImageType = CAP_FORMAT_PDF Then
                    Call MsgBox(LoadResString(MSG_PDFSETUPFAILED1) & vbCrLf & _
                        LoadResString(MSG_PDFSETUPFAILED2), vbOKOnly + vbExclamation, _
                        LoadResString(TITLE_DATAVERIFYFAIL))
                    Call SetImageType(CAP_FORMAT_MTIFF_G4)                      ' Use default image format
                    Call frmAdobeAcrobatSetup.pdfImageFormat3.Init(SetupData)   ' Initialize PDF control
                    Call frmAdobeAcrobatSetup.pdfImageFormat3.RestoreDefaults   ' Clear PDF settings
                Else
                    Call SetImageType(.ImageType)                               ' Select saved PDF image format
                    Call frmAdobeAcrobatSetup.pdfImageFormat3.Init(SetupData)   ' Initialize PDF control
                End If
            Else
                Call SetImageType(CAP_FORMAT_MTIFF_G4)                          ' Use default image format
                Call frmAdobeAcrobatSetup.pdfImageFormat3.Init(SetupData)       ' Initialize PDF control
                Call frmAdobeAcrobatSetup.pdfImageFormat3.RestoreDefaults       ' Clear PDF settings
            End If
            
            'ODBC Connection
            txtDSN.Text = .CustomProperties("ODBC DSN").Value
            txtDB.Text = .CustomProperties("ODBC DB").Value
            chkWinAuth.Value = IIf(.CustomProperties("WinAuth").Value = "True", vbChecked, vbUnchecked)
            txtUID.Text = IIf(.CustomProperties("WinAuth").Value = "False", .CustomProperties("ODBC UID").Value, "")
            txtUID.Enabled = IIf(chkWinAuth.Value = vbChecked, False, True)
            txtPWD.Text = IIf(.CustomProperties("WinAuth").Value = "False", .CustomProperties("ODBC PWD").Value, "")
            txtPWD.Enabled = IIf(chkWinAuth.Value = vbChecked, False, True)
        End With
        
    Exit Sub
    
FUWS_SetImageFilesDefault:
    chkReleaseImageFiles.Value = vbChecked
    Resume Next

FUWS_SetFullTextDefault:
    chkReleaseOCRFullText.Value = vbChecked
    Resume Next
    
FUWS_SetKofaxPDFDefault:
    chkReleaseKofaxPDF.Value = vbChecked
    
FUWS_AdobeWaitForStatus:
    frmAdobeAcrobatSetup.chkWaitForStatus.Value = vbChecked
    Resume Next
    
FUWS_AdobeDeleteOnHung:
    frmAdobeAcrobatSetup.chkDeleteOnHung.Value = vbUnchecked
    Resume Next

FUWS_LogAndPropError:
        
    Dim iRet As Integer
        
    iRet = MsgBox(LoadResString(MSG_USEDEFAULT), _
                  vbExclamation + vbYesNo, _
                  LoadResString(TITLE_RSETUPERROR))
    If (iRet = vbYes) Then
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, False)
        Call Err.Clear
        Call FillUIWithDefaults
    Else
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)
    End If
    
End Sub

'*************************************************
' FindIndexValue
'-------------------------------------------------
' Purpose:  This function loops through the list
'           of Index Values for a specified
'           source name and type
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    -1 if Index Value was NOT found
'           Otherwise the index in fIndexList
'*************************************************
Private Function FindIndexValue(sName As String, IndexType As KfxLinkSourceType) As Integer
    Dim I As Integer

        On Error GoTo FIV_Error
        
        ' Initialize to indicate not found
        FindIndexValue = -1
        
        ' Search through the list
        For I = 0 To fIndexCount - 1
1800        If (fIndexList(I).Source = sName And _
                        (IndexType = -1 Or _
                        IndexType = fIndexList(I).SourceType)) Then
                ' We found it
                FindIndexValue = I
                Exit Function
            End If
        Next
            
        Exit Function

FIV_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Function

'*************************************************
' cmdPDFBrowse_Click
'-------------------------------------------------
' Purpose:  Initialize and display the dialog
'           allowing the user to browse for the
'           directory where Kofax PDF files
'           will be stored during Release. Mark
'           the data dirty if the user selects
'           a directory.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    We store the Help Context ID in the
'           dialog's Tag property since it is
'           used for multiple purposes.
'*************************************************
Private Sub cmdPDFBrowse_Click()
    Dim NewFolder As String
    
        On Error GoTo PDFBrowse_Error
        
     NewFolder = BrowseFolders(Me.hwnd, LoadResString(TITLE_SELECTPDFDIR), BIF_RETURNONLYFSDIRS)
        
        If NewFolder <> "" Then
         txtKofaxPDFDir.Text = NewFolder
         Me.Dirty = True
        End If
        
        Exit Sub

PDFBrowse_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

Private Sub cmdSettings_Click()
    frmAdobeAcrobatSetup.Show vbModal
End Sub

'*************************************************
' Form_Activate
'-------------------------------------------------
' Purpose:  Initialize the state of the form and
'           set focus to the index tab.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Activate()

        On Error GoTo FormActivate_Error
        
        ' If this is the first time the form is
        ' being displayed, set the form status
        ' to OK and display the Index Tab
        If Me.FormStatus = rtLOADING Then
            Me.FormStatus = rtOK
1830        tabText.Tab = INDEX_TAB
            txtName.SelStart = Len(txtName.Text)
            txtName.SetFocus
        End If

        Exit Sub

FormActivate_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' Form_KeyDown
'-------------------------------------------------
' Purpose:  Allow the form to check keystrokes.
'           Ctrl+Tab and Ctrl+Shift+Tab move
'           between tabs on the SSTab control.
'           F1 displays online help for the
'           current tab on the SSTab Control.
' Inputs:   KeyCode   the key that was pressed
'           Shift     flags for Alt, Shift, Ctrl
' Outputs:  None
' Returns:  None
' Notes:    KeyPreview property must be True
'*************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        On Error GoTo FKD_Error

        If KeyCode = vbKeyTab Then
            If Shift = vbCtrlMask Then
                If tabText.Tab = LAST_TAB Then
                    ' Wrap around to the first tab
1860                tabText.Tab = FIRST_TAB
                Else
                    ' Move to the next tab
1870                tabText.Tab = tabText.Tab + 1
                End If
            ElseIf Shift = vbCtrlMask + vbShiftMask Then
                If tabText.Tab = FIRST_TAB Then
                    ' Wrap around to the last tab
1880                tabText.Tab = LAST_TAB
                Else
                    ' Move to the previous tab
1890                tabText.Tab = tabText.Tab - 1
                End If
            End If
        ElseIf KeyCode = vbKeyF1 Then
            ' F1 usually displays the help context ID specified
            ' by the control with focus.  We are overriding that
            ' functionality so that F1 always shows help for the
            ' current tab.  We also ignore the F1 key when any
            ' control key modifier is selected
            If Shift = 0 Then cmdHelp_Click
            KeyCode = 0
        End If
    
        Exit Sub

FKD_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' Form_Load
'-------------------------------------------------
' Purpose:  Initialize the user interface and add
'           internationalization support.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Load()
    
        On Error GoTo FormLoad_Error
        
        Me.FormStatus = rtLOADING
        
        ' Allow international support
        Dim InternationalEngine As New CInternationalSupport
        InternationalEngine.FixAllFonts Me
            
        ' Handle initialization of the form lables and captions
1900    Call FillUIWithCaptions

        Me.Dirty = False
        Exit Sub
        
FormLoad_Error:
        
    ' If an error occurs during the load, log the error
    ' and set the form status to FATAL ERROR
    Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, False)
    Me.FormStatus = rtFATAL_ERROR
End Sub

'*************************************************
' Form_Paint
'-------------------------------------------------
' Purpose:  When this form is first shown, it
'           sometimes starts in the background.
'           We activate the form via AppActivate
'           during the first paint event.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    If we did not limit the activate to
'           the first paint event the form would
'           pop up whenever a large screen covers
'           it and is then moved.
'*************************************************
Private Sub Form_Paint()
    Static Doneit As Boolean
    
    If (Not Doneit) Then
        AppActivate Me.Caption
        Doneit = True
    End If
    
End Sub

'*************************************************
' Form_QueryUnload
'-------------------------------------------------
' Purpose:  This event is called first whenever
'           the form is about to unload. When the
'           user clicks OK or Cancel we start to
'           unload the form.  In this event, we
'           simply validate that all changes are
'           saved and hide the form.  The form is
'           actually unloaded by the ReleaseSetup
'           class.  That time, the form is not
'           visible and we allow it to unload.
' Inputs:   None
' Outputs:  Cancel      flag to abort Unload event
'           UnloadMode  cause of the Unload event
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Results As Integer

        On Error GoTo FQU_Error

        ' If the form is visible then we only validate
        ' that the data is saved and then hide the form.
        ' We do not allow it to unload yet.
        If Me.Visible Then
        
            ' Don't let the form unload
            Cancel = True
            
            ' Check the form status and if changes have been made,
            ' allow the user to save.  Otherwise just exit.
            If Me.Dirty = True Then
                Results = MsgBox(LoadResString(MSG_SAVESETTINGS), _
                                vbYesNoCancel + vbQuestion, _
                                LoadResString(TITLE_SAVESETTINGS))
                
                If Results = vbYes Then
                    ' Try and save
1920                If (VerifyReleaseSettings()) Then
1930                    Call SaveReleaseSettings
                    Else
                        Exit Sub
                    End If
                ElseIf Results = vbCancel Then
                    ' Go back to the form
                    Exit Sub
                End If
            End If
            
            Me.Hide
        End If

        Exit Sub

FQU_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' Form_Unload
'-------------------------------------------------
' Purpose:  Clean up any data objects that were
'           allocated by this form.
' Inputs:   None
' Outputs:  Cancel  flag to abort Unload event
' Returns:  None
' Notes:    None
'*************************************************
Private Sub Form_Unload(Cancel As Integer)

        On Error Resume Next
        
1940    Unload oMenu
1950    Set SetupData = Nothing
1960    Set oTextFile = Nothing
1970    Set oMenu = Nothing

End Sub

'*************************************************
' GetBoxPosition
'-------------------------------------------------
' Purpose:  This routine will find a control's
'           absolute X and Y coordinates on the
'           main form.  It takes into account
'           all containers.
' Inputs:   TheControl      specified control
' Outputs:  Left            x coordinate
'           Top             y coordinate
' Returns:  None
' Notes:    This routine must be called ByRef.
'           The left and top values will not be
'           returned if it is called ByVal.
'*************************************************
Sub GetBoxPosition(TheControl As Control, left As Integer, top As Integer)
    Dim TheParent As Object
    
    On Error GoTo GBP_Error

        ' Initialize variables and get the control's
        ' position within its current container
2000    Set TheParent = TheControl
        left = TheControl.left
        top = TheControl.top
            
        ' Continue loop while parent control is an object.
        ' This should always be true but test just in case.
2010    Do While IsObject(TheParent)
            ' If the current object is the form, exit the loop
2020        If TypeOf TheParent.Container Is Form Then
                Exit Do
            Else
                ' Otherwise, set TheParent variable to
                ' the current controls contain and add
                ' this controls left and top values to
                ' our current values.  This loop will take
                ' care of containers that reside in containers.
2030            Set TheParent = TheParent.Container
                left = left + TheParent.left
                top = top + TheParent.top
            End If
        Loop
            
        ' Now set where in the text box the upper left
        ' corner of the menu will appear.  Currently set
        ' to appear left aligned and just below the text box.
        top = top + (TheControl.Height)
        left = left
        
        Exit Sub

GBP_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' GetProductCode
'-------------------------------------------------
' Purpose:  Determine the running product.
' Inputs:   None
' Outputs:  Enum for running product
' Returns:  None
' Notes:    None
'*************************************************
Private Function GetProductCode() As ProductCode
    
        On Error GoTo ProductCode_Error
    
        Dim oGenericSetupData As Object
        Set oGenericSetupData = SetupData
    
        ' Assume Ascent Capture. Ascent Capture currently
        ' does not support these properties.

        GetProductCode = pcAscentCapture

        On Error Resume Next

        ' Now read the code from the SetupData object
    
1905    GetProductCode = oGenericSetupData.ProductCode
        Exit Function
        
ProductCode_Error:
    ' A failure here means we had a bad SetuData object
    Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)
End Function

'*************************************************
' GetProductName
'-------------------------------------------------
' Purpose:  Return the running product title
' Inputs:   None
' Outputs:  None
' Returns:  Title of running product
' Notes:    None
'*************************************************
Private Function GetProductName() As String
    
        On Error GoTo ProductName_Error
    
        Dim oGenericSetupData As Object
        Set oGenericSetupData = SetupData
    
        ' Assume Ascent Capture. Ascent Capture currently
        ' does not support these properties.

        GetProductName = LoadResString(TITLE_FORM)
        On Error Resume Next
    
        ' Now read the code and name from the late bound
        ' SetupData object

1907    GetProductName = oGenericSetupData.ProductName
        Exit Function
        
ProductName_Error:
     Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)
End Function

'*************************************************
' InitializeIndexValues
'-------------------------------------------------
' Purpose:  This routine will initialize the
'           Index Values with all Batch Fields
'           and Index Fields
' Inputs:   oSetupData  ReleaseSetupData object
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub InitializeIndexValues(oSetupData As ReleaseSetupData)
    Dim BField As AscentRelease.BatchField
    Dim IField As AscentRelease.IndexField
    Dim I As Integer
    
        On Error GoTo IIV_LogAndPropError
            
        ' Resize the list of Index Values
        fIndexCount = oSetupData.BatchFields.Count + oSetupData.IndexFields.Count
        If fIndexCount > 0 Then
2050        ReDim fIndexList(fIndexCount - 1)
        Else
            ReDim fIndexList(0)
        End If
        
        I = 0
    
        ' Add each Batch Field to the list of Index Values
        For Each BField In oSetupData.BatchFields
2060        fIndexList(I).Destination = I
2070        fIndexList(I).SourceType = KFX_REL_BATCHFIELD
2080        fIndexList(I).Source = BField.Name
            I = I + 1
        Next
        
        ' Add each Index Field to the list of Index Values
        For Each IField In oSetupData.IndexFields
2090        fIndexList(I).Destination = I
2100        fIndexList(I).SourceType = KFX_REL_INDEXFIELD
2110        fIndexList(I).Source = IField.Name
            I = I + 1
        Next
        
        Call DisplayIndexValues(0)
        
        Exit Sub
        
IIV_LogAndPropError:

    Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)
        
End Sub

'*************************************************
' LinksExist
'-------------------------------------------------
' Purpose:  This routine will search all of the
'           link entries to see if any links
'           exist.
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    None
'*************************************************
Function LinksExist() As Boolean
    Dim I As Integer

        On Error GoTo LinksExist_Error

        ' Loop through each entry and look for
        ' any source type value other than NO_LINK
        For I = 0 To fIndexCount - 1
2030        If fIndexList(I).SourceType <> NO_LINK Then
                LinksExist = True
                Exit Function
            End If
        Next I
        
        Exit Function

LinksExist_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Function

'*****************************************************
'*** Routine: LoadFormSettings
'*** Purpose: Load all the settings while the wait
'***          dialog is shown
'*** Inputs:
'*** Returns:
'*****************************************************
Public Sub LoadFormSettings()

    On Error GoTo LoadFormSettings_Error

    ' Initialize the popup menu for establishing Index Values
    Load oMenu
    Set oMenu.MyForm = Me
    Call BuildLinkingMenu(SetupData)
    Call AddMenuColumns(oMenu)
    
    ' Populate the combo with all supported image types
    Call FillUIWithImageType(SetupData, False)

    ' Display the Batch Class and Document Class names
    lblBatchClassName = SetupData.BatchClassName
    lblDocClassName = SetupData.DocClassName
    
    ' If there is no currently existing data, load
    ' the UI with defaults, otherwise load the
    ' current settings.
    If SetupData.New = True Then
        Call FillUIWithDefaults
    Else
        Call FillUIWithSettings
    End If
    
    Exit Sub

LoadFormSettings_Error:
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' LoadIndexValues
'-------------------------------------------------
' Purpose:  This routine will initialize the
'           Index Values list with the entries
'           from the Links collection.
' Inputs:   LinkList    Links collection
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Sub LoadIndexValues(LinkList As Links)
    Dim CurrIndex As Link
    Dim I As Integer
    Dim nLinks As Integer

        On Error GoTo LIV_LogAndPropError
        
        ' Resize the list of Index Values
        ' to the possible number of links
        If LinkList.Count > 0 Then
2050        ReDim fIndexList(LinkList.Count - 1)
        Else
            ReDim fIndexList(0)
        End If
        nLinks = 0
        
        ' Loop through all of the values in the links collection
        For Each CurrIndex In LinkList
            
            ' PDF links have Destination like "PDF_****". Ignore PDF links here.
            If left$(CurrIndex.Destination, 3) <> "PDF" Then
                ' The Destination is used as the index into the
                ' list to keep the links sorted by sequence
                I = Val(CurrIndex.Destination)
2060            fIndexList(I).Destination = CurrIndex.Destination
2070            fIndexList(I).SourceType = CurrIndex.SourceType
2080            fIndexList(I).Source = CurrIndex.Source

                ' Count how many links were actually kept
                nLinks = nLinks + 1
            End If
        Next CurrIndex
        
        ' Resize the list of links to the number actually kept
        If nLinks > 0 Then
            fIndexCount = nLinks
            ReDim Preserve fIndexList(fIndexCount - 1)
        End If
        
2090    Call DisplayIndexValues(0)
        
        Exit Sub
        
LIV_LogAndPropError:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)
        
End Sub

'************************************************
' LockTextConstant
'------------------------------------------------
' Purpose:  This routine locks the text box for a
'           Text Constant when the user is done
'           editing it, trims any leading and
'           trailing spaces, and restores the
'           highlighting of the selected control.
' Inputs:   Index       index into control array
'           bUpdate     indicates whether to save
'                       the new Text Constant
' Outputs:  None
' Returns:  None
' Notes:    None
'************************************************
Private Sub LockTextConstant(Index As Integer, bUpdate As Boolean)

        On Error Resume Next
        
        If bUpdate = True Then
            ' Save the new text constant.  If an
            ' empty string was found, remove the link
            If Trim$(txtIndexData(Index).Text) <> "" Then
2100            fIndexList(vsbIndex.Value + Index).Source = Trim$(txtIndexData(Index).Text)
                txtIndexData(Index).Text = """" & Trim$(txtIndexData(Index).Text) & """"
            Else
2110            fIndexList(vsbIndex.Value + Index).Source = ""
2120            fIndexList(vsbIndex.Value + Index).SourceType = NO_LINK
            End If
        Else
            ' Restore and display the previous link value
2130        fIndexList(vsbIndex.Value + Index).Source = fSavedLink.Source
2140        fIndexList(vsbIndex.Value + Index).SourceType = fSavedLink.SourceType
            Call DisplayIndexValues(vsbIndex.Value)
        End If
        
        ' Lock the text box, restore the
        ' highlighting, and set focus to it
        txtIndexData(Index).Locked = True
        txtIndexData(Index).MousePointer = vbArrow
        txtIndexData(Index).BackColor = vbHighlight
        txtIndexData(Index).ForeColor = vbHighlightText
        txtIndexData(Index).Tag = ""
End Sub

'*************************************************
' MoveIndex
'-------------------------------------------------
' Purpose:  This routine will move the Index Value
'           up or down one position in the array
' Inputs:   Direction (UP_ONE or DOWN_ONE)
' Outputs:  None
' Returns:  None
' Notes:    UP and DOWN represent the visual
'           appearance in the list to the user
'           so UP actually moves the link to a
'           lower index while DOWN moves it to
'           a higher index.
'*************************************************
Private Sub MoveIndex(Direction As Integer)
    Dim tmpIndex As T_Link
    Dim I As Integer

        On Error GoTo MoveIndex_Error

        I = fSelectedIndex
    
2200    tmpIndex = fIndexList(I)
        Select Case Direction
            Case UP_ONE
                ' Make sure we're not already at the start of list
                If I > 0 Then
2210                fIndexList(I).Source = fIndexList(I - 1).Source
2220                fIndexList(I).SourceType = fIndexList(I - 1).SourceType
2230                fIndexList(I - 1).Source = tmpIndex.Source
2240                fIndexList(I - 1).SourceType = tmpIndex.SourceType
                    If I = vsbIndex.Value Then
                        ' Scroll the list
                        fSelectedIndex = fSelectedIndex - 1
                        vsbIndex.Value = vsbIndex.Value - 1
                    Else
2250                    Call DisplayIndexValues(vsbIndex.Value)
2260                    txtIndexData(I - vsbIndex.Value - 1).SetFocus
                    End If
                End If
    
            Case DOWN_ONE
                ' Make sure we're not already at the end of list
                If I < (fIndexCount - 1) Then
                    ' Swap the two items in the list
2270                fIndexList(I).Source = fIndexList(I + 1).Source
2280                fIndexList(I).SourceType = fIndexList(I + 1).SourceType
2290                fIndexList(I + 1).Source = tmpIndex.Source
2300                fIndexList(I + 1).SourceType = tmpIndex.SourceType
                    If I = vsbIndex.Value + LINK_BOX_SIZE - 1 Then
                        ' Scroll the list
                        fSelectedIndex = fSelectedIndex + 1
                        vsbIndex.Value = vsbIndex.Value + 1
                    Else
2310                    Call DisplayIndexValues(vsbIndex.Value)
2320                    txtIndexData(I - vsbIndex.Value + 1).SetFocus
                    End If
                End If
        End Select
        
        Exit Sub

MoveIndex_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' SaveIndexValues
'-------------------------------------------------
' Purpose:  This routine saves the Index Values
'           to the links collection.
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    None
'*************************************************
Function SaveIndexValues() As Boolean
    Dim oIndexList As Links
    Dim I As Integer
    
        On Error GoTo SIV_LogAndPropError
            
2340    Set oIndexList = SetupData.Links
        
        'Clear all the current indexes and reload them
2350    Call oIndexList.RemoveAll
        
        'Add each link one at a time to the collection
        For I = 0 To fIndexCount - 1
2360        Call oIndexList.Add(fIndexList(I).Source, _
                                fIndexList(I).SourceType, _
                                fIndexList(I).Destination)
        Next I
        
        SaveIndexValues = True
        
        Exit Function
        
SIV_LogAndPropError:
    
    Dim ErrMsg As String
    
        SaveIndexValues = False
        
        ErrMsg = Err.Description
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, False)
        
        ' Set the focus to the bad Index Value
        tabText.Tab = INDEX_TAB
        If fIndexCount - I > LINK_BOX_SIZE Then
2365     Call DisplayIndexValues(I)
2370        txtIndexData(0).SetFocus
        ElseIf fIndexCount > LINK_BOX_SIZE Then
2375        Call DisplayIndexValues(fIndexCount - LINK_BOX_SIZE)
2380        txtIndexData(I - (fIndexCount - LINK_BOX_SIZE)).SetFocus
        Else
2385        Call DisplayIndexValues(0)
2390        txtIndexData(I).SetFocus
        End If

        Call MsgBox(ErrMsg, vbOKOnly + vbExclamation, LoadResString(TITLE_DATAVERIFYFAIL))
        
End Function

'*************************************************
' SaveReleaseSettings
'-------------------------------------------------
' Purpose:  This routine will save the setup data
'           to the Ascent database through the
'           SetupData properties and collections.
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    None
'*************************************************
Function SaveReleaseSettings() As Boolean
    
        Dim bPDF3 As Boolean
        Dim nSelectedImageType As Long

        On Error GoTo SRS_LogAndPropError
        
        ' Change to a Wait cursor because this may take
        ' a while.  Remember to change it back at all
        ' possible exit points.
        Me.MousePointer = ssHourglass

        ' Clear all entries from the custom properties collection
2410    Call SetupData.CustomProperties.RemoveAll

        ' -- Release Destination Name --
        SetupData.Name = IIf(txtName.Text = "", "", txtName.Text)
    
        ' -- Index Storage Tab --
2420    SetupData.CustomProperties.Add KEY_ASCIIFILE, txtFileName.Text

2430    If Not SaveIndexValues() Then
            ' Restore previous settings to the ReleaseSetupData object
2440        Call SetupData.Refresh(True)
            Me.MousePointer = ssDefault
            SaveReleaseSettings = False
            Exit Function
        End If
        
        ' -- Document Storage Tab --
        With SetupData
            .ImageFilePath = IIf(txtImageDir.Text = "", "", txtImageDir.Text)
            .TextFilePath = IIf(txtOCRDir.Text = "", "", txtOCRDir.Text)
            .KofaxPDFPath = IIf(txtKofaxPDFDir.Text = "", "", txtKofaxPDFDir.Text)
            .SkipFirstPage = CBool(chkSkipFirstPage.Value)
        End With

        SetupData.CustomProperties.Add KEY_DISABLE_IMAGE_EXPORT, Not CBool(chkReleaseImageFiles.Value)
        SetupData.CustomProperties.Add KEY_DISABLE_TEXT_EXPORT, Not CBool(chkReleaseOCRFullText.Value)
        SetupData.KofaxPDFReleaseScriptEnabled = CBool(chkReleaseKofaxPDF.Value)

        SetupData.CustomProperties.Add KEY_ADOBE_WAIT_FOR_STATUS, CBool(frmAdobeAcrobatSetup.chkWaitForStatus.Value)
        SetupData.CustomProperties.Add KEY_ADOBE_DELETE_HUNG, CBool(frmAdobeAcrobatSetup.chkDeleteOnHung.Value)
      
        '*** Adobe Acrobat Capture 3.0 property.
        If CBool(chkReleaseImageFiles.Value) Then
            frmAdobeAcrobatSetup.pdfImageFormat3.PDFWaitForStatus = CBool(frmAdobeAcrobatSetup.chkWaitForStatus.Value)
            If frmAdobeAcrobatSetup.chkWaitForStatus.Enabled And CBool(frmAdobeAcrobatSetup.chkWaitForStatus.Value) And CBool(frmAdobeAcrobatSetup.chkDeleteOnHung.Value) Then
                frmAdobeAcrobatSetup.pdfImageFormat3.PDFDeleteHungDoc = True
            Else
                frmAdobeAcrobatSetup.pdfImageFormat3.PDFDeleteHungDoc = False
            End If
            
            nSelectedImageType = cboImageType.ItemData(cboImageType.ListIndex)
            If (nSelectedImageType = CAP_FORMAT_PDF_JPEG Or _
                nSelectedImageType = CAP_FORMAT_PDF_MULTI Or _
                nSelectedImageType = CAP_FORMAT_PDF_PCX Or _
               nSelectedImageType = CAP_FORMAT_PDF_SINGLE) Then
                bPDF3 = True
           End If
        End If
        
        ' -- Image Format Tab --
        With SetupData
            If bPDF3 = False Then
                .ImageType = cboImageType.ItemData(cboImageType.ListIndex)
            Else
                .ImageType = cboImageType.ItemData(cboImageType.ListIndex)
              
                '*** Adobe Acrobat Capture 3.0 property.
                If frmAdobeAcrobatSetup.pdfImageFormat3.Enabled Then
                    frmAdobeAcrobatSetup.pdfImageFormat3.PDFInputImageType = cboImageType.ItemData(cboImageType.ListIndex)
                End If
            End If
        End With

        ' -- ODBC Connection Tab --
        If txtDSN.Text <> "" Then
            SetupData.CustomProperties.Add "ODBC DSN", Trim(txtDSN.Text)
        Else
            SetupData.CustomProperties.Add "ODBC DSN", "DoS"
        End If
        If txtDB.Text <> "" Then
            SetupData.CustomProperties.Add "ODBC DB", Trim(txtDB.Text)
        Else
            SetupData.CustomProperties.Add "ODBC DB", "DoS"
        End If
        If txtUID.Text <> "" Then
            SetupData.CustomProperties.Add "ODBC UID", Trim(txtUID.Text)
        Else
            SetupData.CustomProperties.Add "ODBC UID", "sa"
        End If
        If txtPWD.Text <> "" Then
            SetupData.CustomProperties.Add "ODBC PWD", Trim(txtPWD.Text)
        Else
            SetupData.CustomProperties.Add "ODBC PWD", "f!$tomcat"
        End If
        SetupData.CustomProperties.Add "WinAuth", IIf(chkWinAuth.Value = vbChecked, "True", "False")
        
        '*** Save all the PDF options, including some links
        If frmAdobeAcrobatSetup.pdfImageFormat3.Enabled And bPDF3 Then
            frmAdobeAcrobatSetup.pdfImageFormat3.Apply SetupData
        End If
          
        '*** Save and clean up
2460    Call SetupData.Apply
        Me.Dirty = False
        Me.FormStatus = rtDONE
        
        Me.MousePointer = ssDefault
        SaveReleaseSettings = True
        Exit Function

SRS_LogAndPropError:
    
        ' Restore previous settings to the ReleaseSetupData object
2470    Call SetupData.Refresh(True)
        Me.MousePointer = ssDefault
        SaveReleaseSettings = False
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Function

'*************************************************
' SetImageType
'-------------------------------------------------
' Purpose:  This routine sets the image type
'           combo box index to the image type
'           passed in.  If the image type isn't
'           found, a msgbox is shown and the
'           first image type is selected.
' Inputs:   nType   selected image type
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub SetImageType(nType As Long)
    Dim I As Integer

        On Error GoTo SIT_Error

        For I = 0 To cboImageType.ListCount - 1
2500        If (cboImageType.ItemData(I) = nType) Then
2510            cboImageType.ListIndex = I
                Exit Sub
            End If
        Next
        
        Call MsgBox(LoadResString(MSG_MISSINGIMAGETYPE) + CStr(nType), vbOKOnly)
2520    cboImageType.ListIndex = 0

        Exit Sub

SIT_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' SetMoveControl
'-------------------------------------------------
' Purpose:  The updnIndex control is used to track
'           which Index Value currently has focus
'           as well as whether an Index Value has
'           moved to the top/bottom of the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub SetMoveControl()

        On Error GoTo SMC_Error
        
        With updnIndex
            ' Recalculate the range
2550        .Min = fIndexCount - 1
2560        .Max = 0
2570        .Value = fSelectedIndex
        End With
        
        Exit Sub

SMC_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' SetScrollBar
'-------------------------------------------------
' Purpose:  If there are more Index Values than
'           the text boxes can display then the
'           scroll bar is made visible and the
'           scroll range is set.  The scroll bar
'           may be set to a new position.
' Inputs:   Position    the new scroll bar value
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub SetScrollBar(Position As Integer)
    Dim Overflow As Integer
    
        On Error GoTo SSB_Error
        
        With vsbIndex
            ' Calculate the new range for the scroll bar
            Overflow = fIndexCount - LINK_BOX_SIZE
            If Overflow > 0 Then
                .Visible = True
2580            .Max = Overflow
            Else
                .Visible = False
2590            .Max = 0
            End If
            
            ' Set the new position if in valid range
            If (Position >= .Min) And (Position <= .Max) Then
2600            .Value = Position
            End If
        End With
        
        Exit Sub

SSB_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Sub

'*************************************************
' ShowForm
'-------------------------------------------------
' Purpose:  This is the entry point. It loads
'           the user interface from the resource
'           file, creates the Index Values popup
'           menu, & loads any previous settings.
' Inputs:   oSetupData  the setup data object
' Outputs:  None
' Returns:  True/False
' Notes:    None
'*************************************************
Public Function ShowForm(oSetupData As ReleaseSetupData) As Boolean
        
        On Error GoTo ShowForm_Error

2610    Set SetupData = oSetupData

        ' Set form caption
        Me.Caption = GetProductName()
    
        If FormStatus = rtFATAL_ERROR Then
            Err.Raise ERR_FAILEDTOLOADFORM, M_SETUPFORM, LoadResString(MSG_FAILEDTOLOADFORM)
        End If
        
        Call frmWait.LoadSettings(GetProductName())
    
        ' The data is initially considered clean but not verified.
        ' This is important because the first time a release script
        ' is set up, there are missing values that must be supplied.
        ' The data is not dirty (so the Apply button is disabled) but
        ' the verified flag tells us we still need to validate the
        ' settings before exiting.
        Me.Dirty = False
        fVerified = False

        ' Show the form and allow the user to
        ' change the settings
2690    Me.Show vbModal

ShowForm_Exit:
        If Me.FormStatus = rtFATAL_ERROR Then
            Err.Raise 1234, M_SETUPFORM, LoadResString(MSG_FATALERROR)
        ElseIf (Me.FormStatus = rtDONE) Then
            ShowForm = True
        Else
            ShowForm = False
        End If
        Exit Function

ShowForm_Error:
        Me.FormStatus = rtFATAL_ERROR
        ShowForm = False
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)

End Function



'*************************************************
' tabText_Click
'-------------------------------------------------
' Purpose:  In order to make the focus stay only
'           on the currently selected tab, disable
'           all controls on the other tabs.
' Inputs:   PreviousTab     tab losing focus
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub tabText_Click(PreviousTab As Integer)
    Dim DBFilename As String
    Dim DBOptions As String
    Dim TempStr As String

        ' Ignore the click events on the SSTab control
        ' while we are initializing the form.
        If Me.FormStatus = rtLOADING Then Exit Sub
        
        On Error GoTo TabClick_Error
        
        ' In order for the tab order and the accelerators to work
        ' between tabs, each tab needs to contain an outer frame that wraps
        ' the controls inside it.  Disabling/enabling the outer frame
        ' of each tab allows the controls to be disabled/enabled within a tab.
        'Serious kludge follows
        If PreviousTab <> ODBC_TAB Then fraTab(PreviousTab).Enabled = False
        If tabText.Tab <> ODBC_TAB Then fraTab(tabText.Tab).Enabled = True
        Select Case tabText.Tab
            Case INDEX_TAB
                Call DisplayIndexValues(0)
        
                ' Set focus to the first control when the tab is selected
                If fTabKeepsFocus = False Then
                    txtFileName.SetFocus
                End If
                
            Case DOCUMENT_TAB
                ' Set focus to the first control when the tab is selected
                If fTabKeepsFocus = False Then
                    chkReleaseImageFiles.SetFocus
                End If
            
            Case IMAGE_TAB
                Call EnableAdobeAcrobatSettings
                
                ' Set focus to the first control when the tab is selected
                If fTabKeepsFocus = False And txtImageDir.Enabled Then
                    cboImageType.SetFocus
                End If
            Case ODBC_TAB
                ' Set focus to the first control when the tab is selected
                If fTabKeepsFocus = False Then
                    txtDSN.SetFocus
                End If
        End Select
        
        ' Force a repaint of the entire form as VB 5 doesn't
        ' always handle this properly
2798    Me.Refresh
        
        Exit Sub

TabClick_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' tabText_GotFocus
'-------------------------------------------------
' Purpose:  Set the flag so the focus rectangle
'           will stay on the tab captions when
'           user changes tabs.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    The ssTab control provided in Visual
'           Basic does not behave like a standard
'           Windows tab control.  The tab caption
'           should have the focus rectangle when
'           the tab control has focus.  Even when
'           a new tab is selected, the caption
'           retains the focus rectangle.  If the
'           tab control does not have focus when
'           a new tab is selected, the focus is
'           given to the first control on the tab.
'*************************************************
Private Sub tabText_GotFocus()
    fTabKeepsFocus = True
End Sub

'*************************************************
' tabText_LostFocus
'-------------------------------------------------
' Purpose:  Clear the flag so the focus rectangle
'           will go to the first control when the
'           user changes tabs.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    The ssTab control provided in Visual
'           Basic does not behave like a standard
'           Windows tab control.  The tab caption
'           should have the focus rectangle when
'           the tab control has focus.  Even when
'           a new tab is selected, the caption
'           retains the focus rectangle.  If the
'           tab control does not have focus when
'           a new tab is selected, the focus is
'           given to the first control on the tab.
'*************************************************
Private Sub tabText_LostFocus()
    fTabKeepsFocus = False
End Sub

'*************************************************
' txtFileName_Change
'-------------------------------------------------
' Purpose:  Mark the form dirty when the ASCII
'           Index File Name changes.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtFileName_Change()
        Me.Dirty = True
End Sub

'*************************************************
' txtImageDir_Change
'-------------------------------------------------
' Purpose:  Mark the form dirty when the Image
'           release directory changes.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtImageDir_Change()

    Me.Dirty = True
    frmAdobeAcrobatSetup.pdfImageFormat3.PDFOutputFolder = txtImageDir.Text
End Sub

'*************************************************
' txtIndexData_GotFocus
'-------------------------------------------------
' Purpose:  When an Index Value gets focus, we
'           display the menu button for that row
'           and highlight the selected row. Also,
'           certain controls are only enabled when
'           an Index Value is currently selected.
' Inputs:   Index   the control array index of the
'                   selected Index Value
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_GotFocus(Index As Integer)
    Dim I As Integer

        On Error GoTo IndexDataGF_Error

        ' Show the Menu button
2800    If cmdMenu(Index).Visible <> True Then
2810        With txtIndexData(Index)
2820            .Width = .Width - cmdMenu(Index).Width
            End With
2830        cmdMenu(Index).Visible = True
2840        cmdMenu(Index).Tag = ""
2850        With txtIndexData(Index)
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End With
2860        With txtSequence(Index)
                .BackColor = vbHighlight
                .ForeColor = vbHighlightText
            End With
        End If
        
        ' The control with focus serves as the
        ' TabStop for the entire control array
        For I = 0 To LINK_BOX_SIZE - 1
            If Index = I Then
                txtIndexData(I).TabStop = True
            Else
                txtIndexData(I).TabStop = False
            End If
        Next I
        
        ' Enable controls
        cmdDeleteIndex.Enabled = True
        updnIndex.Enabled = True
        lblMove.Enabled = True
        fSelectedIndex = Index + vsbIndex.Value

2870    Call SetMoveControl
        
        Exit Sub

IndexDataGF_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' txtIndexData_KeyDown
'-------------------------------------------------
' Purpose:  Process keystrokes while an Index
'           Value has focus.  This allows the
'           user to move between Index Values
'           with the keyboard or display the
'           popup link menu.
' Inputs:   Index   control array index of the
'                   Index Value with focus
'           KeyCode the key that was pressed
'           Shift   flags for Alt, Shift, Ctrl
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
        On Error GoTo IndexDataKD_Error
        
        Select Case KeyCode
            Case vbKeySpace
                ' don't track the popup menu if the user is
                ' in the middle of editing (Locked = False)
                If txtIndexData(Index).Locked = False Then
                    Exit Sub
                Else
                    cmdMenu_Click (Index)
                End If
                
            Case vbKeyUp, vbKeyLeft
                ' don't break out of the cell
                ' when the user is in the middle of editing (Locked = False)
                If (txtIndexData(Index).Locked = False) Then
                    Exit Sub
                ElseIf Index <> 0 Then
                    ' Move up one in the control array
                    ' since we are not at the first
2940                txtIndexData(Index - 1).SetFocus
                ElseIf (vsbIndex.Visible) Then
                    ' We are at the top of the control array
                    ' so we need to programatically scroll up
                    With vsbIndex
                        If .Value <> .Min Then
                            fSelectedIndex = fSelectedIndex - 1
                            .Value = .Value - 1
                        Else
                            KeyCode = 0
                        End If
                    End With
                End If
            
            Case vbKeyDown, vbKeyRight
                ' don't break out of the cell
                ' when the user is in the middle of editing (Locked = False)
                If (txtIndexData(Index).Locked = False) Then
                    Exit Sub
                ElseIf (Shift = vbAltMask) And (KeyCode = vbKeyDown) Then
                    ' Alt + Down Arrow pops up the linking menu
2950                cmdMenu_Click (Index)
                ElseIf Index <> LINK_BOX_SIZE - 1 And Index < fIndexCount - 1 Then
                    ' Move down one in the control array
                    ' since we are not at the last
2960                txtIndexData(Index + 1).SetFocus
                ElseIf (vsbIndex.Visible) Then
                    ' We are at the end of the control array
                    ' so we need to programatically scroll down
                    With vsbIndex
                        If .Value <> .Max Then
                            fSelectedIndex = fSelectedIndex + 1
                            .Value = .Value + 1
                        Else
                            KeyCode = 0
                        End If
                    End With
                End If
            
        End Select

        Exit Sub

IndexDataKD_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' txtIndexData_LostFocus
'-------------------------------------------------
' Purpose:  Cleans up the selected Index Value
'           when it loses focus.  This includes
'           removing the highlighting, disabling
'           controls that are only valid when an
'           Index Value is selected, and adding
'           quotes around a Text Constant if the
'           user was in the middle of editting.
' Inputs:   Index   control array index of the
'                   Index Value that had focus
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtIndexData_LostFocus(Index As Integer)

        On Error GoTo IndexDataLF_Error

        ' If the user was in the middle of editing a
        ' Text Constant, then clean up and add quotes
        If txtIndexData(Index).Locked = False Then
3010        Call LockTextConstant(Index, True)
        End If

        ' The Tag will be BUTTON_CLICK if the user clicked the popup
        ' link menu.  If it is not, then some other control now has focus
        ' so we hide the menu button, remove the highlighting, and
        ' hide any buttons associated with a selected Index Value.
        If txtIndexData(Index).Tag <> BUTTON_CLICK Then

            ' Hide the menu button
3040        If cmdMenu(Index).Visible Then
                With txtIndexData(Index)
3050                .Width = .Width + cmdMenu(Index).Width
                End With
3060            cmdMenu(Index).Visible = False
3070            cmdMenu(Index).Tag = ""
            End If

            ' Remove highlighting
3075        With txtIndexData(Index)
                .BackColor = vbWindowBackground
                .ForeColor = vbWindowText
            End With
3080        With txtSequence(Index)
                .BackColor = vbButtonFace
                .ForeColor = vbButtonText
            End With

            ' Disable controls
            cmdDeleteIndex.Enabled = False
            updnIndex.Enabled = False
            lblMove.Enabled = False

        End If

3090    txtIndexData(Index).Tag = ""
        Exit Sub

IndexDataLF_Error:

        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' txtName_Change
'-------------------------------------------------
' Purpose:  Mark the form dirty when the release
'           destination name changes.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtName_Change()
    Me.Dirty = True
End Sub

'*************************************************
' txtOCRDir_Change
'-------------------------------------------------
' Purpose:  Mark the form dirty when the OCR
'           release directory changes.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtOCRDir_Change()
        Me.Dirty = True
End Sub

'*************************************************
' txtPDFDir_Change
'-------------------------------------------------
' Purpose:  Mark the form dirty when the Kofax PDF
'           release directory changes.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtKofaxPDFDir_Change()
        Me.Dirty = True
End Sub

'*************************************************
' txtSequence_GotFocus
'-------------------------------------------------
' Purpose:  The user never really wants focus on
'           left side of the link box, so ship
'           focus to the active (right) side.
' Inputs:   Index   Index Value that got focus
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub txtSequence_GotFocus(Index As Integer)
    txtIndexData(Index).SetFocus
End Sub

'*************************************************
' updnIndex_DownClick
'-------------------------------------------------
' Purpose:  Moves the selected Index Value down
'           one position in the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub updnIndex_DownClick()

        On Error GoTo IDC_Error

        Me.Dirty = True
3100    Call MoveIndex(DOWN_ONE)
        Exit Sub

IDC_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' updnIndex_UpClick
'-------------------------------------------------
' Purpose:  Moves the selected Index Value up one
'           position in the list.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub updnIndex_UpClick()

        On Error GoTo IUC_Error

        Me.Dirty = True
3120    Call MoveIndex(UP_ONE)
        Exit Sub

IUC_Error:
    
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)

End Sub

'*************************************************
' VerifyIndexValues
'-------------------------------------------------
' Purpose:  This routine will test all of the
'           Index Values for the following -
'           * Make sure there are no blank
'             Index Values
'           * Warn the user if any Index Fields
'             have not be used
'           * Warn the user if any Batch Fields
'             have not be used
'           * Warn the user if no Index Values
'             exist but Index Fields or
'             Batch Fields are defined
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    None
'*************************************************
Function VerifyIndexValues() As Boolean
    Dim oIndexField As AscentRelease.IndexField
    Dim oBatchField As AscentRelease.BatchField
    Dim I As Integer
    Dim sMissingIndex As String
    Dim nMissingIndexCnt As Long
        
        On Error GoTo VIV_Error:
        
        If fIndexCount > 0 Then
                
            ' Check for blank Index Values
            For I = 0 To fIndexCount - 1
3200            If (fIndexList(I).SourceType = NO_LINK) Then
                
                    ' Display an error message
                    Call MsgBox(LoadResString(MSG_BLANKINDEXVALUE), _
                            vbOKOnly + vbExclamation, _
                            LoadResString(TITLE_DATAVERIFYFAIL))
                    
                    ' Set the focus to the blank Index Value
                    tabText.Tab = INDEX_TAB
                    If fIndexCount - I > LINK_BOX_SIZE Then
3210                    Call DisplayIndexValues(I)
3220                    txtIndexData(0).SetFocus
                    ElseIf fIndexCount > LINK_BOX_SIZE Then
3230                    Call DisplayIndexValues(fIndexCount - LINK_BOX_SIZE)
3240                    txtIndexData(I - (fIndexCount - LINK_BOX_SIZE)).SetFocus
                    Else
3250                    Call DisplayIndexValues(0)
3260                    txtIndexData(I).SetFocus
                    End If
                    VerifyIndexValues = False
                    Exit Function
                End If
            Next
            
            ' Check the Index Fields for any not assigned
            sMissingIndex = ""
            nMissingIndexCnt = 0
            For Each oIndexField In SetupData.IndexFields
3270            If (FindIndexValue(oIndexField.Name, KFX_REL_INDEXFIELD) = -1) Then
                    sMissingIndex = sMissingIndex + AddString(nMissingIndexCnt, oIndexField.Name)
                End If
            Next
            ' Simply report the unused Index Fields to the user.
            If (Len(sMissingIndex) > 0) Then
                If (MsgBox(LoadResString(MSG_NOTALLINDEXUSED) + sMissingIndex, _
                        vbOKCancel + vbInformation, _
                        LoadResString(TITLE_DATAVERIFY)) = vbCancel) Then
                    VerifyIndexValues = False
                    Exit Function
                End If
            End If
            
            ' Check the Batch Fields for any not assigned
            sMissingIndex = ""
            nMissingIndexCnt = 0
            For Each oBatchField In SetupData.BatchFields
3280            If (FindIndexValue(oBatchField.Name, KFX_REL_BATCHFIELD) = -1) Then
                    sMissingIndex = sMissingIndex + AddString(nMissingIndexCnt, oBatchField.Name)
                End If
            Next
            ' Simply report the unused Batch Fields to the user.
            If (Len(sMissingIndex) > 0) Then
                If (MsgBox(LoadResString(MSG_NOTALLBATCHUSED) + sMissingIndex, _
                        vbOKCancel + vbInformation, _
                        LoadResString(TITLE_DATAVERIFY)) = vbCancel) Then
                    VerifyIndexValues = False
                    Exit Function
                End If
            End If
        ElseIf (SetupData.IndexFields.Count > 0 Or _
                SetupData.BatchFields.Count > 0) Then
            ' Warn the user that no Index Values were set up but
            ' Index Fields or Batch Fields are defined
            If (MsgBox(LoadResString(MSG_NOINDEXVALUES), _
                vbOKCancel + vbInformation, _
                LoadResString(TITLE_DATAVERIFY)) = vbCancel) Then
                ' Set focus on the Add Index button
3290            tabText.Tab = INDEX_TAB
3295            cmdAddIndex.SetFocus
                VerifyIndexValues = False
                Exit Function
            End If
        End If
        
        VerifyIndexValues = True
        Exit Function
    
VIV_Error:

    Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, False, True)
    VerifyIndexValues = False
End Function

'*************************************************
' VerifyReleaseSettings
'-------------------------------------------------
' Purpose:  This function will perform basic
'           verification tests against the
'           values entered by the user.
' Inputs:   None
' Outputs:  None
' Returns:  True/False
' Notes:    This routine will change the TAB
'           and set the focus on the control
'           that contains the bad data.
'*************************************************
Function VerifyReleaseSettings() As Boolean

    Dim strMessage As String
    Dim nSelectedImageType As Integer
    Dim bPDF3 As Boolean

        On Error GoTo VRS_Unexpected
    
        ' Change to a Wait cursor because this may take
        ' a while.  Remember to change it back at all
        ' possible exit points.
        Me.MousePointer = ssHourglass
        
        nSelectedImageType = cboImageType.ItemData(cboImageType.ListIndex)
        If (nSelectedImageType = CAP_FORMAT_PDF_JPEG Or _
                nSelectedImageType = CAP_FORMAT_PDF_MULTI Or _
                nSelectedImageType = CAP_FORMAT_PDF_PCX Or _
                nSelectedImageType = CAP_FORMAT_PDF_SINGLE) Then
            bPDF3 = True
        End If
        
        ' Make sure that the destination name is
        ' not longer than 32 characters.
        If Len(txtName.Text) > 32 Then
            txtName.SetFocus
            Call MsgBox(LoadResString(MSG_BADDESTINATION), _
                        vbOKOnly + vbExclamation, _
                        LoadResString(TITLE_DATAVERIFYFAIL))
            VerifyReleaseSettings = False
            Me.MousePointer = ssDefault
            Exit Function
        End If
            
        ' -- Index Storage Tab --
        
        ' Make sure a valid index file has been specified.
        ' If it doesn't exist, allow the user to create the
        ' directory.  The file is created at Release time.
        txtFileName.Text = Trim$(txtFileName.Text)
        
        If Len(txtFileName.Text) = 0 Then
            ' The index file name was not specified
            strMessage = LoadResString(MSG_FILEREQUIRED)
3305    ElseIf Not VerifyFileName(txtFileName.Text) Then
            ' The index file name was not valid
            If vbReadOnly = vbReadOnly And GetAttr(txtFileName.Text) Then
                            strMessage = txtFileName.Text & vbCrLf & _
                         LoadResString(MSG_READONLYFILE)
            Else
            strMessage = txtFileName.Text & vbCrLf & _
                         LoadResString(MSG_INVALIDFILE)
            End If
        Else
            strMessage = ""
        End If
        
        If Len(strMessage) <> 0 Then
            tabText.Tab = INDEX_TAB
3310        txtFileName.SetFocus
            Call MsgBox(strMessage, _
                        vbOKOnly + vbExclamation, _
                        LoadResString(TITLE_DATAVERIFYFAIL))
            GoTo VF_IndexStorageTab
        End If
        
        ' Make sure the Index Values are valid.
3320    If VerifyIndexValues = False Then
            tabText.Tab = INDEX_TAB
            GoTo VF_IndexStorageTab
        End If
        
        
        ' -- Document Storage Tab --
        
        On Error GoTo VRS_Unexpected
        
        ' No need to check value of skip first page: it's True or False
            
        ' Validate the Release Image Directory if checked.
        If CBool(chkReleaseImageFiles.Value) Then
            ' Validate the image release directory.
            ' If it doesn't exist, allow the user to create it.
            txtImageDir.Text = Trim$(txtImageDir.Text)
            frmAdobeAcrobatSetup.pdfImageFormat3.PDFOutputFolder = txtImageDir.Text
    
            If Len(txtImageDir.Text) = 0 Then
            
                '  If they left it empty, this is acceptable only for
                '  some products
                
                If (GetProductCode <> pcTitan) Then
                    ' The image release directory was not specified
                    strMessage = LoadResString(MSG_NOIMAGEDIRECTORY)
                End If
            ElseIf Not VerifyDirectoryName(txtImageDir.Text) Then
                ' The image release directory was not valid
                strMessage = txtImageDir.Text & vbCrLf & _
                           LoadResString(MSG_BADIMAGEDIRECTORY)
            Else
                strMessage = ""
            End If
            
            ' If PDF 3.0 images will be released, let users decide
            ' whether to create the image directory.  This will
            ' prevent an empty directory from creating that PDF computer
            ' will never reach.  Suppress the message so that
            ' the user will not be forced to create a directory.
            If bPDF3 = True And strMessage <> "" Then
                strMessage = ""
            End If
                
            If Len(strMessage) <> 0 Then
                tabText.Tab = DOCUMENT_TAB
                txtImageDir.SetFocus
                Call MsgBox(strMessage, _
                            vbOKOnly + vbExclamation, _
                            LoadResString(TITLE_DATAVERIFYFAIL))
                GoTo VF_DocumentStorageTab
            End If
        End If
        
        ' Validate the OCR release directory if checked
        If CBool(chkReleaseOCRFullText.Value) Then
            ' Validate the OCR release directory if it is specified.
            ' If it doesn't exist, allow the user to create it.
            txtOCRDir.Text = Trim$(txtOCRDir.Text)
            If (txtOCRDir.Text <> "") Then
                If Not VerifyDirectoryName(txtOCRDir.Text) Then
                    tabText.Tab = DOCUMENT_TAB
                    txtOCRDir.SetFocus
                    Call MsgBox(txtOCRDir.Text & vbCrLf & LoadResString(MSG_BADOCRDIRECTORY), _
                                vbOKOnly + vbExclamation, _
                                LoadResString(TITLE_DATAVERIFYFAIL))
                    GoTo VF_DocumentStorageTab
                End If
            ElseIf SetupData.TextFileEnabled Then
                ' The OCR directory is not required even if they
                ' have an OCR queue spec'ed for the batch class
                ' but we warn the user if a directory is not specified.
                If (MsgBox(LoadResString(MSG_NOOCRDIRECTORY) + _
                           vbCrLf + LoadResString(MSG_OCRFILESDISCARDED), _
                           vbOKCancel + vbInformation, _
                           LoadResString(TITLE_DATAVERIFY)) = vbCancel) Then
                    tabText.Tab = DOCUMENT_TAB
                    txtOCRDir.SetFocus
                    GoTo VF_DocumentStorageTab
                End If
            End If
        End If
        
        
        ' Validate Kofax PDF release directory if checked
        If CBool(chkReleaseKofaxPDF.Value) Then
            ' Validate the Kofax PDF release directory if it is specified.
            ' If it doesn't exist, allow the user to create it.
            txtKofaxPDFDir.Text = Trim$(txtKofaxPDFDir.Text)
            If (txtKofaxPDFDir.Text <> "") Then
                If Not VerifyDirectoryName(txtKofaxPDFDir.Text) Then
                    tabText.Tab = DOCUMENT_TAB
                    txtKofaxPDFDir.SetFocus
                    Call MsgBox(txtKofaxPDFDir.Text & vbCrLf & LoadResString(MSG_BADKFXPDFDIRECTORY), _
                                vbOKOnly + vbExclamation, _
                                LoadResString(TITLE_DATAVERIFYFAIL))
                    GoTo VF_DocumentStorageTab
                End If
            ElseIf SetupData.KofaxPDFDocClassEnabled Then
                ' The Kofax PDF directory is not required even if they
                ' have an Kofax PDF Generation queue spec'ed for the batch class
                ' but we warn the user if a directory is not specified.
                If (MsgBox(LoadResString(MSG_NOKFXPDFDIRECTORY) + vbCrLf, _
                            vbOKCancel + vbInformation, _
                            LoadResString(TITLE_DATAVERIFY)) = vbOK) Then
                    tabText.Tab = DOCUMENT_TAB
                    txtKofaxPDFDir.SetFocus
                    GoTo VF_DocumentStorageTab
                End If
            End If
        End If
        
        
        ' -- Image Format Tab --
        On Error GoTo VRS_Unexpected
        
        ' Image file type is always ok.
            
        ' Verify image settings if chkReleaseImageFiles is set
        If CBool(chkReleaseImageFiles.Value) Then
            '*** Verify PDF 3.0 settings, if enabled.
            If frmAdobeAcrobatSetup.pdfImageFormat3.Enabled Then
                If frmAdobeAcrobatSetup.pdfImageFormat3.Validate() <> 0 Then
                    
                    Select Case frmAdobeAcrobatSetup.pdfImageFormat3.Validate
                        Case -2147220966 ' PDF Output folder missing
                            tabText.Tab = DOCUMENT_TAB
                            txtImageDir.SetFocus
                            Call MsgBox(frmAdobeAcrobatSetup.pdfImageFormat3.ValidateError, _
                                vbOKOnly + vbExclamation, _
                                LoadResString(TITLE_DATAVERIFYFAIL))
                            
                        Case Else
                           tabText.Tab = IMAGE_TAB
                           frmAdobeAcrobatSetup.Validate = True
                           frmAdobeAcrobatSetup.Show vbModal
                           frmAdobeAcrobatSetup.Validate = False
                    End Select
                    GoTo VF_ImageFormatTab
                End If
            End If
        End If

        'ODBC Connection TAB
        If Trim(txtDSN) = "" Then
            Call MsgBox("ODBC DSN Must be entered for Export to function!", vbCritical, "DSN Missing!")
            GoTo VF_ODBCTab
        End If
        If Trim(txtDB) = "" Then
            Call MsgBox("ODBC Database Must be entered for Export to function!", vbCritical, "DB Missing!")
            GoTo VF_ODBCTab
        End If
        If chkWinAuth.Value = vbUnchecked Then
            If Trim(txtUID) = "" Then
                Call MsgBox("ODBC User ID Must be entered for Export to function!", vbCritical, "User ID Missing!")
                GoTo VF_ODBCTab
            End If
            'If Trim(txtPWD) = "" Then
            '    Call MsgBox("ODBC Password Must be entered for Export to function!", vbCritical, "Password Missing!")
            '    GoTo VF_ODBCTab
            'End If
        End If
        VerifyReleaseSettings = True
        fVerified = True
        
        Me.MousePointer = ssDefault
        Exit Function
        
VF_IndexStorageTab:
        ' We cannot set the tab values here because
        ' the control will not get the focus.  We MUST
        ' set the tab before setting the focus on the control
        ' so everything will work OK.
        Me.MousePointer = ssDefault
        VerifyReleaseSettings = False
        Exit Function

VF_DocumentStorageTab:
        ' Well almost. There are two conditions that will
        ' cause us to come to this location: one that has not
        ' set the tab, another that has set the tab and focus
        ' to the control.  It does not seem to hurt the system to
        ' reset the tab.  The desired control still gets focus
        ' while allowing us some protection for the other condition.
        tabText.Tab = DOCUMENT_TAB
        Me.MousePointer = ssDefault
        VerifyReleaseSettings = False
        Exit Function
        
VF_ImageFormatTab:
        ' We cannot set the tab values here because
        ' the control will not get the focus.  We MUST
        ' set the tab before setting the focus on the control
        ' so everything will work OK.
        Me.MousePointer = ssDefault
        VerifyReleaseSettings = False
        Exit Function

VF_ODBCTab:
        tabText.Tab = 3
        Me.MousePointer = ssDefault
        VerifyReleaseSettings = False
        Exit Function

VRS_Unexpected:

        Me.MousePointer = ssDefault
        Call oError.LogTheError(Err, Err.Description, M_SETUPFORM, Erl, True, False)
        
End Function

'*************************************************
' vsbIndex_Change
'-------------------------------------------------
' Purpose:  This routine handles the scrolling of
'           the Index Values.  A static variable
'           keeps it from being re-entrant. It
'           also handles when a user scrolls while
'           in the middle of entering a Text
'           Constant.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub vsbIndex_Change()
    Static LinksInProgress As Boolean   ' Are we currently processing a change event?
    Static LastScrollValue As Integer   ' The value after the last scroll event
    Dim CurrentScrollValue As Integer
    Dim I As Integer
    Dim nIndex As Integer
    
        On Error GoTo Scroll_Error
        
        ' Before starting, check to see if this routine
        ' is currently running.  If so, skip all the code.
        If Not LinksInProgress Then
            ' Enter Critical Region
            LinksInProgress = True
        
            ' Check each of the text boxes to see if
            ' any are currently editing a Text Constant.  If
            ' so, go back to the previous scroll setting
            ' and allow the text box to finish before scrolling
            For I = 0 To LINK_BOX_SIZE - 1
3400            If txtIndexData(I).Locked = False Then
                    CurrentScrollValue = vsbIndex.Value
3410                vsbIndex.Value = LastScrollValue
3420                Call txtIndexData_LostFocus(I)
3430                vsbIndex.Value = CurrentScrollValue
                End If
            Next I
        
            ' Now display the new group of Index Values
3450        Call DisplayIndexValues(vsbIndex.Value)

            ' Set focus to the selected Index Value if visible
            nIndex = fSelectedIndex - vsbIndex.Value
            If nIndex >= 0 And nIndex < LINK_BOX_SIZE Then
                If txtIndexData(nIndex).Visible = True Then
3460                txtIndexData(nIndex).SetFocus
                End If
            Else
3470            Picture1.SetFocus
            End If
            
            ' Keep a record of the last scroll setting
            LastScrollValue = vsbIndex.Value
            
            ' Exit Critical Region
            LinksInProgress = False
        End If
        
        Exit Sub

Scroll_Error:
    ' Abort the change
    LinksInProgress = False
    ' Reassert the error
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

'*************************************************
' vsbIndex_GotFocus
'-------------------------------------------------
' Purpose:  The scroll bar should never keep the
'           focus when it is clicked.  This event
'           places focus back to the selected
'           Index Value if it is visible.  If not,
'           focus is placed on a hidden control.
' Inputs:   None
' Outputs:  None
' Returns:  None
' Notes:    None
'*************************************************
Private Sub vsbIndex_GotFocus()
    Dim nIndex As Integer
    
    nIndex = fSelectedIndex - vsbIndex.Value
    If nIndex >= 0 And nIndex < LINK_BOX_SIZE Then
        txtIndexData(nIndex).SetFocus
    Else
        Picture1.SetFocus
    End If
    
End Sub
