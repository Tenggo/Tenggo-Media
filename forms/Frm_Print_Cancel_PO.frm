VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frm_Print_Cancel_PO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation Order Print"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   7035
      Left            =   0
      ScaleHeight     =   6975
      ScaleWidth      =   8730
      TabIndex        =   0
      Top             =   -15
      Width           =   8790
      Begin VB.PictureBox Pc_Browse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   2205
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   8505
         TabIndex        =   77
         Top             =   75
         Visible         =   0   'False
         Width           =   8535
         Begin MSFlexGridLib.MSFlexGrid Grd_Cancel 
            Height          =   1890
            Left            =   180
            TabIndex        =   78
            Top             =   120
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3334
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777088
            SelectionMode   =   1
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1230
         Left            =   75
         TabIndex        =   71
         Top             =   15
         Width           =   4770
         Begin VB.ComboBox Cbo_Month 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   690
            Width           =   1815
         End
         Begin VB.ComboBox Cbo_Year 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Month : "
            Height          =   270
            Left            =   450
            TabIndex        =   75
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Year : "
            Height          =   270
            Left            =   465
            TabIndex        =   74
            Top             =   345
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2190
         Left            =   75
         TabIndex        =   50
         Top             =   3915
         Width           =   8550
         Begin VB.Label Lbl_Material 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5835
            TabIndex        =   70
            Top             =   975
            Width           =   1890
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Material :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4695
            TabIndex        =   69
            Top             =   990
            Width           =   1080
         End
         Begin VB.Label Lbl_Nett_Rate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1425
            TabIndex        =   68
            Top             =   1290
            Width           =   2100
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Nett Rate :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   375
            TabIndex        =   67
            Top             =   1335
            Width           =   960
         End
         Begin VB.Label Lbl_Date 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1425
            TabIndex        =   66
            Top             =   255
            Width           =   2100
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   705
            TabIndex        =   65
            Top             =   285
            Width           =   645
         End
         Begin VB.Label Lbl_Size 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1425
            TabIndex        =   64
            Top             =   600
            Width           =   2100
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Size :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   570
            TabIndex        =   63
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Lbl_Paper 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5835
            TabIndex        =   62
            Top             =   285
            Width           =   1890
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Paper :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4890
            TabIndex        =   61
            Top             =   330
            Width           =   885
         End
         Begin VB.Label Lbl_Color 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   5835
            TabIndex        =   60
            Top             =   630
            Width           =   1890
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Color :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4920
            TabIndex        =   59
            Top             =   675
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nett :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   375
            TabIndex        =   58
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Lbl_Total_Nett 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            TabIndex        =   57
            Top             =   1635
            Width           =   2100
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Rate :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4125
            TabIndex        =   56
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   570
            TabIndex        =   55
            Top             =   975
            Width           =   780
         End
         Begin VB.Label Lbl_Satuan 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            TabIndex        =   54
            Top             =   945
            Width           =   2100
         End
         Begin VB.Label Lbl_Gross_Rate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5835
            TabIndex        =   53
            Top             =   1320
            Width           =   1890
         End
         Begin VB.Label Lbl_Total_Gross 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5835
            TabIndex        =   52
            Top             =   1665
            Width           =   1890
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Gross :"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4350
            TabIndex        =   51
            Top             =   1710
            Width           =   1425
         End
      End
      Begin Crystal.CrystalReport Crpt 
         Left            =   2595
         Top             =   4695
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   7650
         TabIndex        =   25
         Top             =   6135
         Width           =   975
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   9045
            TabIndex        =   37
            Top             =   0
            Width           =   975
            Begin VB.CommandButton Command11 
               Caption         =   "Cl&ose"
               Height          =   495
               Left            =   120
               TabIndex        =   38
               Top             =   210
               Width           =   720
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   8055
            TabIndex        =   34
            Top             =   0
            Width           =   915
            Begin VB.PictureBox Picture7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   60
               ScaleHeight     =   495
               ScaleWidth      =   735
               TabIndex        =   35
               Top             =   195
               Width           =   795
               Begin VB.CommandButton Command10 
                  Caption         =   "&Print"
                  Height          =   495
                  Left            =   15
                  TabIndex        =   36
                  Top             =   0
                  Width           =   720
               End
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   2385
            TabIndex        =   27
            Top             =   0
            Width           =   2415
            Begin VB.PictureBox Picture6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   90
               ScaleHeight     =   495
               ScaleWidth      =   2175
               TabIndex        =   28
               Top             =   180
               Width           =   2235
               Begin VB.CommandButton Command9 
                  Caption         =   "&Add"
                  Height          =   495
                  Left            =   15
                  TabIndex        =   33
                  Top             =   0
                  Width           =   720
               End
               Begin VB.CommandButton Command8 
                  Caption         =   "&Edit"
                  Height          =   495
                  Left            =   735
                  TabIndex        =   32
                  Top             =   0
                  Width           =   720
               End
               Begin VB.CommandButton Command7 
                  Caption         =   "&Cancel"
                  Height          =   495
                  Left            =   735
                  TabIndex        =   31
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   720
               End
               Begin VB.CommandButton Command6 
                  Caption         =   "&Delete"
                  Height          =   495
                  Left            =   1455
                  TabIndex        =   30
                  Top             =   0
                  Width           =   720
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "&Save"
                  Height          =   495
                  Left            =   15
                  TabIndex        =   29
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   720
               End
            End
         End
         Begin VB.CommandButton Cmd_Close 
            Caption         =   "Cl&ose"
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   210
            Width           =   720
         End
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6660
         TabIndex        =   22
         Top             =   6135
         Width           =   915
         Begin VB.PictureBox Picture5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   60
            ScaleHeight     =   495
            ScaleWidth      =   735
            TabIndex        =   23
            Top             =   195
            Width           =   795
            Begin VB.CommandButton Cmd_Print 
               Caption         =   "&Print"
               Height          =   495
               Left            =   15
               TabIndex        =   24
               Top             =   0
               Width           =   720
            End
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2640
         TabIndex        =   18
         Top             =   6135
         Width           =   3780
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   90
            ScaleHeight     =   480
            ScaleWidth      =   3525
            TabIndex        =   19
            Top             =   180
            Width           =   3585
            Begin VB.CommandButton Cmd_Add 
               Caption         =   "&Add"
               Height          =   480
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   1080
            End
            Begin VB.CommandButton Cmd_Abort 
               Caption         =   "Ca&ncel"
               Height          =   480
               Left            =   2445
               TabIndex        =   21
               Top             =   0
               Width           =   1080
            End
            Begin VB.CommandButton Cmd_Cancel 
               Caption         =   "&Cancel Order"
               Height          =   480
               Left            =   1080
               TabIndex        =   20
               Top             =   0
               Width           =   1365
            End
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   90
         TabIndex        =   12
         Top             =   6135
         Width           =   2310
         Begin VB.PictureBox Picture4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   135
            ScaleHeight     =   495
            ScaleWidth      =   1980
            TabIndex        =   13
            Top             =   180
            Width           =   2040
            Begin VB.CommandButton Cmd_First 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               Picture         =   "Frm_Print_Cancel_PO.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Previous 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   495
               Picture         =   "Frm_Print_Cancel_PO.frx":014A
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Next 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   990
               Picture         =   "Frm_Print_Cancel_PO.frx":0294
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   0
               Width           =   495
            End
            Begin VB.CommandButton Cmd_Last 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1485
               Picture         =   "Frm_Print_Cancel_PO.frx":03DE
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2595
         Left            =   75
         TabIndex        =   1
         Top             =   1260
         Width           =   8565
         Begin VB.TextBox Txt_Brand_Variant 
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1425
            TabIndex        =   79
            Top             =   1350
            Width           =   3045
         End
         Begin VB.ComboBox Cbo_Job_Id 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   615
            Width           =   1770
         End
         Begin VB.ComboBox Cbo_Job_No 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   990
            Width           =   1770
         End
         Begin VB.ComboBox Cbo_Brand 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   270
            Width           =   3045
         End
         Begin VB.ComboBox Cbo_Note 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1650
            Width           =   1890
         End
         Begin VB.ComboBox Cbo_Media 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1710
            Width           =   3045
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Brand Variant :"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   90
            TabIndex        =   80
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label Lbl_Job_id 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6180
            TabIndex        =   47
            Tag             =   " "
            Top             =   2115
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Job Id :"
            Height          =   270
            Left            =   615
            TabIndex        =   46
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Note Code :"
            Height          =   195
            Left            =   4830
            TabIndex        =   44
            Top             =   1695
            Width           =   945
         End
         Begin VB.Label Lbl_Cancel_Date 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5835
            TabIndex        =   42
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Cancel Date :"
            Height          =   270
            Left            =   4575
            TabIndex        =   41
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label lbl_Cancel_No 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5835
            TabIndex        =   40
            Top             =   255
            Width           =   1890
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cancel No :"
            Height          =   270
            Left            =   4830
            TabIndex        =   39
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Lbl_Order_Date 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5835
            TabIndex        =   11
            Top             =   1290
            Width           =   1890
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Order Date :"
            Height          =   270
            Left            =   4710
            TabIndex        =   10
            Top             =   1305
            Width           =   1065
         End
         Begin VB.Label Lbl_PO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   5835
            TabIndex        =   9
            Top             =   945
            Width           =   1890
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Order No :"
            Height          =   270
            Left            =   4830
            TabIndex        =   8
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Job No :"
            Height          =   270
            Left            =   615
            TabIndex        =   7
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Brand :"
            Height          =   270
            Left            =   405
            TabIndex        =   6
            Top             =   300
            Width           =   945
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Client :"
            Height          =   270
            Left            =   390
            TabIndex        =   5
            Top             =   2085
            Width           =   945
         End
         Begin VB.Label Lbl_Client 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1425
            TabIndex        =   4
            Top             =   2055
            Width           =   3045
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Media Name :"
            Height          =   195
            Left            =   210
            TabIndex        =   3
            Top             =   1725
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "Frm_Print_Cancel_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''************************************************
' Form              : Frm_Print_Cancel_PO
' Function          : Generate CO Print
' Created Date      : 16/Aug/2001
' By                :
' Last Update       : 8 September 2004 -> Add Brand Variant -> Diyah
'                   : 28 Jan 2005 SOB BU 2
' Dy 24.10.2007 ->jika Cancel update juga ke Client_purchase_order_detail
'************************************************

Option Explicit
Dim rs_Cancel_Print As New ADODB.Recordset
Dim Sql As String
Dim Old_note As String
Dim rs_Cancel_Temp As New ADODB.Recordset
Dim isNew As Boolean

Private Sub Get_Brand_Info()
    Dim rs_brand As New ADODB.Recordset
    Dim rs_Client As New ADODB.Recordset
    
    Sql = "SELECT de_Flag, Percent_MSC, MSC_On_Flag, VAT_Percent,Brand_Name, Client_Name FROM Ib_Print_Schedule WHERE Job_Id='" & Trim(Cbo_Job_Id.Text) & "' AND Job_No='" & Trim(Cbo_Job_No.Text) & "'"
    rs_brand.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    With rs_brand
        If .EOF = False Then
            Sql = "SELECT brand.*,client.* FROM Brand, client WHERE brand_code ='" & Left(Cbo_Job_Id.Text, 4) & "'"
            Sql = Sql & " and brand.client_code=client.client_code"
            rs_Client.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
            
            If Not rs_Client.EOF Then
                Brand_InFo_Print.ULI = IIf(Trim(rs_Client.Fields("special_client_flag")) = 1, True, False)
                
            End If
            
            rs_Client.Close
            Set rs_Client = Nothing
            
            Brand_InFo_Print.DE_Flag = IIf(.Fields("DE_Flag") = 1, True, False)
            Brand_InFo_Print.Client_Name = IIf(IsNull(.Fields("client_name")), "", .Fields("client_name"))
            Brand_InFo_Print.Brand_Name = IIf(IsNull(.Fields("Brand_Name")), "", .Fields("Brand_Name"))
            
            Brand_InFo_Print.MSC = IIf(IsNull(.Fields("Percent_MSC")), 0, .Fields("Percent_MSC")) / 100
            Select Case .Fields("MSC_On_Flag")
                Case Is = 0
                    Brand_InFo_Print.MSC_Nett_Flag = False
                Case Is = 1
                    Brand_InFo_Print.MSC_Nett_Flag = True
                Case Is = 2
                    Brand_InFo_Print.MSC_Nett_Flag = False
                Case Is = 3
                    Brand_InFo_Print.MSC_Nett_Flag = False
                Case Is = 4
                    Brand_InFo_Print.MSC_Nett_Flag = True
                Case Else
                    Brand_InFo_Print.MSC_Nett_Flag = True
            End Select
            
            Brand_InFo_Print.Media_Agency_Bonus = IIf(IsNull(.Fields("Percent_MSC")), 0, .Fields("Percent_MSC")) / 100
            
            Select Case .Fields("MSC_On_Flag")
                Case Is = 0
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 1
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = True
                Case Is = 2
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 3
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = False
                Case Is = 4
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = True
                Case Else
                    Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag = True
            End Select
            Brand_InFo_Print.Vat = IIf(IsNull(.Fields("VAT_Percent")), 0, .Fields("VAT_Percent"))
        End If
    End With
    
    If rs_brand.State = adStateOpen Then
        rs_brand.Close
    End If
    
    Set rs_brand = Nothing
End Sub

Private Sub Load_Media()
    Dim rs_Load_Media As New ADODB.Recordset
    
    Sql = "SELECT DISTINCT print_code, media_name FROM po_print WHERE job_no='" & Cbo_Job_No.Text & "' AND  post_flag='Unpost'"
    Sql = Sql & " AND job_id='" & Cbo_Job_Id.Text & "' "
    rs_Load_Media.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    Cbo_Media.Clear
    While Not rs_Load_Media.EOF And Not rs_Load_Media.BOF
        Cbo_Media.AddItem rs_Load_Media(0) & " --> " & rs_Load_Media(1)
        rs_Load_Media.MoveNext
    Wend
        
    rs_Load_Media.Close
    Set rs_Load_Media = Nothing
End Sub

Private Sub Load_Brand()
    Dim rs_brand As New ADODB.Recordset
    
    Cbo_Brand.Clear
    With rs_brand
        Sql = "SELECT * FROM brand WHERE brand_code IN (SELECT brand_code FROM Media_Security_Catalog WHERE User_name='" & strLogin_User & "' AND position IN('Implementor','Admin') AND Valid_until > getdate())"
        .Open Sql, ConnERP, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF And Not .BOF
            Cbo_Brand.AddItem .Fields("Brand_Code").Value & " --> " & .Fields("Brand_Name").Value
            .MoveNext
        Wend
    End With
    
    rs_brand.Close
    Set rs_brand = Nothing
End Sub

Private Sub Tombol(show As Boolean)
    Cmd_First.Enabled = show
    Cmd_Previous.Enabled = show
    Cmd_Next.Enabled = show
    Cmd_Last.Enabled = show
    
    Cmd_Add.Enabled = show
    Cmd_Cancel.Enabled = show
    Cmd_Abort.Enabled = show
    
    Cmd_Print.Enabled = show
End Sub

Private Sub Load_PO_Cancel()
    If rs_Cancel_Print.State = adStateOpen Then
        rs_Cancel_Print.Close
    End If
    recDate.Requery
    
    Sql = "SELECT * FROM po_print WHERE job_no='" & Cbo_Job_No.Text & "' AND print_code='" & Get_Print_Kode(Trim(Cbo_Media.Text), "-->") & "' AND cancel_no is not null"
    Sql = Sql & " AND job_id='" & Cbo_Job_Id.Text & "'"
    rs_Cancel_Print.Open Sql, ConnERP, adOpenDynamic, adLockOptimistic
    
    If Not rs_Cancel_Print.EOF And Not rs_Cancel_Print.BOF Then
        Tombol True
        Cmd_Cancel.Enabled = False
        Cmd_Abort.Enabled = False
        fill_Data
    Else
        Tombol False
        Cmd_Add.Enabled = True
        Empty_form
        MsgBox "Data is Empty", vbExclamation, strApplication_Name
    End If
End Sub

Private Sub fill_Data()
    Dim rs_Brand_Name As New ADODB.Recordset
    
    With rs_Cancel_Print
        Lbl_Job_id.Caption = .Fields("job_id").Value
        Lbl_PO.Caption = .Fields("po_number").Value
        Lbl_Order_Date.Caption = Format(.Fields("order_date").Value, "dd/mmm/yyyy")
        Lbl_Date.Caption = Format(IIf(IsNull(.Fields("replace_date").Value), .Fields("booking_date").Value, .Fields("replace_date").Value), "dd/mmm/yyyy")
        
        If IsNull(.Fields("Brand_Variant_Code").Value) = False And IsNull(.Fields("Brand_Variant_name").Value) = False Then
            Txt_Brand_Variant.Text = .Fields("Brand_variant_code").Value & "-->" & .Fields("Brand_variant_name").Value
        Else
            Txt_Brand_Variant.Text = ""
        End If
        
        Lbl_Size.Caption = .Fields("size").Value
        Lbl_Satuan.Caption = .Fields("satuan").Value
        Lbl_Nett_Rate.Caption = Format(.Fields("nett_rate").Value, "#,##0")
        Lbl_Gross_Rate.Caption = Format(.Fields("gross_rate").Value, "#,##0")
              
        If .Fields("Bonus_Flag").Value <> "0" Then
            Lbl_Total_Nett.Caption = 0
            Lbl_Total_Gross.Caption = 0
        Else
            Lbl_Total_Nett.Caption = Format(.Fields("Cancel_Total_Nett").Value, "#,##0")
            Lbl_Total_Gross.Caption = Format(.Fields("Cancel_Total_Gross").Value, "#,##0")
        End If
 
        Lbl_Paper.Caption = .Fields("paper").Value
        Lbl_Color.Caption = .Fields("color").Value
        Lbl_Material.Caption = .Fields("Material").Value & " - " & .Fields("Material_name").Value
        
        Cbo_Note.Clear
        Cbo_Note.AddItem .Fields("note_code").Value
        Cbo_Note.Text = .Fields("note_code").Value
        
        lbl_Cancel_No.Caption = IIf(IsNull(.Fields("cancel_no").Value), "", .Fields("cancel_no").Value)
        Lbl_Cancel_Date.Caption = Format(IIf(IsNull(.Fields("cancel_date").Value), "", .Fields("cancel_date").Value), "dd/mmm/yyyy")
      End With
End Sub

Private Sub Fill_Data_Temp()
    Dim rs_Brand_Name As New ADODB.Recordset
    Dim rs_client_name As New ADODB.Recordset
    Dim rs_bonus As New ADODB.Recordset
    
    With rs_Cancel_Temp
        Lbl_Job_id.Caption = .Fields("job_id").Value
        
        Sql = "SELECT a.client_name FROM client a,brand b WHERE a.client_code=b.client_code AND b.brand_code='" & Left(Trim(Cbo_Brand.Text), 4) & "'"
        rs_client_name.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        
        Lbl_Client.Caption = rs_client_name(0)
        
        rs_client_name.Close
        
        Lbl_PO.Caption = .Fields("po_number").Value
        Lbl_Order_Date.Caption = Format(.Fields("order_date").Value, "dd/mmm/yyyy")
        Lbl_Date.Caption = Format(IIf(IsNull(.Fields("replace_date").Value), .Fields("booking_date").Value, .Fields("replace_date").Value), "dd/mmm/yyyy")
        Lbl_Size.Caption = .Fields("size").Value
        Lbl_Satuan.Caption = .Fields("satuan").Value
        Lbl_Nett_Rate.Caption = Format(.Fields("nett_rate").Value, "#,##0")
        Lbl_Gross_Rate.Caption = Format(.Fields("gross_rate").Value, "#,##0")
        
        Sql = "SELECT bonus_flag FROM ib_print_schedule WHERE job_no='" & Cbo_Job_No.Text & "' AND job_id='" & Cbo_Job_Id.Text & "'"
        rs_bonus.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        If rs_bonus(0) = "1" Then
            Lbl_Total_Nett.Caption = 0
            Lbl_Total_Gross.Caption = 0
        Else
            Lbl_Total_Nett.Caption = Format(.Fields("total_Nett").Value, "#,##0")
            Lbl_Total_Gross.Caption = Format(.Fields("total_gross").Value, "#,##0")
        End If
        
        rs_bonus.Close
        Set rs_bonus = Nothing
        
        Lbl_Paper.Caption = .Fields("paper").Value
        Lbl_Color.Caption = .Fields("color").Value
        Lbl_Material.Caption = .Fields("material").Value & " - " & .Fields("material_name").Value
        
        Cbo_Note.Clear
        Cbo_Note.AddItem .Fields("note_code").Value
        Cbo_Note.Text = .Fields("note_code").Value
        
        lbl_Cancel_No.Caption = ""
        Lbl_Cancel_Date.Caption = ""
        
    End With
End Sub

Private Sub cbo_brand_Click()
    Dim rs_client_name As New ADODB.Recordset
    
    isNew = False
    
    Cbo_Job_No.Clear
    Cbo_Media.Clear
    
    Tombol False
    
    Empty_form
    
    Txt_Brand_Variant.Text = ""
    Lbl_Client.Caption = ""
    
    Load_Job_ID
End Sub

Private Sub Load_Job_ID()
    Dim rs_Load_Job_Id As New ADODB.Recordset
    
    Sql = "SELECT DISTINCT(Job_ID) FROM po_print WHERE substring(rtrim(ltrim(Job_ID)),1,4)='" & Left(Cbo_Brand.Text, 4) & "' AND post_flag='Unpost'"
    Sql = Sql & " AND substring(po_number,3,4)='" & Cbo_Year.Text & "' AND substring(po_number,1,2)='" & Left(Cbo_Month.Text, 2) & "'"
    
    rs_Load_Job_Id.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    Cbo_Job_Id.Clear
    While Not rs_Load_Job_Id.EOF And Not rs_Load_Job_Id.BOF
        Cbo_Job_Id.AddItem rs_Load_Job_Id(0)
        rs_Load_Job_Id.MoveNext
    Wend
    
    rs_Load_Job_Id.Close
    Set rs_Load_Job_Id = Nothing
End Sub

Private Sub Load_Job_No()
    Dim rs_Load_job_no As New ADODB.Recordset
    
    Sql = "SELECT DISTINCT(job_no) FROM po_print WHERE job_id='" & Cbo_Job_Id.Text & "' AND post_flag='Unpost'"
    Sql = Sql & " AND substring(po_number,3,4)='" & Cbo_Year.Text & "' AND substring(po_number,1,2)='" & Left(Cbo_Month.Text, 2) & "'"
    rs_Load_job_no.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    Cbo_Job_No.Clear
    While Not rs_Load_job_no.EOF And Not rs_Load_job_no.BOF
        Cbo_Job_No.AddItem rs_Load_job_no(0)
        rs_Load_job_no.MoveNext
    Wend
    
    rs_Load_job_no.Close
    Set rs_Load_job_no = Nothing
End Sub

Private Sub Cbo_Job_Id_Click()
    If Trim(Cbo_Brand.Text) = "" Then
        MsgBox "Select Brand First", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    Cbo_Media.Clear
    Cbo_Job_No.Clear
    
    Tombol False
    
    Empty_form
    
    Txt_Brand_Variant.Text = ""
    
    Load_Job_No
End Sub

Private Sub Cbo_Job_No_Click()
    Dim Rs_Brand_Variant As New ADODB.Recordset
    
    isNew = False
    
    If Trim(Cbo_Brand.Text) = "" Then
      MsgBox "Select Brand First", vbExclamation, strApplication_Name
      Exit Sub
    End If
    
    If Trim(Cbo_Job_Id.Text) = "" Then
      MsgBox "Select job_id First", vbExclamation, strApplication_Name
      Exit Sub
    End If
    
    Get_Brand_Info
    Lbl_Client.Caption = Brand_InFo_Print.Client_Name
    Cbo_Media.Clear
    
    Tombol False
    Empty_form
    
    Txt_Brand_Variant.Text = ""

    Load_Media
    
    Sql = "SELECT Brand_variant_Code,Brand_Variant_Name FROM PO_Print WHERE JOB_ID='" & Cbo_Job_Id.Text & "' AND JOB_No='" & Cbo_Job_No & "'"
    Rs_Brand_Variant.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Not Rs_Brand_Variant.EOF Then
        Txt_Brand_Variant.Text = Rs_Brand_Variant("Brand_variant_Code").Value & "-->" & Rs_Brand_Variant("Brand_variant_Name").Value
        Rs_Brand_Variant.Close
    End If
End Sub

Private Sub Cbo_Media_Click()
    If Trim(Me.Cbo_Media.Text) = "" Then
        MsgBox "Select Media First", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    Load_PO_Cancel
End Sub

Private Sub Cbo_Month_Click()
    Cbo_Brand.Clear
    Cbo_Job_No.Clear
    Cbo_Media.Clear
    Cbo_Job_Id.Clear
    Lbl_Client.Caption = ""
    
    Tombol False
    
    Empty_form
    Load_Brand
End Sub

Private Sub Cbo_Year_Click()
    Cbo_Brand.Clear
    Cbo_Job_No.Clear
    Cbo_Media.Clear
    Cbo_Job_Id.Clear
    Lbl_Client.Caption = ""
    
    Tombol False
    
    Empty_form
    Load_Brand
End Sub

Private Sub Cmd_Abort_Click()
    If Cmd_Abort.Caption = "Ca&ncel" Then
        isNew = False
        Cbo_Year.Enabled = True
        Cbo_Month.Enabled = True
        Cmd_Cancel.Enabled = False
        Cmd_Add.Enabled = True
        Cmd_Print.Enabled = True
        Cmd_Close.Enabled = True
        Cmd_Abort.Enabled = False
        
        Cbo_Media.Enabled = True
        Cbo_Brand.Enabled = True
        Cbo_Job_No.Enabled = True
        Cbo_Job_Id.Enabled = True
        Pc_Browse.Visible = False
        Cmd_Cancel.Caption = "&Cancel Order"
        Cbo_Note.Enabled = False
        Load_PO_Cancel
    End If
End Sub
Private Sub Create_temp_data()
    Set rs_Cancel_Temp = Nothing
    Set rs_Cancel_Temp = New ADODB.Recordset
    
    With rs_Cancel_Temp.Fields
        .Append "PO_Number", adVarChar, 12, adFldMayBeNull
        .Append "Job_No", adVarChar, 13, adFldMayBeNull
        .Append "Job_id", adVarChar, 13, adFldMayBeNull
        .Append "Booking_Date", adDate, 8, adFldMayBeNull
        .Append "Size", adVarChar, 30, adFldMayBeNull
        .Append "satuan", adVarChar, 30, adFldMayBeNull
        .Append "paper", adVarChar, 30, adFldMayBeNull
        .Append "Color", adVarChar, 30, adFldMayBeNull
        .Append "Material", adVarChar, 1, adFldMayBeNull
        .Append "Material_name", adVarChar, 50, adFldMayBeNull
        .Append "Order_Date", adDate, 8, adFldMayBeNull
        .Append "Nett_Rate", adCurrency, 9, adFldMayBeNull
        .Append "gross_Rate", adCurrency, 9, adFldMayBeNull
        .Append "total_Nett", adCurrency, 9, adFldMayBeNull
        .Append "total_gross", adCurrency, 9, adFldMayBeNull
        .Append "note_code", adVarChar, 15, adFldMayBeNull
        .Append "replace_Date", adDate, 8, adFldMayBeNull
    End With
    
    rs_Cancel_Temp.Open , , adOpenDynamic, adLockOptimistic
End Sub

Private Sub Prepare_Data()
    Dim rs_Not_cancel As New ADODB.Recordset
    
    Create_temp_data
    
    recDate.Requery
    Sql = "select * FROM po_print WHERE job_no='" & Cbo_Job_No.Text & "' AND print_code='" & Get_Print_Kode(Trim(Cbo_Media.Text), "-->") & "' AND cancel_no is null AND print_flag > 0 AND UPPER(post_flag) = 'UNPOST'" ' AND substring(po_number,3,4)='" & CStr(Year(rs_date(0))) & "' ORDER BY po_number"
    
    rs_Not_cancel.CursorLocation = adUseClient
    rs_Not_cancel.Open Sql, ConnERP, adOpenDynamic, adLockOptimistic
    
    With rs_Cancel_Temp
        While Not rs_Not_cancel.EOF And Not rs_Not_cancel.BOF
            .AddNew
            .Fields("po_number").Value = rs_Not_cancel.Fields("po_number").Value
            .Fields("Job_No").Value = rs_Not_cancel.Fields("Job_No").Value
            .Fields("Job_id").Value = rs_Not_cancel.Fields("Job_id").Value
            
            .Fields("booking_date").Value = rs_Not_cancel.Fields("booking_date").Value
            .Fields("size").Value = rs_Not_cancel.Fields("size").Value
            .Fields("satuan").Value = rs_Not_cancel.Fields("satuan").Value
            .Fields("paper").Value = rs_Not_cancel.Fields("paper").Value
            .Fields("color").Value = rs_Not_cancel.Fields("color").Value
            .Fields("material").Value = rs_Not_cancel.Fields("material").Value
            .Fields("material_name").Value = rs_Not_cancel.Fields("material_name").Value
            .Fields("order_date").Value = rs_Not_cancel.Fields("order_date").Value
            .Fields("nett_rate").Value = rs_Not_cancel.Fields("nett_rate").Value
            .Fields("gross_rate").Value = rs_Not_cancel.Fields("gross_rate").Value
            .Fields("total_nett").Value = rs_Not_cancel.Fields("total_nett").Value
            .Fields("total_gross").Value = rs_Not_cancel.Fields("total_nett").Value
            .Fields("note_code").Value = rs_Not_cancel.Fields("note_code").Value
            .Fields("replace_date").Value = rs_Not_cancel.Fields("replace_date").Value
            .Update
            rs_Not_cancel.MoveNext
        Wend
        
        rs_Not_cancel.Close
        Set rs_Not_cancel = Nothing
    End With
End Sub

Private Sub Show_data_To_Grd()
    Dim Rs_Data_PO As New ADODB.Recordset
    Dim baris As Integer
    Dim Kolom As Integer
           
    Sql = "SELECT * FROM po_print WHERE job_no='" & Cbo_Job_No.Text & "'"
    Sql = Sql & " AND job_id='" & Cbo_Job_Id.Text & "'"
    Sql = Sql & " AND print_code ='" & Get_Print_Kode(Trim(Cbo_Media.Text), "-->") & "'"
    Sql = Sql & " AND cancel_no is null"
    Sql = Sql & " AND post_flag='Unpost'"
    Rs_Data_PO.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Rs_Data_PO.EOF = True Then
         isNew = False
         MsgBox "No PO Can't Cancel", vbExclamation, strApplication_Name
    Else
        lbl_Cancel_No.Caption = ""
        Lbl_Cancel_Date.Caption = ""
        Lbl_PO.Caption = ""
        Lbl_Order_Date.Caption = ""
        Cbo_Note.Clear

         'set button
         Cbo_Brand.Enabled = False
        Cbo_Job_Id.Enabled = False
        Cbo_Job_No.Enabled = False
        Cbo_Media.Enabled = False
        
        Cmd_Add.Enabled = False
        Cmd_Cancel.Enabled = True
        Cmd_Abort.Caption = "Ca&ncel"
        Cmd_Abort.Enabled = True

        Cmd_First.Enabled = False
        Cmd_Previous.Enabled = False
        Cmd_Next.Enabled = False
        Cmd_Last.Enabled = False
        
        Cmd_Print.Enabled = False
        Cmd_Close.Enabled = False
        
        Cbo_Year.Enabled = False
        Cbo_Month.Enabled = False
        
        Pc_Browse.Visible = True
        Pc_Browse.Left = Frame2.Left
        Pc_Browse.Top = Frame2.Top
        
        With Grd_Cancel
           .cols = 12
           .Rows = 1
           .FormatString = " PO No | Booking Date | Total Gross | Total Nett | Size |  Satuan  | Paper | Color  | Material | Order Date | Gross Rate | Nett Rate "
           
           .ColWidth(0) = 1400 'PO
           .ColWidth(1) = 1400 'media
           .ColWidth(2) = 1300 'gross
           .ColWidth(3) = 1300 'nett
           .ColWidth(4) = 800 'size
           .ColWidth(5) = 1500 'satuan
           .ColWidth(6) = 500  'color
           .ColWidth(7) = 500 'paper
           .ColWidth(8) = 1500 'material
           .ColWidth(9) = 1500 'order date
           .ColWidth(10) = 1500 'gross rate
           .ColWidth(11) = 1500 'nett rate
           
           baris = 1
           While Not Rs_Data_PO.EOF
               .AddItem Rs_Data_PO("PO_Number") & vbTab _
                         & Format(IIf(IsNull(Rs_Data_PO("Replace_Date")), Rs_Data_PO("Booking_Date"), Rs_Data_PO("Replace_Date")), "dd/mmm/yyyy") & vbTab _
                         & Format(Rs_Data_PO("Total_Gross"), "#,##0") & vbTab _
                         & Format(Rs_Data_PO("Total_nett"), "#,##0") & vbTab _
                         & Rs_Data_PO("size") & vbTab _
                         & Rs_Data_PO("Satuan") & vbTab _
                         & Rs_Data_PO("color") & vbTab _
                         & Rs_Data_PO("paper") & vbTab _
                         & IIf(IsNull(Rs_Data_PO("Material_code_Replace")), Rs_Data_PO("Material") & " - " & Rs_Data_PO("Material_Name"), Rs_Data_PO("Material_code_Replace") & " - " & Rs_Data_PO("Material_Replace")) & vbTab _
                         & Format(Rs_Data_PO("Order_Date"), "dd/mmm/yyyy") & vbTab _
                         & Format(Rs_Data_PO("Gross_Rate"), "#,##0.0") & vbTab _
                         & Format(Rs_Data_PO("Nett_Rate"), "#,##0.0")
               Rs_Data_PO.MoveNext
           Wend
        End With
    End If
    
    Rs_Data_PO.Close
    Set Rs_Data_PO = Nothing
End Sub

Private Sub Cmd_Add_Click()
    'Check apakah Implemetor Brand
    If Not IsValidAccess(strLogin_FullName, "Implementor", Left(Cbo_Brand.Text, 4)) Then
        MsgBox "Access Denied...", vbExclamation, strApplication_Name
        Exit Sub
    End If

    isNew = True
    Show_data_To_Grd
End Sub

Function Generate_Cancel_No()
    Dim rs_LastNo As New ADODB.Recordset
    Dim Last_No As Double
    Dim Nomer As String
    Dim a As Integer
    
    recDate.Requery
    
    Sql = "SELECT * FROM Last_Cancel_PO_Print WHERE year= " & CInt(Mid(Trim(Lbl_PO.Caption), 3, 4))
    rs_LastNo.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Not rs_LastNo.EOF And Not rs_LastNo.BOF Then
        Last_No = rs_LastNo("last_number") + 1
    Else
        Last_No = 1
    End If
    
    rs_LastNo.Close
    Set rs_LastNo = Nothing
    
    Nomer = ""
    Nomer = Format(CStr(Last_No), "00000")
    
    Generate_Cancel_No = Left(Trim(Lbl_PO.Caption), 6) & "C" & Nomer
    
    Sql = "DELETE FROM Last_Cancel_PO_Print WHERE year=" & CInt(Mid(Trim(Lbl_PO.Caption), 3, 4))
    ConnERP.Execute Sql
    Sql = "INSERT INTO Last_Cancel_PO_Print VALUES(" & CInt(Mid(Trim(Lbl_PO.Caption), 3, 4)) & "," & Last_No & "  ) "
    ConnERP.Execute Sql
End Function

Function MSC_Persen(Kode As String) As Double
    Dim Rs_Msc As New ADODB.Recordset
    
    Sql = "SELECT msc FROM brand WHERE brand_code='" & Kode & "'"
    Rs_Msc.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    MSC_Persen = CDbl(Rs_Msc(0) / 100)
    
    Rs_Msc.Close
    Set Rs_Msc = Nothing
End Function

Function is_MSC(Kode As String) As Boolean
    Dim Rs_Msc As New ADODB.Recordset
    
    Sql = "SELECT msc_nett_flag FROM brand WHERE brand_code='" & Kode & "'"
    Rs_Msc.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Rs_Msc(0) = 1 Or Rs_Msc(0) = 4 Then
        is_MSC = True
    Else
        is_MSC = False
    End If
    
    Rs_Msc.Close
    Set Rs_Msc = Nothing
End Function

Function is_Bonus(Brand As String) As Boolean
    Dim Rs_Msc As New ADODB.Recordset
    
    Sql = "SELECT Media_agency_bonus_nett_flag FROM brand WHERE brand_code='" & Brand & "'"
    Rs_Msc.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Rs_Msc(0) = 1 Or Rs_Msc(0) = 4 Then
        is_Bonus = True
    Else
        is_Bonus = False
    End If
    
    Rs_Msc.Close
    Set Rs_Msc = Nothing
End Function

Function Bonus_Persen(Brand As String) As Double
    Dim Rs_Msc As New ADODB.Recordset
    
    Sql = "SELECT media_agency_bonus FROM brand WHERE brand_code='" & Brand & "'"
    Rs_Msc.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    Bonus_Persen = CDbl(Rs_Msc(0) / 100)
    
    Rs_Msc.Close
    Set Rs_Msc = Nothing
End Function


Private Sub cmd_cancel_Click()
    Dim rs_NOte_cancel As New ADODB.Recordset
    Dim V_MSC As Double, V_DE As Double
    Dim Rs_De As New ADODB.Recordset
    Dim rs_bonus As New ADODB.Recordset
    Dim V_Nett As Double
    Dim V_gross As Double
    Dim rs_Nett As New ADODB.Recordset
    Dim Rs_Data_PO As New ADODB.Recordset
    
    Dim Is_have_PO As Integer
    Dim Vat As Double
    
    If Cmd_Cancel.Caption = "&Cancel Order" Then
       
        Cmd_Cancel.Caption = "&Save"
        Cbo_Note.Enabled = True
        Old_note = Cbo_Note.Text
        Load_Note
        
        Sql = "SELECT note_code_cancel FROM print_information"
        rs_NOte_cancel.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        Cbo_Note.Text = rs_NOte_cancel(0)
        rs_NOte_cancel.Close
        Set rs_NOte_cancel = Nothing
       
        With Grd_Cancel
            If .Row > 0 Then
                Lbl_PO.Caption = .TextMatrix(.Row, 0)
                Lbl_Order_Date.Caption = .TextMatrix(.Row, 9)
                Lbl_Date.Caption = .TextMatrix(.Row, 1)
                Lbl_Size.Caption = .TextMatrix(.Row, 4)
                Lbl_Satuan.Caption = .TextMatrix(.Row, 5)
                Lbl_Nett_Rate.Caption = .TextMatrix(.Row, 11)
                Lbl_Total_Nett.Caption = .TextMatrix(.Row, 3)
                Lbl_Paper.Caption = .TextMatrix(.Row, 6)
                Lbl_Color.Caption = .TextMatrix(.Row, 7)
                Lbl_Material.Caption = .TextMatrix(.Row, 8)
                Lbl_Gross_Rate.Caption = .TextMatrix(.Row, 10)
                Lbl_Total_Gross.Caption = .TextMatrix(.Row, 2)

                Pc_Browse.Visible = False
            Else
                MsgBox "Select PO first", vbExclamation, strApplication_Name
            End If
        End With
    Else
        On Error GoTo my_err
        
        Is_have_PO = Get_Client_Have_PO(Trim(Mid(Cbo_Brand.Text, 1, 4)))
        Load_Taxes_Data
        Vat = Taxes.Vat
        Vat = Vat / 100
        
        
        ConnERP.BeginTrans
        If Brand_InFo_Print.ULI Then
            Sql = "SELECT bonus_flag FROM ib_print_schedule WHERE job_no='" & Cbo_Job_No.Text & "' AND job_id='" & Cbo_Job_Id.Text & "'"
            rs_bonus.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
            
            If rs_bonus(0) = "1" Then
                '=========== masuk ke budget
                Sql = "SELECT total_nett, total_gross FROM po_print WHERE po_number='" & Lbl_PO.Caption & "'"
                rs_Nett.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                
                V_Nett = rs_Nett(0)
                V_gross = rs_Nett(1)
                
                rs_Nett.Close
                Set rs_Nett = Nothing
            
                If Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag Then
                    V_MSC = Brand_InFo_Print.Media_Agency_Bonus * V_Nett
                Else
                    V_MSC = Brand_InFo_Print.Media_Agency_Bonus * V_gross
                End If
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_DE = Rs_De(0)
                
                Rs_De.Close
                Set Rs_De = Nothing
                
                Sql = "UPDATE ULI_Budget_Control_Actual SET spots= spots - 1, purchase_order_msc = purchase_order_msc - " & V_MSC & ", purchase_order_de = purchase_order_de - " & V_DE & ", purchase_order_estimate = purchase_order_estimate - " & (V_DE + V_MSC) & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
'                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
'                    Sql = Sql & " MSC=msc - " & V_MSC & ","
'                    Sql = Sql & " DE = DE - " & V_DE & ", "
'                    Sql = Sql & " sub_Total = sub_Total - " & (V_DE + V_MSC) & ","
'                    Sql = Sql & " VAT= (" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "),"
'                    Sql = Sql & " Grand_total=(sub_Total - " & (V_DE + V_MSC) & ")+((" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "))"
'                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
'                    CONNERp.Execute Sql

                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " MSC=msc - " & V_MSC & ","
                    Sql = Sql & " DE = DE - " & V_DE & ", "
                    Sql = Sql & " sub_Total = sub_Total - (" & V_DE & " + " & V_MSC & ")"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                
                End If
                
            ElseIf rs_bonus(0) = "0" Then
                '=========== masuk ke budget
                Sql = "SELECT total_nett, total_gross FROM po_print WHERE po_number='" & Lbl_PO.Caption & "'"
                rs_Nett.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_Nett = rs_Nett(0)
                V_gross = rs_Nett(1)
                
                rs_Nett.Close
                Set rs_Nett = Nothing
            
                If Brand_InFo_Print.MSC_Nett_Flag Then
                    V_MSC = Brand_InFo_Print.MSC * V_Nett
                Else
                    V_MSC = Brand_InFo_Print.MSC * V_gross
                End If
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
               
                V_DE = Rs_De(0)
               
                Rs_De.Close
                Set Rs_De = Nothing
                
                Sql = "UPDATE ULI_Budget_Control_Actual SET spots= spots - 1, purchase_order_netto = purchase_order_netto - " & CDbl(Lbl_Total_Nett.Caption) & ", purchase_order_msc = purchase_order_msc - " & V_MSC & ", purchase_order_de = purchase_order_de - " & V_DE & " , purchase_order_estimate = purchase_order_estimate - " & (V_DE + V_MSC + CDbl(Lbl_Total_Nett.Caption)) & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
'                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
'                    Sql = Sql & " Nett=Nett -" & CDbl(Lbl_Total_Nett.Caption) & ",MSC=msc - " & V_MSC & ","
'                    Sql = Sql & " DE = DE - " & V_DE & ", "
'                    Sql = Sql & " sub_Total = sub_Total -  (" & V_DE & " + " & V_MSC & " + " & CDbl(Lbl_Total_Nett.Caption) & "),"
'                    Sql = Sql & " VAT= (" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "),"
'                    Sql = Sql & " Grand_total=(sub_Total - " & (V_DE + V_MSC) & ")+((" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "))"
'                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
'                    CONNERp.Execute Sql
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " Nett=Nett -" & CDbl(Lbl_Total_Nett.Caption) & ","
                    Sql = Sql & " MSC=msc - " & V_MSC & ","
                    Sql = Sql & " DE = DE - " & V_DE & " "
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " sub_Total = nett+msc+de"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                End If
            ElseIf rs_bonus(0) = "2" Then
                '=========== masuk ke budget
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_DE = Rs_De(0)
                
                Rs_De.Close
                Set Rs_De = Nothing
                
                Sql = "UPDATE ULI_Budget_Control_Actual SET spots= spots - 1, purchase_order_de = purchase_order_de - " & V_DE & " , purchase_order_estimate = purchase_order_estimate - " & V_DE & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
'                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
'                    Sql = Sql & " DE = DE - " & V_DE & ", "
'                    Sql = Sql & " sub_Total = sub_Total - " & V_DE & ","
'                    Sql = Sql & " VAT= (" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "),"
'                    Sql = Sql & " Grand_total=(sub_Total - " & (V_DE + V_MSC) & ")+((" & Vat & ")*(sub_Total - " & (V_DE + V_MSC) & "))"
'                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
'                    CONNERp.Execute Sql
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " DE = DE - " & V_DE & ", "
                    Sql = Sql & " sub_Total = sub_Total - " & V_DE & ""
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
               End If
            End If
            rs_bonus.Close
            Set rs_bonus = Nothing
        Else
            'SOB BU2
            Sql = "SELECT bonus_flag FROM ib_print_schedule WHERE job_no='" & Cbo_Job_No.Text & "' AND job_id='" & Cbo_Job_Id.Text & "'"
            rs_bonus.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
            
            If rs_bonus(0) = "1" Then
                '=========== masuk ke budget
                Sql = "SELECT total_nett, total_gross FROM po_print WHERE po_number='" & Lbl_PO.Caption & "'"
                rs_Nett.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                
                V_Nett = rs_Nett(0)
                V_gross = rs_Nett(1)
                
                rs_Nett.Close
                Set rs_Nett = Nothing
            
                If Brand_InFo_Print.Media_Agency_Bonus_Nett_Flag Then
                    V_MSC = Brand_InFo_Print.Media_Agency_Bonus * V_Nett
                Else
                    V_MSC = Brand_InFo_Print.Media_Agency_Bonus * V_gross
                End If
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_DE = Rs_De(0)
                
                Rs_De.Close
                Set Rs_De = Nothing
                
                Sql = "UPDATE BU2_Budget_Control_Quotation SET Actual_Spots = Actual_Spots - 1, "
                Sql = Sql & " Actual_MSC = Actual_MSC - " & V_MSC & ", "
                Sql = Sql & " Actual_DE = Actual_DE - " & V_DE & ", "
                Sql = Sql & " Actual_Total = Actual_Total - " & (V_DE + V_MSC) & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " MSC=msc - " & V_MSC & ","
                    Sql = Sql & " DE = DE - " & V_DE & ", "
                    Sql = Sql & " sub_Total = sub_Total - (" & V_DE & " + " & V_MSC & ")"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                End If
            ElseIf rs_bonus(0) = "0" Then
                '=========== masuk ke budget
                Sql = "SELECT total_nett, total_gross FROM po_print WHERE po_number='" & Lbl_PO.Caption & "'"
                rs_Nett.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_Nett = rs_Nett(0)
                V_gross = rs_Nett(1)
                
                rs_Nett.Close
                Set rs_Nett = Nothing
            
                If Brand_InFo_Print.MSC_Nett_Flag Then
                    V_MSC = Brand_InFo_Print.MSC * V_Nett
                Else
                    V_MSC = Brand_InFo_Print.MSC * V_gross
                End If
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
               
                V_DE = Rs_De(0)
               
                Rs_De.Close
                Set Rs_De = Nothing
                
                Sql = "UPDATE BU2_Budget_Control_Quotation SET Actual_Spots = Actual_Spots - 1, "
                Sql = Sql & " Actual_Netto = Actual_Netto - " & CDbl(Lbl_Total_Nett.Caption) & ", "
                Sql = Sql & " Actual_MSC = Actual_MSC - " & V_MSC & ", "
                Sql = Sql & " Actual_DE = Actual_DE - " & V_DE & ", "
                Sql = Sql & " Actual_Total = Actual_Total - " & (V_DE + V_MSC + CDbl(Lbl_Total_Nett.Caption)) & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " Nett=Nett -" & CDbl(Lbl_Total_Nett.Caption) & ","
                    Sql = Sql & " MSC=msc - " & V_MSC & ","
                    Sql = Sql & " DE = DE - " & V_DE & " "
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " sub_Total = nett+msc+de"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                End If
            ElseIf rs_bonus(0) = "2" Then
                '=========== masuk ke budget
                
                Sql = "SELECT de_money FROM ib_print_schedule WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_no='" & Cbo_Job_No.Text & "'"
                Rs_De.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
                V_DE = Rs_De(0)
                
                Rs_De.Close
                Set Rs_De = Nothing
                
                
                Sql = "UPDATE BU2_Budget_Control_Quotation SET Actual_Spots = Actual_Spots - 1, "
                Sql = Sql & " Actual_DE = Actual_DE - " & V_DE & ", "
                Sql = Sql & " Actual_Total = Actual_Total - " & (V_DE) & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                ConnERP.Execute Sql
                
                If Is_have_PO = 1 Then
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " DE = DE - " & V_DE & ", "
                    Sql = Sql & " sub_Total = sub_Total - " & V_DE & ""
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Client_Purchase_Order_Detail SET "
                    Sql = Sql & " VAT= (" & Vat & ")*(sub_total),"
                    Sql = Sql & " Grand_total=(sub_total)+((" & Vat & ")*(sub_total))"
                    Sql = Sql & " WHERE job_id='" & Cbo_Job_Id.Text & "' AND job_number='" & Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
               End If
            End If
            rs_bonus.Close
            Set rs_bonus = Nothing
        End If
                
        recDate.Requery
        lbl_Cancel_No.Caption = Generate_Cancel_No
        Lbl_Cancel_Date.Caption = Format(recDate(0), "dd/mmm/yyyy")
        
        
        Sql = "SELECT * FROM PO_PRINT WHERE po_number='" & Lbl_PO.Caption & "'"
        Rs_Data_PO.CursorLocation = adUseClient
        Rs_Data_PO.Open Sql, ConnERP, adOpenForwardOnly, adLockReadOnly
        
        If Not Rs_Data_PO.EOF Then
            
            'update quot
            Select Case Rs_Data_PO("Bonus_Flag")
                Case "0":   Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Gross=PO_Total_Gross - " & Rs_Data_PO("Total_Gross") & ", PO_Total_Nett = PO_Total_Nett - " & Rs_Data_PO("Total_Nett")
                    Sql = Sql & ", PO_MSC = PO_MSC - " & Rs_Data_PO("Total_MSC") & ", PO_DE = PO_DE - " & Rs_Data_PO("DE") & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
    
                    Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Cost = PO_MSC + PO_Total_Nett + PO_DE "
                    Sql = Sql & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                
                Case "1":   Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Gross=PO_Total_Gross - " & Rs_Data_PO("Total_Gross") & ", PO_Total_Nett = PO_Total_Nett - " & Rs_Data_PO("Total_Nett")
                    Sql = Sql & ", PO_MSC = PO_MSC - " & Rs_Data_PO("Total_MSC") & ", PO_DE = PO_DE - " & Rs_Data_PO("DE") & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Cost = PO_MSC + PO_DE "
                    Sql = Sql & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                
                Case "2":   Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Gross=0, PO_Total_Nett = 0"
                    Sql = Sql & ", PO_MSC = 0, PO_DE = PO_DE - " & Rs_Data_PO("DE") & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
                    
                    Sql = "UPDATE Ib_Print_Schedule SET PO_Total_Cost = PO_DE "
                    Sql = Sql & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
                    ConnERP.Execute Sql
            End Select
            
            Sql = "UPDATE Ib_Print_Schedule SET PO_VAT = PO_Total_Cost * " & (Brand_InFo_Print.Vat / 100)
            Sql = Sql & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
            ConnERP.Execute Sql
            
            Sql = "UPDATE Ib_Print_Schedule SET PO_Grand_Total = PO_Total_Cost + PO_VAT"
            Sql = Sql & " WHERE Job_Id='" & Me.Cbo_Job_Id.Text & "' AND Job_No='" & Me.Cbo_Job_No.Text & "'"
            ConnERP.Execute Sql
            
        End If
        
        Rs_Data_PO.Close
        Set Rs_Data_PO = Nothing
        
        ConnERP.CommitTrans
        MsgBox "Data has been saved", vbExclamation, strApplication_Name
        
        Cbo_Year.Enabled = True
        Cbo_Month.Enabled = True
        Cmd_Cancel.Enabled = False
        Cmd_Add.Enabled = True
        Cmd_Print.Enabled = True
        Cmd_Close.Enabled = True
        Cmd_Abort.Enabled = False
        
        Cbo_Media.Enabled = True
        Cbo_Brand.Enabled = True
        Cbo_Job_No.Enabled = True
        Cbo_Job_Id.Enabled = True
        Pc_Browse.Visible = False
        Cmd_Cancel.Caption = "&Cancel Order"
        Cbo_Note.Enabled = False
        isNew = False
        Load_PO_Cancel
    End If
    
    Exit Sub
my_err:
    ConnERP.RollbackTrans
    MsgBox Err.Description, vbExclamation, strApplication_Name
End Sub

Private Sub Cmd_Close_Click()
    Unload Me
End Sub

Private Sub Cmd_First_Click()
    If isNew = False Then
        rs_Cancel_Print.MoveFirst
        fill_Data
    End If
End Sub

Private Sub Cmd_Last_Click()
    If isNew = False Then
        rs_Cancel_Print.MoveLast
        fill_Data
    End If
End Sub

Private Sub Cmd_Next_Click()
    If isNew = False Then
        rs_Cancel_Print.MoveNext
        If rs_Cancel_Print.EOF = True Then
        
            rs_Cancel_Print.MoveLast
            MsgBox "Last Record", vbExclamation, strApplication_Name
        End If
        
        fill_Data
    End If
End Sub

Private Sub Cmd_Previous_Click()
    If isNew = False Then
        rs_Cancel_Print.MovePrevious
        
        If rs_Cancel_Print.BOF = True Then
            rs_Cancel_Print.MoveFirst
            MsgBox "First Record", vbExclamation, strApplication_Name
        End If
        
        fill_Data
    End If
End Sub
Private Sub Empty_form()
    
    Lbl_Job_id.Caption = ""
    Lbl_PO.Caption = ""
    Lbl_Order_Date.Caption = ""
    Cbo_Note.Clear
    
    Lbl_Date.Caption = ""
    Lbl_Size.Caption = ""
    Lbl_Satuan.Caption = ""
    Lbl_Nett_Rate.Caption = 0
    Lbl_Gross_Rate.Caption = 0
    Lbl_Total_Nett.Caption = 0
    Lbl_Total_Gross.Caption = 0
    Lbl_Paper.Caption = ""
    Lbl_Color.Caption = ""
    Lbl_Material.Caption = ""
    
    lbl_Cancel_No.Caption = ""
    Lbl_Cancel_Date.Caption = ""
End Sub

Function Is_mmCl() As Boolean
    Dim rs_Cek_MMC As New ADODB.Recordset
    
    Sql = "SELECT isMMC FROM print_size_catalog WHERE upper(size_code)='" & UCase(Trim(Lbl_Satuan.Caption)) & "'"
    rs_Cek_MMC.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    If Not rs_Cek_MMC.BOF And Not rs_Cek_MMC.EOF Then
        If rs_Cek_MMC(0) = 1 Then
            Is_mmCl = True
        Else
            Is_mmCl = False
        End If
    Else
        Is_mmCl = False
    End If
    
    rs_Cek_MMC.Close
    Set rs_Cek_MMC = Nothing
    
End Function

Private Sub Cmd_Print_Click()
    Dim rs_Supervisor As New ADODB.Recordset
    Dim rs_PO_Manager As New ADODB.Recordset
    Dim Posisi As Integer
    Dim Rs_Received_By As New ADODB.Recordset
    Dim strSql As String
    
    Dim strSign As String
    Dim StrDescription  As String
    Dim strSize As String
    
    If Trim(Lbl_PO.Caption) = "" Then
        MsgBox "No Data to Print", vbExclamation, strApplication_Name
        Exit Sub
    End If
    
    Sql = "UPDATE po_print SET  cancel_Print_Flag = cancel_Print_Flag + 1  WHERE po_number ='" & Trim(Lbl_PO.Caption) & "'"
    ConnERP.Execute Sql
    
    Sql = "UPDATE po_print SET cancel_Print_Flag='1' WHERE po_number ='" & Trim(Lbl_PO.Caption) & "'"
    ConnERP.Execute Sql
    
    strPaperTypeOld = strPaperTypeCO
    
    If UCase(Trim(strPaperTypeCO)) = "A4" Then
        If MsgBox("Do you want to print Pre-Printed PO?", vbYesNo, strApplication_Name) = vbYes Then
            strPaperTypeCO = "LETTER"
        End If
    End If

    If UCase(Trim(strPaperTypeCO)) = "A4" Then
        sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source= c:\ERPTempDB\Rpt.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
    
        Set dfwConn = New Connection
        dfwConn.Open sConnect
        
        strSql = "DELETE FROM COPrint"
        dfwConn.Execute strSql
        
        strSql = ""
        strSql = strSql & " SELECT Client_Code,Business_Unit FROM client WHERE Client_name = '" & Trim(Lbl_Client.Caption) & "'"
        If rs_Get_Address.State = adStateOpen Then
          rs_Get_Address.Close
          Set rs_Get_Address = Nothing
        End If
        rs_Get_Address.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        
        strSql = "SELECT A.PO_Number, A.Job_ID, A.Job_No, "
        strSql = strSql & " A.Brand_Variant_Name , A.Media_Name, "
        strSql = strSql & " A.Company_Name, A.Address1, A.Address2, "
        strSql = strSql & " A.Address3, A.Booking_Date, "
        strSql = strSql & " A.Booking_Month, A.Material_Name, "
        strSql = strSql & " A.Cancel_No, A.Cancel_Date, "
        strSql = strSql & " A.Replace_Date, A.Cancel_Print_Flag, "
        strSql = strSql & " B.Paper_Name,"
        strSql = strSql & " C.Color_Name,"
        strSql = strSql & " D.Description,"
        strSql = strSql & " E.Brand_Name , E.Client_Name"
        strSql = strSql & " FROM"
        strSql = strSql & " PO_Print A,"
        strSql = strSql & " Print_Paper_Catalog B,"
        strSql = strSql & " Print_Color_Catalog C,"
        strSql = strSql & " Print_Catalog_Note D,"
        strSql = strSql & " IB_Print_Schedule E"
        strSql = strSql & " WHERE"
        strSql = strSql & " A.Paper = B.Paper_Code AND"
        strSql = strSql & " A.Color = C.Color_Code AND"
        strSql = strSql & " A.Note_Code = D.Note_Code AND"
        strSql = strSql & " A.Job_ID = E.Job_Id AND"
        strSql = strSql & " A.Job_No = E.Job_No AND"
        strSql = strSql & " A.Cancel_No IS NOT NULL"
        strSql = strSql & " AND A.cancel_no='" & Trim(lbl_Cancel_No.Caption) & "'"
        strSql = strSql & " Order By A.Cancel_No Asc"
    
        Rs_Received_By.Open strSql, ConnERP, adOpenStatic, adLockReadOnly
        
        If MsgBox("Do you want print to Winfax?", vbYesNo + vbQuestion, strApplication_Name) = vbYes Then
            'IntPOCopy = 1
            blnIsWinfax = True
        Else
'            IntPOCopy = GetCopyNumber(Left(Trim(Cbo_Job_Id.Text), 4))
            blnIsWinfax = False
        End If
        
        IntPOCopy = GetCopyNumber(Left(Trim(Cbo_Job_Id.Text), 4))
        While Not Rs_Received_By.EOF
            strQuery = "INSERT INTO COPrint VALUES ( "
            strQuery = strQuery & " '" & Rs_Received_By.Fields("PO_Number").Value & "' ,"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("Job_id").Value & "' ,"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("Job_no").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("Brand_variant_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("media_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("company_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("address1").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("address2").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("address3").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("booking_date").Value & "',"
            strQuery = strQuery & " " & Rs_Received_By.Fields("booking_month").Value & ","
            strQuery = strQuery & " '" & Rs_Received_By.Fields("material_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("cancel_no").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("cancel_date").Value & "',"
            strQuery = strQuery & " '" & IIf(IsNull(Rs_Received_By.Fields("replace_date").Value), "9/9/9999", Rs_Received_By.Fields("replace_date").Value) & "',"
            strQuery = strQuery & " " & Rs_Received_By.Fields("cancel_print_flag").Value & ","
            strQuery = strQuery & " '" & Rs_Received_By.Fields("paper_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("COLOR_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("Description").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("Brand_name").Value & "',"
            strQuery = strQuery & " '" & Rs_Received_By.Fields("client_name").Value & "'"
            
            For intIndex = 1 To IntPOCopy
                If rs_Get_Address("Client_Code").Value = "XLUM" Or rs_Get_Address("Client_Code").Value = "CMAN" Then
                    dfwConn.Execute strQuery & ",'01',0," & intIndex & ")"
                'ElseIf rs_Get_Address("Client_Code").Value = "CBCI" Or rs_Get_Address("Client_Code").Value = "CBC2" Then
                '    dfwConn.Execute strQuery & ",'05',0," & intIndex & ")"
                Else
                    dfwConn.Execute strQuery & ",'" & rs_Get_Address("business_unit").Value & "',0," & intIndex & ")"
                End If
            Next intIndex
            
            Rs_Received_By.MoveNext
        Wend
        Rs_Received_By.Close
        Set Rs_Received_By = Nothing
        
        rs_Get_Address.Close
    End If
    
    With Crpt
        .Reset
        Me.MousePointer = vbHourglass
        
        If Is_mmCl = True Then
            Posisi = InStr(1, Lbl_Size.Caption, "x")
            .Formulas(0) = "size='" & Mid(Lbl_Size.Caption, 1, Posisi - 1) & " col x " & Mid(Lbl_Size.Caption, Posisi + 1, Len(Lbl_Size.Caption)) & " mm" & "'"
            strSize = Mid(Lbl_Size.Caption, 1, Posisi - 1) & " col x " & Mid(Lbl_Size.Caption, Posisi + 1, Len(Lbl_Size.Caption)) & " mm"
        Else
            .Formulas(0) = "size='" & Lbl_Size.Caption & " x " & Lbl_Satuan.Caption & "'"
            strSize = Lbl_Size.Caption & " x " & Lbl_Satuan.Caption
        End If
        
        Sql = "SELECT Name FROM User_id "
        Sql = Sql & "WHERE user_name=(SELECT TOP 1 User_name FROM Media_Security_catalog "
        Sql = Sql & "WHERE UPPER(LTRIM(RTRIM(Position)))='SUPERVISOR' AND "
        Sql = Sql & " brand_code ='" & Left(Cbo_Job_No.Text, 4) & "' AND Valid_until >= getdate() ORDER BY Valid_Until DESC )"
        rs_Supervisor.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        
        Sql = "SELECT manager_po, HAL_PO FROM print_information"
        rs_PO_Manager.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
        
        If rs_Supervisor.BOF And rs_Supervisor.EOF Then
            .Formulas(1) = "tertanda='" & "( " & rs_PO_Manager(0) & " )" & "'"
            strSign = rs_PO_Manager(0)
        Else
            .Formulas(1) = "tertanda='" & "( " & rs_Supervisor(0) & " )" & "'"
            strSign = rs_Supervisor(0)
        End If
        
        rs_Supervisor.Close
        Set rs_Supervisor = Nothing

        .Formulas(2) = "hal_cancel='" & rs_PO_Manager(1) & "'"
        StrDescription = rs_PO_Manager(1)
        
        .Formulas(3) = "total_nett_cost='" & Lbl_Total_Nett.Caption & "'"
        
        rs_PO_Manager.Close
        Set rs_PO_Manager = Nothing
        
        If UCase(Trim(strPaperTypeCO)) = "A4" Then
            .ReportFileName = strReport_Dir & strPathPO & "\print\cO_Print_A4.RPT"
            '.ReportFileName = App.Path & "\report\print\co_Print_A4.RPT"
            If blnIsWinfax = True Then
                .SelectionFormula = "{COPrint.Copy_No}=1"
            End If
        Else
            If Destination_Print_PO = True Then
                .ReportFileName = strReport_Dir + "\print\cancel_order.rpt"
            Else
                .ReportFileName = strReport_Dir + "\print\CO_Print_Fax.rpt"
            End If
            Sql = "{po_print.cancel_no}='" & lbl_Cancel_No.Caption & "'"
            .SelectionFormula = Sql
        End If
        '.ReportFileName = App.Path + "\report\cancel_order.rpt"
        
        .WindowState = crptMaximized
        .WindowTitle = "Cancel Order " & Lbl_PO.Caption
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .Connect = "DSN=" & Server_Name & ";UID=" & Login_User & ";PWD=" & Login_Password & ";DSQ=" & Database_Name & ""
        .RetrieveDataFiles
        .Action = 1
        
                
    End With
    
    strPaperTypeCO = strPaperTypeOld
    
    Me.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    Dim Th As Integer
    
    RemoveMenus Me, True
    
    Load_Brand
    
    Pc_Browse.Visible = False
    
    Tombol False
    'isi tahun
        Me.Cbo_Year.Clear
    For Th = 2000 To 2016
        Me.Cbo_Year.AddItem Th
    Next
    
    recDate.Requery
    Me.Cbo_Year.Text = Year(recDate(0))
    
    'isi bulan
    Me.Cbo_Month.Clear
    For Th = 1 To 12
        Me.Cbo_Month.AddItem Format(Th, "00") & " - " & Get_Month_Name(Th)
    Next
    Me.Cbo_Month.Text = Format(month(recDate(0)), "00") & " - " & Get_Month_Name(month(recDate(0)))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs_Cancel_Print.State = adStateOpen Then
        rs_Cancel_Print.Close
        Set rs_Cancel_Print = Nothing
    End If
    
    If rs_Cancel_Temp.State = adStateOpen Then
        rs_Cancel_Temp.Close
        Set rs_Cancel_Temp = Nothing
    End If
End Sub

Private Sub Load_Note()
    Dim rs_Load_note As New ADODB.Recordset
    
    Sql = "SELECT note_code FROM print_catalog_note"
    rs_Load_note.Open Sql, ConnERP, adOpenStatic, adLockReadOnly
    
    Cbo_Note.Clear
    While Not rs_Load_note.EOF And Not rs_Load_note.BOF
        Cbo_Note.AddItem rs_Load_note(0)
        rs_Load_note.MoveNext
    Wend
    
    rs_Load_note.Close
    Set rs_Load_note = Nothing
End Sub
