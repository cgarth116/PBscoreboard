VERSION 5.00
Object = "{EA4C06C4-DD2F-41A9-AEF0-9FB0C8C9BAB9}#1.0#0"; "SComm32x.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PB_ScoreBoard"
   ClientHeight    =   7410
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10605
   DrawStyle       =   4  'Dash-Dot-Dot
   ForeColor       =   &H000000FF&
   Icon            =   "buzzer_1_5_4.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "buzzer_1_5_4.frx":0CCA
   ScaleHeight     =   364.1
   ScaleMode       =   0  'User
   ScaleWidth      =   477.285
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo14 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9240
      TabIndex        =   94
      Text            =   "Табло в пит"
      Top             =   4152
      Width           =   1215
   End
   Begin VB.ComboBox Combo13 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   93
      Text            =   "Выберите пульт старшего судьи поля"
      Top             =   5160
      Width           =   2415
   End
   Begin VB.ComboBox Combo12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      TabIndex        =   92
      Text            =   "Пит B"
      Top             =   4844
      Width           =   1215
   End
   Begin VB.ComboBox Combo11 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8040
      TabIndex        =   91
      Text            =   "Пит А"
      Top             =   4844
      Width           =   1215
   End
   Begin VB.ComboBox Combo10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      TabIndex        =   90
      Text            =   "База B"
      Top             =   4498
      Width           =   1215
   End
   Begin VB.CommandButton Command55 
      Caption         =   "-"
      Height          =   547
      Left            =   2400
      TabIndex        =   89
      Top             =   6179
      Width           =   375
   End
   Begin VB.CommandButton Command54 
      Caption         =   "+"
      Height          =   549
      Left            =   2400
      TabIndex        =   88
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   1080
      TabIndex        =   87
      Text            =   "Text17"
      Top             =   5010
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      TabIndex        =   85
      Text            =   "Табло в поле"
      Top             =   4152
      Width           =   1215
   End
   Begin SCommLib.SComm SComm2 
      Left            =   2640
      Top             =   6840
      _ExtentX        =   953
      _ExtentY        =   979
      CommPort        =   2
      Settings        =   "9600,N,8,1"
      OverlappedIO    =   0   'False
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Блокируем порт"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   5495
      Width           =   1215
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "buzzer_1_5_4.frx":711C
      Left            =   8040
      List            =   "buzzer_1_5_4.frx":711E
      TabIndex        =   83
      Text            =   "База А"
      Top             =   4498
      Width           =   1215
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Инфо табло на TV"
      Height          =   255
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Инфо табло турнира"
      Height          =   255
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   600
      TabIndex        =   80
      Text            =   "Text16"
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   840
      TabIndex        =   79
      Text            =   "Text15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command50 
      Caption         =   "окончание_матча_main"
      Height          =   375
      Left            =   6360
      TabIndex        =   78
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Overtime_mane"
      Height          =   375
      Left            =   5280
      TabIndex        =   77
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XBall                10 мин\5 win"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   346
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1245
      Width           =   2415
   End
   Begin VB.CommandButton Command47 
      Caption         =   "-"
      Height          =   195
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3645
      Width           =   255
   End
   Begin VB.CommandButton Command46 
      Caption         =   "+"
      Height          =   195
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5 мин"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "OverTime"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Тест кнопок"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5495
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   70
      Text            =   "Text4"
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Обратная связь"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7080
      Width           =   2377
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   68
      Text            =   "Команды на поле"
      Top             =   1080
      Width           =   1695
   End
   Begin SCommLib.SComm SComm1 
      Left            =   2040
      Top             =   6840
      _ExtentX        =   953
      _ExtentY        =   979
      InputLen        =   2
      RThreshold      =   3
      Settings        =   "115200,N,8,1"
      CommName        =   "Последовательный порт (COM1)"
      OverlappedIO    =   0   'False
   End
   Begin VB.CommandButton Command32 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10200
      TabIndex        =   63
      Top             =   1530
      Width           =   255
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Гейм без очка"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Overtime"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Время"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command33 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10200
      TabIndex        =   64
      Top             =   1665
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Перерыв"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Перерыв"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   549
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Пит-Стоп"
      Top             =   5088
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "Время Матча"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Звук"
      Top             =   2490
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Продолжительность Пит-Стопа"
      Top             =   1875
      Width           =   2400
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "Формат игр"
      Top             =   0
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1200
      TabIndex        =   47
      Text            =   "Text7"
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sound Woman"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   9240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sound Мan"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sound 10 sec."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   43
      Text            =   "Выберите порт подключения приёмника"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Command30"
      Height          =   255
      Left            =   1080
      TabIndex        =   42
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   255
      Left            =   1320
      TabIndex        =   41
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Command28"
      Height          =   255
      Left            =   840
      TabIndex        =   40
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Command27"
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   840
      TabIndex        =   36
      Text            =   "Combo6"
      Top             =   5520
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   480
      TabIndex        =   35
      Text            =   "Combo5"
      Top             =   5520
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Готовятся в пите"
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   5520
      TabIndex        =   33
      Text            =   "Combo4"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   360
      TabIndex        =   32
      Text            =   "Combo3"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Смена пары команд"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command25 
      Caption         =   "-"
      Height          =   549
      Left            =   5280
      TabIndex        =   30
      Top             =   6180
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      Caption         =   "+"
      Height          =   549
      Left            =   5280
      TabIndex        =   29
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Caption         =   "-"
      Height          =   735
      Left            =   6000
      TabIndex        =   28
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      Caption         =   "+"
      Height          =   855
      Left            =   6000
      TabIndex        =   27
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   204
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1038
      Width           =   513
   End
   Begin VB.CommandButton Command20 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   204
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   464
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1038
      Width           =   511
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BUZZER"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Добавить команду в список"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   305
      Left            =   7440
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   305
      Left            =   7440
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "-"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Гейм команды"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "-"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Гейм команды"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Сброс PitTime"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Старт Пит-Стоп"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "30 сек"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   811
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1 мин"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   811
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2 мин"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   366
      Left            =   9645
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   811
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "TimeOut"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Окончание матча"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Поле в игре"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   120
      Width           =   2160
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2040
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2 мин"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command39 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   60
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command40 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9000
      TabIndex        =   61
      Top             =   2895
      Width           =   255
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custom"
      Height          =   366
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XBallLight         10 мин\4 win"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   346
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   915
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XBallSuperLight 8 мин\3 win"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   346
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   585
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XBallUltraLight   6 мин\2 win"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   346
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   270
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   360
      TabIndex        =   86
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image5 
      Height          =   1095
      Left            =   8040
      Picture         =   "buzzer_1_5_4.frx":7120
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2400
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   66
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   1680
      TabIndex        =   25
      Top             =   2280
      Width           =   4395
   End
   Begin VB.Image Image2 
      Height          =   1470
      Left            =   5760
      Picture         =   "buzzer_1_5_4.frx":82CB
      Top             =   5460
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   360
      Picture         =   "buzzer_1_5_4.frx":A9B6
      Top             =   5460
      Width           =   1950
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3240
      TabIndex        =   57
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5040
      TabIndex        =   56
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2520
      TabIndex        =   55
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   480
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   45.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   2760
      TabIndex        =   26
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   6480
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   360
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'счетчик для переключателя "Старт-стоп"
'переменная для чтения сом порта

Dim t_1main As Single
Dim t_1pit As Single
Dim k1_board As String
Dim k1_score_board As String
Dim k2_board As String
Dim k2_score_board As String
Dim t1_board As String
Dim t2_board As String
Dim i As Long
Dim button_yes As Integer
Dim button_identification As String
Dim team_counter As Integer
Dim null_group(100) As String
Dim group_counter As Integer
Dim competition_info As String
Dim Space As String
Dim day_end As Integer
Dim period As String
Dim Qualifying_position_1 As Integer
Dim Qualifying_position_2 As Integer
Dim Qualifying As Integer
Dim g As Integer
Dim array_group(7) As String
Dim result_table_sort1 As Integer
Dim result_table_sort2 As Integer
Dim result_table_N As Integer
Dim result_table_array(200) As result_table
Dim win_los As Integer
Dim result_table_index As Integer
Dim result_table_position As Integer
Dim result_info As String
Dim result_info_2 As String
Dim result_info_counter As Integer
Dim overtime_info As String
Dim color_active As Double
Dim color_unactive As Double
Dim color_active_timeout As Double
Dim color_unactive_timeout As Double
Dim score_end As Integer
Dim license_counter As Integer
Dim license As Integer
Dim license_info As String
Dim Password As String
Dim Password_q As Long
Dim MyText As String
Dim k As String
Dim end_game As Integer
Dim second_counter As Integer
Dim trackList(17) As String
Dim trackNumber As Integer
Dim match_point As Integer
Dim Buffer() As Byte
Dim sc As String
Dim a As Integer
Dim c As Integer
Dim a_break As Integer
Dim a1 As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim overtime_time As Integer
Dim overtime As Integer
Dim overtime1 As Integer
Dim overtime_save As Integer
'счетчик очков команд
Dim team_1 As Integer
Dim team_2 As Integer
'счетчик времени отображаемого в данный момент в основном времени
Dim Finish As Single
Dim finish_custom As Single
Dim finish_break_custom As Single
Dim finish_break_custom1 As Single
Dim t1 As Single
'счетчик времени для расчетов основного времени
Dim Finish_main As Single
'счетчик времени отображаемого в данный момент в break
Dim Finish_break As Single
'счетчик времени для расчетов в break
Dim Finish_main_break As Single
Dim a_2 As Integer
Dim a_break_2 As Integer
Dim a1_2 As Integer
'счетчик очков команд
Dim team_1_2 As Integer
Dim team_2_2 As Integer
'счетчик времени отображаемого в данный момент в основном времени
Dim Finish_2 As Single
'счетчик времени для расчетов основного времени
Dim Finish_main_2 As Single
'счетчик времени отображаемого в данный момент в break
Dim Finish_break_2 As Single
'счетчик времени для расчетов в break
Dim Finish_main_break_2 As Single
Dim Start As Single
'глобальная переменная для запуска цикла секундомера
Dim Switch As Boolean
Dim Combo1_save As String
Dim Combo2_save As String
Dim label1_save As Integer
Dim label2_save As Integer
Dim label1_1_save As Integer
Dim label2_1_save As Integer
Dim command1_save As String
Dim command2_save As String
Dim command1_1_save As String
Dim command1_2_save As String
Dim label3_save As Single
Dim label3_1_save As Single
Dim Finish_save As Single
Dim Finish1_save As Single
Dim Finish_main1_save As Single
Dim Finish_main_save As Single
Dim PushButton As Long
'подключаем звуковые файлы
'Private Declare Function sndPlaySound Lib "winmm.dll" Alias _
'      "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
'      Long) As Long
   
Dim dx As New DirectX8 'DirectX
Dim ds As DirectSound8 'DirectSound
   
Dim dsBuffer1 As DirectSoundSecondaryBuffer8
Dim dsBuffer2 As DirectSoundSecondaryBuffer8
Dim dsBuffer3 As DirectSoundSecondaryBuffer8
Dim dsBuffer4 As DirectSoundSecondaryBuffer8
Dim dsBuffer5 As DirectSoundSecondaryBuffer8
Dim dsBuffer6 As DirectSoundSecondaryBuffer8
Dim dsBuffer7 As DirectSoundSecondaryBuffer8
Dim dsBuffer8 As DirectSoundSecondaryBuffer8
Dim dsBuffer9 As DirectSoundSecondaryBuffer8
Dim dsBuffer10 As DirectSoundSecondaryBuffer8
Dim dsBuffer11 As DirectSoundSecondaryBuffer8
Dim dsBuffer12 As DirectSoundSecondaryBuffer8
Dim dsBuffer13 As DirectSoundSecondaryBuffer8
Dim dsBuffer14 As DirectSoundSecondaryBuffer8
Dim dsBuffer15 As DirectSoundSecondaryBuffer8
Dim dsBuffer16 As DirectSoundSecondaryBuffer8
Dim Bufferdesc As DSBUFFERDESC

'Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
   
Private Type result_table
    Command_name As String
    result_score_match As Integer
    result_win_lose As Integer
    result_by_game As String
    result_group As String
    counter_match As Integer
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Sub Sleep(millisecondsTimeout As Integer)
'End Sub
'End Sub

'End Sub

'убиваем при выходе чтобы не exe не висел в памяти
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

'звук гудка, лежит в папке с файлом
Private Sub Sound_buzzer()
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer1 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\buzzer.wav", Bufferdesc)
dsBuffer1.Play DSBPLAY_DEFAULT
End Sub


Private Sub Sound_buzzer_stop()
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer1 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\buzzer.wav", Bufferdesc)
dsBuffer1.Play DSBPLAY_DEFAULT
End Sub

'звук секунда, лежит в папке с файлом
Private Sub Sound_seconds()
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer2 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\beep_seconds.wav", Bufferdesc)
dsBuffer2.Play DSBPLAY_DEFAULT
End Sub

'звук 10, лежит в папке с файлом
Private Sub Sound_10()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer3 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\10_woman.wav", Bufferdesc)
dsBuffer3.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer4 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\10_man.wav", Bufferdesc)
dsBuffer4.Play DSBPLAY_DEFAULT
End If
End Sub

'звук 30, лежит в папке с файлом
Private Sub Sound_30()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer5 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\30_woman.wav", Bufferdesc)
dsBuffer5.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer6 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\30_man.wav", Bufferdesc)
dsBuffer6.Play DSBPLAY_DEFAULT
End If
End Sub

'звук минута, лежит в папке с файлом
Private Sub Sound_minute()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer7 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\minute_woman.wav", Bufferdesc)
dsBuffer7.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer8 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\minute_man.wav", Bufferdesc)
dsBuffer8.Play DSBPLAY_DEFAULT
End If
End Sub

'звук 60, лежит в папке с файлом
Private Sub Sound_60()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer9 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\60_woman.wav", Bufferdesc)
dsBuffer9.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer10 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\60_man.wav", Bufferdesc)
dsBuffer10.Play DSBPLAY_DEFAULT
End If
End Sub

'звук dtmw, лежит в папке с файлом
Private Sub Sound_dtmw()
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer11 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\dtmw.wav", Bufferdesc)
dsBuffer11.Play DSBPLAY_DEFAULT
End Sub

'звук time, лежит в папке с файлом
Private Sub Sound_time()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer12 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\time_woman.wav", Bufferdesc)
dsBuffer12.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer13 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\time_man.wav", Bufferdesc)
dsBuffer13.Play DSBPLAY_DEFAULT
End If
End Sub

'звук field in game, лежит в папке с файлом
Private Sub Sound_field_in_game()
If Command36.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer14 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\field_in_game_woman.wav", Bufferdesc)
dsBuffer14.Play DSBPLAY_DEFAULT
End If
If Command35.BackColor = color_active Then
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer15 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\field_in_game_man.wav", Bufferdesc)
dsBuffer15.Play DSBPLAY_DEFAULT
End If
End Sub

'звук dtmw, лежит в папке с файлом
Private Sub Sound_dtmw_stop()
Set ds = dx.DirectSoundCreate(vbNullString)
ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
Set dsBuffer1 = ds.CreateSoundBufferFromFile(App.Path + "\sounds\dtmw_stop.wav", Bufferdesc)
dsBuffer1.Play DSBPLAY_DEFAULT
End Sub

Private Sub Combo10_Click()
If Combo10.Text = Combo8.Text Then
    MsgBox ("Указана одна и та же кнопка для обоих баз.")
    Combo10.Text = "База В"
End If
End Sub

Private Sub Combo11_Click()
If Combo11.Text = Combo12.Text Then
    MsgBox ("Указана одна и та же кнопка для обоих питов.")
    Combo11.Text = "Пит А"
End If
End Sub

Private Sub Combo12_Click()
If Combo12.Text = Combo11.Text Then
    MsgBox ("Указана одна и та же кнопка для обоих питов.")
    Combo12.Text = "База В"
End If
End Sub

Private Sub Combo3_Click()
Command16.Caption = "Гейм команды" & vbCrLf & Combo3.Text
Command7.Caption = "ТimeOut " & vbCrLf & Combo3.Text
Form3.Text1.Text = Combo3.Text
End Sub

Private Sub Combo4_Click()
Command14.Caption = "Гейм команды" & vbCrLf & Combo4.Text
Command6.Caption = "ТimeOut " & vbCrLf & Combo4.Text
Form3.Text2.Text = Combo4.Text
End Sub

Private Sub Combo8_Click()
If Combo8.Text = Combo10.Text Then
    MsgBox ("Указана одна и та же кнопка для обоих баз.")
    Combo8.Text = "База А"
End If
End Sub

Private Sub Command11_Click()
If Command11.Caption = "Поле в игре" Then
Call Sound_dtmw
Call Sound_field_in_game
'прячем кнопку окончание матча
Command5.Visible = False
'скрываем кнопку продолжить гейм
Command11.Visible = False
'скрываем кнопку овертайм
If overtime = 0 Then
Command41.Visible = False
End If
Call Command4_Click
End If
End Sub

'вводим название команд
Private Sub Command18_Click()
If Text1.Text = "" Then
Else
result_table_N = result_table_N + 1
Combo1.AddItem Text1
Combo2.AddItem Text1
Combo3.AddItem Text1
Combo4.AddItem Text1
List1.AddItem Text1.Text
Text1.Text = ""
End If
End Sub

'изменяем вручную секунды в основном времени +1
Private Sub Command22_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label3.Caption >= "0.00,00" Then
    Finish = Finish - 1
    If Finish > 0 Then
    Finish = 0
    End If
    Label3.Caption = TimeInLabel(-Finish)
    Form3.Label3.Caption = TimeInLabel(-Finish)
    End If
End If
End If
End Sub

'изменяем вручную секунды в основном времени -1
Private Sub Command23_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label3.Caption >= "0.00,00" Then
    Finish = Finish + 1
    If Finish > 0 Then
    Finish = 0
    End If
    Label3.Caption = TimeInLabel(-Finish)
    Form3.Label3.Caption = TimeInLabel(-Finish)
    End If
End If
End If
End Sub

'изменяем вручную секунды в перерыве +1
Private Sub Command24_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label4.Caption >= "0.00,00" Then
    Finish_break = Finish_break - 1
    If Finish_break > 0 Then
    Finish_break = 0
    End If
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
    End If
End If
End If
End Sub

'изменяем вручную секунды в перерыве -1
Private Sub Command25_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label4.Caption >= "0.00,00" Then
    Finish_break = Finish_break + 1
    If Finish_break > 0 Then
    Finish_break = 0
    End If
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
    End If
End If
End If
End Sub

'сохраняем данные о матче
Private Sub Command26_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
'сохраняем название играющих команд
Combo5.Text = Combo3.Text
Combo6.Text = Combo4.Text
'сохраняем счет
label1_1_save = team_1
label2_1_save = team_2
'сохраняем овертайм
overtime1 = overtime
'Инфо табло для ожидающих команд
Label7.Caption = team_1
Label8.Caption = team_2
'Инфо табло для ожидающих команд
If i1 <> 0 Then
    Label9.Caption = "матч окончен"
Else
    Label9.Caption = Label3.Caption
End If
'сохраняем время
'label 5 и 6 временные переменные

Label5.Caption = Label3.Caption
Finish1_save = Finish
Finish_main1_save = Finish_main
'сохраняем таймауты
Command28.BackColor = Command7.BackColor
Command30.BackColor = Command6.BackColor
'меняем
'меняем местами команды
Combo3.Text = Combo1.Text
Form3.Text1.Text = Combo3.Text
Combo4.Text = Combo2.Text
Form3.Text2.Text = Combo4.Text
Combo1.Text = Combo5.Text
Combo2.Text = Combo6.Text
Label10.Caption = Board_team(t_1main, t_1pit)
'меняем счет
Label1.Caption = label1_save
Form3.Label1.Caption = label1_save
team_1 = label1_save
Label2.Caption = label2_save
Form3.Label2.Caption = label2_save
team_2 = label2_save
label1_save = label1_1_save
label2_save = label2_1_save
'меняем овертаймы
overtime = overtime_save
overtime_save = overtime1
'меняем таймауты
Command7.BackColor = Command27.BackColor
Form3.Command1.BackColor = Command7.BackColor
Command27.BackColor = Command28.BackColor
Command6.BackColor = Command29.BackColor
Form3.Command2.BackColor = Command6.BackColor
Command29.BackColor = Command30.BackColor
'меняем время матча
Label3.Caption = Label6.Caption
Form3.Label3.Caption = Label6.Caption
Label6.Caption = Label5.Caption
Finish = Finish_save
't_1main = -Finish
Finish_save = Finish1_save
Finish_main = Finish_main_save
Finish_main_save = Finish_main1_save
'убираем кнопку overtime
If overtime = 0 Then
    Command41.Visible = False
Else
    Command41.Visible = True
    Command41.Caption = "Идет Overtime"
    Command41.BackColor = &HFF&
End If
'выводим название в timeout и гейм
Command16.Caption = "Гейм команды " & vbCrLf & Combo3.Text
Command7.Caption = "ТimeOut " & vbCrLf & Combo3.Text
Command14.Caption = "Гейм команды " & vbCrLf & Combo4.Text
Command6.Caption = "ТimeOut " & vbCrLf & Combo4.Text
'зеленим кнопки команд
Command20.BackColor = &HFF00&
Command21.BackColor = &HFF00&
Finish_break = Finish_main_break
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
    Label4.Caption = TimeInLabel(-finish_break_custom)
    Form3.Label4.Caption = TimeInLabel(-finish_break_custom)
Else
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If

Label10.Caption = Board_team(t_1main, t_1pit)
'Sleep (300)
'Label10.Caption = Board_team(t_1main, t_1pit)

If result_info_counter = 0 Then
    result_info_counter = 1
    Else
    result_info_counter = 0
End If
End If
End If
End Sub

Private Sub Command34_Click()
If Command34.BackColor = color_unactive Then
Command34.BackColor = color_active
Else
Command34.BackColor = color_unactive
End If
End Sub

Private Sub Command35_Click()
If Command35.BackColor = color_unactive Then
Command35.BackColor = color_active
Command36.BackColor = color_unactive
End If
End Sub

Private Sub Command36_Click()
If Command36.BackColor = color_unactive Then
Command36.BackColor = color_active
Command35.BackColor = color_unactive
End If
End Sub

'кнопка перерыва между матчами
Private Sub Command37_Click()
' проверяем что матчи еще не начаты
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
Command5.Caption = "Окончание матча"
Label4.Caption = TimeInLabel(-finish_break_custom)
Form3.Label4.Caption = TimeInLabel(-finish_break_custom)
Finish_break = finish_break_custom
Call Command12_Click
End If
End If
End If
End Sub

Private Sub Command38_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If a_break = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command38.BackColor = color_active
Finish_break = finish_break_custom
Finish_main_break = finish_break_custom

If finish_break_custom Mod 60 = 0 Then
    Command38.Caption = -finish_break_custom \ 60 & " мин"
Else
    Command38.Caption = -finish_break_custom \ 60 & " мин " & "30 сек"
End If

Label4.Caption = TimeInLabel(-finish_break_custom)
Form3.Label4.Caption = TimeInLabel(-finish_break_custom)
End If
End If
End If
End If
End Sub

Private Sub Command39_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command38.BackColor = color_active
finish_break_custom = finish_break_custom - 30
Finish_main_break = finish_break_custom
Finish_break = finish_break_custom

If finish_break_custom Mod 60 = 0 Then
    Command38.Caption = -finish_break_custom \ 60 & " мин"
Else
    Command38.Caption = -finish_break_custom \ 60 & " мин " & "30 сек"
End If

Label4.Caption = TimeInLabel(-finish_break_custom)
Form3.Label4.Caption = TimeInLabel(-finish_break_custom)
End If
End If
End If
End Sub

Private Sub Command40_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If finish_break_custom < -30 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command38.BackColor = color_active
finish_break_custom = finish_break_custom + 30
Finish_main_break = finish_break_custom
Finish_break = finish_break_custom

If finish_break_custom Mod 60 = 0 Then
    Command38.Caption = -finish_break_custom \ 60 & " мин"
Else
    Command38.Caption = -finish_break_custom \ 60 & " мин " & "30 сек"
End If

Label4.Caption = TimeInLabel(-finish_break_custom)
Form3.Label4.Caption = TimeInLabel(-finish_break_custom)
End If
End If
End If
End If
End Sub

Private Sub Command41_Click()
'If overtime = 0 Then
    If a_break Mod 2 = 0 Then
        If a Mod 2 = 0 Then
            PushButton = MsgBox("Подтверждаете что команды переходят в Overtime?", 292)
            If PushButton = 6 Then
                If team_1 > team_2 Then
                     team_2 = team_1
                     Label2.Caption = Label1.Caption
                    Else
                     team_1 = team_2
                     Label1.Caption = Label2.Caption
                End If
                     'Label1.Caption = Label2.Caption
                     Form3.Label1.Caption = Label1.Caption
                     Form3.Label2.Caption = Label2.Caption
                        If result_info_counter = 0 Then
                            'если первая пара играет
                            result_info_2 = result_info_2 & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf & "Overtime" & vbCrLf
                            Else
                            'если вторая пара играет
                            result_info = result_info & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf & "Overtime" & vbCrLf

                        End If
                Call Command49_Click
                Call Command12_Click
            End If
        End If
    End If
'End If
End Sub

Private Sub Command42_Click()
    'скрываем кнопку овертайм
    If overtime = 0 Then
    Command41.Visible = False
    End If
    If a_break Mod 2 = 0 Then
        If a Mod 2 = 0 Then
            If Command4.Caption = "Поле в игре" Then
                If Finish1_save <> Finish_main1_save Or Finish <> Finish_main Then
                    'Call Sound_dtmw_stop
                    
                    If result_info_counter = 0 Then
                        'если первая пара играет
                        result_info_2 = result_info_2 & Combo3.Text & " | " & Label1.Caption & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                        Else
                        'если вторая пара играет
                        result_info = result_info & Combo3.Text & " | " & Label1.Caption & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                    End If
                    
                    'меняем признак играющих команд
                    'если 2 пита, то не меням, если 4 то меняем
                    'If Combo1.Text = "" And Combo2.Text = "" Then
                    '    Else
                    'End If
                    
                    Command4.Caption = "Старт гейма"
                    Command11.Caption = "Старт гейма"
                    If Command10.BackColor = color_active Then
                    Finish_break = -30
                    Finish_main_break = -30
                    End If
                    If Command9.BackColor = color_active Then
                    Finish_break = -60
                    Finish_main_break = -60
                    End If
                    If Command8.BackColor = color_active Then
                    Finish_break = -120
                    Finish_main_break = -120
                    End If
                    
                    'скрываем кнопку гейм без очка
                    Command42.Visible = False
                    'скрываем кнопку время
                    Command11.Visible = False
                    If Label3.Caption <> "0.00,00" Then
                        If i1 = 0 Then 'если 4 пита
                            Call Command26_Click 'меняем команды
                        Else 'если 2 пита
                            Finish_break = -120
                            Finish_main_break = -120
                        End If
                        Call Command12_Click 'запускаем пит
                    Else
                        If overtime <> 0 Then
                            If team_1 = team_2 Then
                                Call Command50_Click
                                If i1 = 1 Then 'если 4 пита то запускаем таймер
                                Call Command12_Click
                                End If
                            End If
                        Else
                            If team_1 = team_2 Then
                                Call Command49_Click 'овертайм
                                Call Command12_Click 'запускаем пит
                            Else
                                Call Command50_Click
                                If i1 = 1 Then 'если 4 пита то запускаем таймер
                                Call Command12_Click
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Command43_Click()
If license = 1 Then
license_info = "Программа активирована"
Else
license_info = "Демо режим"
End If
PushButton = MsgBox("PaintBall ScoreBoard v.1.6.1" & vbCrLf & "Лицензия :               " & license_info & vbCrLf & "Разработчик ПО:    Култашев Виктор" & vbCrLf & "Аппаратная часть: Фома Хомячов" & vbCrLf & "Тестирование:        Зарецкий Дмитрий" & vbCrLf & "e-mail:                       fion_alp@mail.ru" & vbCrLf & "tel:                             +79600380828" & vbCrLf & "SSID:                          XXXX", 65 - license)
If PushButton = 2 Then
'формируем код запроса
Password_q = Int(Timer * 1000)
'формируем правильный ответ
Text4.Text = ((Int(Password_q \ 10000) * Int(Password_q \ 10000)) + Int(Password_q \ 7) + 1051981)
Password = InputBox("Полная версия продукта доступна только по паролю." & vbCrLf & "код запроса : " & Password_q & "VX" & Int(Timer * 0.03))
If Password = Text4.Text Then
    license = 1
'Else
'    license = 0
End If
End If
End Sub


Private Sub Command44_Click()
If Command44.BackColor = color_unactive Then
Command44.BackColor = &HFF&
Else
Command44.BackColor = color_unactive
End If
End Sub

Private Sub Command45_Click()
'If a_break Mod 2 = 0 Then
'If a Mod 2 = 0 Then
'If a_break = 0 Then
'If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command45.BackColor = color_active
Command45.Caption = -overtime_time \ 60 & " мин"
'End If
'End If
'End If
'End If
End Sub

Private Sub Command46_Click()
'даем возможность менять время овертайма в любое время
'If a_break Mod 2 = 0 Then
'If a Mod 2 = 0 Then
'If a_break = 0 Then
'If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command45.BackColor = color_active
overtime_time = overtime_time - 60
Command45.Caption = -overtime_time \ 60 & " мин"
'End If
'End If
'End If
'End If
End Sub

Private Sub Command47_Click()
'даем возможность менять время овертайма в любое время
'If a_break Mod 2 = 0 Then
'If a Mod 2 = 0 Then
'If a_break = 0 Then
If overtime_time < -60 Then
'If Finish1_save = Finish_main1_save And Finish = Finish_main Then
Command45.BackColor = color_active
overtime_time = overtime_time + 60
Command45.Caption = -overtime_time \ 60 & " мин"
Command41.Caption = "Overtime"
'End If
'End If
'End If
'End If
End If
End Sub

Private Sub Command48_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
'If Command4.Caption = "Поле в игре" Then
'Else
If a_break = 0 Then
Finish = -600
Finish_main = -600
Finish_save = -600
Finish_main_save = -600
Command48.BackColor = color_active
Command1.BackColor = color_unactive
Command2.BackColor = color_unactive
Command3.BackColor = color_unactive
Command31.BackColor = color_unactive
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-Finish)
'счет окончания матча
score_end = 5
End If
End If
End If
End If
End Sub

Private Sub Command49_Click()
If Combo1.Text = "" And Combo2.Text = "" Then
   i2 = 1
End If
'если одна пара закончила играть
If i1 = 1 Then
   i2 = 1
End If
'скрываем кнопку гейм без очка
Command42.Visible = False
'скрываем кнопку время
Command11.Visible = False

If i2 > 0 Then
                'если 2 пита
                'перерыв пита 2 минуты
                'время матча овертайм
                'если овертаймов не было нажато до этого
                Finish = overtime_time
                Finish_main = Finish
                Finish_break = -120
                Label3.Caption = TimeInLabel(-Finish)
                Form3.Label3.Caption = TimeInLabel(-Finish)
                Label6.Caption = TimeInLabel(-Finish)
                Label4.Caption = TimeInLabel(-Finish_break)
                Form3.Label4.Caption = TimeInLabel(-Finish_break)
                overtime = overtime + 1
                 If team_1 > team_2 Then
                 team_2 = team_1
                 Else
                 team_1 = team_2
                 End If
                Label1.Caption = team_1
                Form3.Label1.Caption = team_1
                Label2.Caption = team_2
                Form3.Label2.Caption = team_2
                Command41.Caption = "Идет Overtime"
                Command41.BackColor = &HFF&
                Command11.Caption = "Время"
                
            Else
            
            'если 4 пита
            Finish = overtime_time
            Finish_main = Finish
            Label3.Caption = TimeInLabel(-Finish)
            Form3.Label3.Caption = TimeInLabel(-Finish)
            Command4.Caption = "Старт гейма"
            overtime = overtime + 1
            If team_1 > team_2 Then
                 team_2 = team_1
                 Else
                 team_1 = team_2
            End If
                Label1.Caption = team_1
                Form3.Label1.Caption = team_1
                Label2.Caption = team_2
                Form3.Label2.Caption = team_2
            Call Command26_Click
            Command11.Caption = "Время"
            
End If
End Sub

Private Sub Command50_Click()
'кнопки гейма красим в серый на всякий случай
Command16.BackColor = &HE0E0E0
Command14.BackColor = &HE0E0E0

'    If Combo1.Text = "" And Combo2.Text = "" Then
'        i1 = i1 + 1
'    End If

'скрываем кнопку гейм без очка
Command42.Visible = False

i1 = i1 + 1

'4 пита
    If i1 = 1 Then
                    
            'записываем результаты матча в файл
            Open App.Path & "\results.txt" For Append As #2
            Print #2, vbCrLf & Time & vbCrLf & "Результаты матча:" & vbCrLf;
            'если 2 пита то пишем только один результат
            If overtime <> 0 Then
               overtime_info = "Матч завершился в OverTime."
               Else
               overtime_info = "Матч завершился в основное время."
            End If
            
            'сохраняем в массив результаты матча
                                'название 1й команды
                                'Open App.Path & "\result_table.txt" For Append As #3
                                result_table_position = 0
                                
                                'проверить While result_table_array(result_table_position).Command_name <> "" вместо нижней строки
                                
                                While result_table_position <= 50
                                    'ищем запись по данной команде либо пустю строку
                                    If result_table_array(result_table_position).Command_name = Combo3.Text Or result_table_array(result_table_position).Command_name = "" Then
                                            'записываем название команды
                                            result_table_array(result_table_position).Command_name = Combo3.Text
                                            'узнаем кол-во оставшихся матчей для команды в отборах
                                            result_table_array(result_table_position).counter_match = result_table_array(result_table_position).counter_match - 1
                                            If Label1.Caption > Label2.Caption Then
                                               'если выиграли то +1
                                                win_lose = 1
                                                Else
                                                'если проиграли то 0
                                                win_lose = 0
                                            End If
                                            Space = Chr(32)
                                            If (team_1 - team_2) >= 0 Then
                                                Space = "+"
                                            End If
                                            'записываем результат по матчам
                                            result_table_array(result_table_position).result_by_game = result_table_array(result_table_position).result_by_game & " " & win_lose & "/" & Space & (team_1 - team_2) & " |"
                                            'записываем итоговую разницу геймов
                                            result_table_array(result_table_position).result_score_match = result_table_array(result_table_position).result_score_match + (team_1 - team_2)
                                            'записываем итог по выигранным матчам
                                            result_table_array(result_table_position).result_win_lose = result_table_array(result_table_position).result_win_lose + win_lose
                                            'Print #3, result_table_array(result_table_position).Command_name & " - " & result_table_array(result_table_position).result_by_game & result_table_array(result_table_position).result_win_lose & "/" & result_table_array(result_table_position).result_score_match & vbCrLf;
                                            'запоминаем позицию первой команды
                                            Qualifying_position_1 = result_table_position
                                            'выходим из цикла
                                            result_table_position = 51
                                        Else
                                    End If
                                    result_table_position = result_table_position + 1
                                Wend
                                
                                'название 2й команды
                                result_table_position = 0
                                
                                'проверить While result_table_array(result_table_position).Command_name <> "" вместо нижней строки
                                
                                While result_table_position <= 50
                                    'ищем запись по данной команде либо пустю строку
                                    If result_table_array(result_table_position).Command_name = Combo4.Text Or result_table_array(result_table_position).Command_name = "" Then
                                            'записываем название команды
                                            result_table_array(result_table_position).Command_name = Combo4.Text
                                            'узнаем кол-во оставшихся матчей для команды в отборах
                                            result_table_array(result_table_position).counter_match = result_table_array(result_table_position).counter_match - 1
                                            If Label1.Caption < Label2.Caption Then
                                               'если выиграли то +1
                                                win_lose = 1
                                                Else
                                                'если проиграли то 0
                                                win_lose = 0
                                            End If
                                            Space = Chr(32)
                                            If (team_2 - team_1) >= 0 Then
                                                Space = "+"
                                            End If
                                            'записываем результат по матчам
                                            result_table_array(result_table_position).result_by_game = result_table_array(result_table_position).result_by_game & " " & win_lose & "/" & Space & (team_2 - team_1) & " |"
                                            'записываем итоговую разницу геймов
                                            result_table_array(result_table_position).result_score_match = result_table_array(result_table_position).result_score_match + (team_2 - team_1)
                                            'записываем итог по выигранным матчам
                                            result_table_array(result_table_position).result_win_lose = result_table_array(result_table_position).result_win_lose + win_lose
                                            'Print #3, result_table_array(result_table_position).Command_name & " - " & result_table_array(result_table_position).result_by_game & result_table_array(result_table_position).result_win_lose & "/" & result_table_array(result_table_position).result_score_match & vbCrLf;
                                            'запоминаем позицию второй команды
                                            Qualifying_position_2 = result_table_position
                                            'выходим из цикла
                                            result_table_position = 51
                                        Else
                                    End If
                                    result_table_position = result_table_position + 1
                                Wend
                                'Close #3
            'If Combo1.Text <> "" Then
               ' Print #2, Combo3.Text & " " & team_1 & " - " & team_2 & " " & Combo4.Text & vbCrLf;
                            If result_info_counter = 0 Then
                                Print #2, result_info_2 & overtime_info & "Оставшееся время : " & Label3.Caption & vbCrLf;
                                Else
                                Print #2, result_info & overtime_info & "Оставшееся время : " & Label3.Caption & vbCrLf;
                            End If
                'Print #2, "Ход матча по геймам:" & vbCrLf & result_info & vbCrLf & team_1 & " - " & team_2 & vbCrLf;
             'End If
            Print #2, vbCrLf;
            Close #2
            
            If Qualifying = 1 Then
              'если отборы закончились то дописываем результаты в файл
              Open App.Path & "\result_table.txt" For Append As #3
                If result_table_array(Qualifying_position_1).result_win_lose > result_table_array(Qualifying_position_2).result_win_lose Then
                    Space = Chr(32)
                    If result_table_array(Qualifying_position_1).result_score_match >= 0 Then
                        Space = "+"
                    End If
                    Print #3, result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf;
                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf
                    Space = Chr(32)
                    If result_table_array(Qualifying_position_2).result_score_match >= 0 Then
                        Space = "+"
                    End If
                    Print #3, result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf & vbCrLf;
                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf & vbCrLf
                Else
                    Space = Chr(32)
                    If result_table_array(Qualifying_position_2).result_score_match >= 0 Then
                        Space = "+"
                    End If
                    Print #3, result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf;
                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf
                    Space = Chr(32)
                    If result_table_array(Qualifying_position_1).result_score_match >= 0 Then
                        Space = "+"
                    End If
                    Print #3, result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf & vbCrLf;
                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf & vbCrLf
                End If
                Close #3
            End If
            
            
            'обнуляем показание секундомера
            If Command10.BackColor = color_active Then
                Finish_break = -30
                Finish_main_break = -30
            End If
            
            If Command9.BackColor = color_active Then
                Finish_break = -60
                Finish_main_break = -60
            End If
            
            If Command8.BackColor = color_active Then
                Finish_break = -120
                Finish_main_break = -120
            End If
            
            If Command38.BackColor = color_active Then
                Label9.Caption = TimeInLabel(-finish_custom)
            End If
            
            Finish = Finish_main
            Label3.Caption = TimeInLabel(-Finish)
            Form3.Label3.Caption = TimeInLabel(-Finish)
            Label4.Caption = TimeInLabel(-Finish_break)
            Form3.Label4.Caption = TimeInLabel(-Finish_break)
            'обнуляем кнопки таймаута команд
            Command7.BackColor = color_unactive_timeout
            Form3.Command1.BackColor = Command7.BackColor
            Command6.BackColor = color_unactive_timeout
            Form3.Command2.BackColor = Command6.BackColor
            Command41.Caption = "Overtime"
            Command41.BackColor = color_unactive
            Command41.Visible = False

            'зеленим кнопки команд
            Command20.BackColor = &HFF00&
            Command21.BackColor = &HFF00&
            Command4.Caption = "Время"
            'скрываем кнопку продолжить гейм
            Command11.Visible = False
            'меняем команды местами
            Call Command26_Click
            Command11.Caption = "Время"
    End If
    
    'обнуляем если два окончания матча
    If i1 = 2 Then
            
            Command12.Visible = False
            If overtime <> 0 Then
               overtime_info = "Матч завершился в OverTime."
               Else
               overtime_info = "Матч завершился в основное время."
            End If
            
                                'сохраняем в массив результаты матча
                                'название 1й команды
                                result_table_position = 0
                                
                                'проверить While result_table_array(result_table_position).Command_name <> "" вместо нижней строки
                                
                                While result_table_position <= 50
                                    'ищем запись по данной команде либо пустю строку
                                    If result_table_array(result_table_position).Command_name = Combo3.Text Or result_table_array(result_table_position).Command_name = "" Then
                                            'записываем название команды
                                            result_table_array(result_table_position).Command_name = Combo3.Text
                                            'узнаем кол-во оставшихся матчей для команды в отборах
                                            result_table_array(result_table_position).counter_match = result_table_array(result_table_position).counter_match - 1
                                            If Label1.Caption > Label2.Caption Then
                                               'если выиграли то +1
                                                win_lose = 1
                                                Else
                                                'если проиграли то 0
                                                win_lose = 0
                                            End If
                                            Space = Chr(32)
                                            If (team_1 - team_2) >= 0 Then
                                                Space = "+"
                                            End If
                                            'записываем результат по матчам
                                            result_table_array(result_table_position).result_by_game = result_table_array(result_table_position).result_by_game & " " & win_lose & "/" & Space & (team_1 - team_2) & " |"
                                            'записываем итоговую разницу геймов
                                            result_table_array(result_table_position).result_score_match = result_table_array(result_table_position).result_score_match + (team_1 - team_2)
                                            'записываем итог по выигранным матчам
                                            result_table_array(result_table_position).result_win_lose = result_table_array(result_table_position).result_win_lose + win_lose
                                            'Print #3, result_table_array(result_table_position).Command_name & " - " & result_table_array(result_table_position).result_by_game & result_table_array(result_table_position).result_win_lose & "/" & result_table_array(result_table_position).result_score_match & vbCrLf;
                                            'запоминаем позицию первой команды
                                            Qualifying_position_1 = result_table_position
                                            'выходим из цикла
                                            result_table_position = 51
                                        Else
                                    End If
                                    result_table_position = result_table_position + 1
                                Wend
                                
                                'название 2й команды
                                result_table_position = 0
                                
                                'проверить While result_table_array(result_table_position).Command_name <> "" вместо нижней строки
                                
                                While result_table_position <= 50
                                    'ищем запись по данной команде либо пустю строку
                                    If result_table_array(result_table_position).Command_name = Combo4.Text Or result_table_array(result_table_position).Command_name = "" Then
                                            'записываем название команды
                                            result_table_array(result_table_position).Command_name = Combo4.Text
                                            'узнаем кол-во оставшихся матчей для команды в отборах
                                            result_table_array(result_table_position).counter_match = result_table_array(result_table_position).counter_match - 1
                                            If Label1.Caption < Label2.Caption Then
                                               'если выиграли то +1
                                                win_lose = 1
                                                Else
                                                'если проиграли то 0
                                                win_lose = 0
                                            End If
                                            Space = Chr(32)
                                            If (team_2 - team_1) >= 0 Then
                                                Space = "+"
                                            End If
                                            'записываем результат по матчам
                                            result_table_array(result_table_position).result_by_game = result_table_array(result_table_position).result_by_game & " " & win_lose & "/" & Space & (team_2 - team_1) & " |"
                                            'записываем итоговую разницу геймов
                                            result_table_array(result_table_position).result_score_match = result_table_array(result_table_position).result_score_match + (team_2 - team_1)
                                            'записываем итог по выигранным матчам
                                            result_table_array(result_table_position).result_win_lose = result_table_array(result_table_position).result_win_lose + win_lose
                                            'Print #3, result_table_array(result_table_position).Command_name & " - " & result_table_array(result_table_position).result_by_game & result_table_array(result_table_position).result_win_lose & "/" & result_table_array(result_table_position).result_score_match & vbCrLf;
                                            'запоминаем позицию второй команды
                                            Qualifying_position_2 = result_table_position
                                            'выходим из цикла
                                            result_table_position = 51
                                        Else
                                    End If
                                    result_table_position = result_table_position + 1
                                Wend
                                                              
                                If Qualifying = 1 Then
                                  'если отборы закончились то дописываем результаты в файл
                                  Open App.Path & "\result_table.txt" For Append As #3
                                  If result_table_array(Qualifying_position_1).result_win_lose > result_table_array(Qualifying_position_2).result_win_lose Then
                                    Space = Chr(32)
                                    If result_table_array(Qualifying_position_1).result_score_match >= 0 Then
                                    Space = "+"
                                    End If
                                    Print #3, result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf;
                                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf
                                    Space = Chr(32)
                                    If result_table_array(Qualifying_position_2).result_score_match >= 0 Then
                                    Space = "+"
                                    End If
                                    Print #3, result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf & vbCrLf;
                                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf & vbCrLf
                                  Else
                                    Space = Chr(32)
                                    If result_table_array(Qualifying_position_2).result_score_match >= 0 Then
                                        Space = "+"
                                    End If
                                    Print #3, result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf;
                                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_2).result_win_lose & "/" & Space & result_table_array(Qualifying_position_2).result_score_match & " " & result_table_array(Qualifying_position_2).Command_name & vbCrLf
                                    Space = Chr(32)
                                    If result_table_array(Qualifying_position_1).result_score_match >= 0 Then
                                        Space = "+"
                                    End If
                                    Print #3, result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf & vbCrLf;
                                    'Form2.Text1.Text = Form2.Text1.Text & period & vbCrLf & vbCrLf
                                    'Form2.Text1.Text = Form2.Text1.Text & result_table_array(Qualifying_position_1).result_win_lose & "/" & Space & result_table_array(Qualifying_position_1).result_score_match & " " & result_table_array(Qualifying_position_1).Command_name & vbCrLf & vbCrLf
                                  End If
                                  Close #3
                                End If
                                
                                
                               'тестируем вывод в другую форму
                               ' If Qualifying = 0 Then
                               '     Form2.Text1.Text = competition_info & vbCrLf & vbCrLf
                               '     Form2.Text1.Text = Form2.Text1.Text & "Отборочные игры" & vbCrLf & vbCrLf & "Group A" & vbCrLf & vbCrLf
                               '     q = 0
                               '     While result_table_array(q).Command_name <> ""
                                '        Form2.Text1.Text = Form2.Text1.Text & result_table_array(q).result_by_game & null_group((result_table_array(q).counter_match)) & " Итог : " & result_table_array(q).result_win_lose & "/" & Space & result_table_array(q).result_score_match & " " & result_table_array(q).Command_name & vbCrLf
                                '        If result_table_array(q).result_group <> result_table_array(q + 1).result_group And result_table_array(q + 1).result_group <> "" Then
                                '           Form2.Text1.Text = Form2.Text1.Text & vbCrLf & result_table_array(q + 1).result_group & vbCrLf & vbCrLf
                                '        End If
                                 '       If result_table_array(q + 1).result_group = "" Then
                                 '           Form2.Text1.Text = Form2.Text1.Text & vbCrLf
                                 '       End If
                                  '      q = q + 1
                                 '   Wend
                               ' End If
                                
                                
                                
                                'сортируем список по группе
                                q = 0
                                'Прокручиваем все группы
                                While q <= 7
                                    result_table_N = 0
                                    'делаем кол-во проходов по кол-ву команд
                                    While result_table_N <= team_counter
                                        'прокручиваем первую часть массива
                                        result_table_position = 0
                                        
                                        'проверить While result_table_array(result_table_position).Command_name <> "" вместо нижней строки
                                        
                                        While result_table_position <= 90
                                        'если соседние элементы из одной группы
                                        If result_table_array(result_table_position).result_group = array_group(q) And result_table_array(result_table_position + 1).result_group = array_group(q) Then
                                            'если следующий элемент больше чем предыдущий
                                            If result_table_array(result_table_position + 1).result_win_lose > result_table_array(result_table_position).result_win_lose Then
                                                result_table_array(100) = result_table_array(result_table_position + 1)
                                                result_table_array(result_table_position + 1) = result_table_array(result_table_position)
                                                result_table_array(result_table_position) = result_table_array(100)
                                            End If
                                        End If
                                        result_table_position = result_table_position + 1
                                        'считаем все команды принадлежащие данной группе
                                        'result_table_N = result_table_N + 1
                                        Wend
                                        result_table_N = result_table_N + 1
                                    Wend
                                    q = q + 1
                                Wend
                                
                                'выводим результаты в файл
                                'если признак отборов то записываем результаты в файл
                                If Qualifying = 0 Then
                                    Open App.Path & "\result_table.txt" For Output As #3
                                    Print #3, "РАСЧЕТ ИДЕТ В ТЕСТОВОМ ВАРИАНТЕ" & vbCrLf & "Сортировка только по кол-ву выигранных геймов(личные встречи не учитываются)" & vbCrLf & "Выложеные результаты подлежат проверке судьёй информатором" & vbCrLf & vbCrLf & "Название мероприятия : " & competition_info & vbCrLf & vbCrLf;
                                    'тестируем вывод в другую форму
                                    'Form2.Text1.Text = competition_info & vbCrLf & vbCrLf
                                    Print #3, "Отборочные игры" & vbCrLf & vbCrLf & "Group A" & vbCrLf & vbCrLf;
                                    'Form2.Text1.Text = Form2.Text1.Text & "Отборочные игры" & vbCrLf & vbCrLf & "Group A" & vbCrLf & vbCrLf
                                    result_table_position = 0
                                    While result_table_array(result_table_position).Command_name <> ""
                                            Space = Chr(32)
                                            If result_table_array(result_table_position).result_score_match >= 0 Then
                                                Space = "+"
                                            End If
                                        Print #3, result_table_array(result_table_position).result_by_game & " Итог : " & result_table_array(result_table_position).result_win_lose & "/" & Space & result_table_array(result_table_position).result_score_match & " " & result_table_array(result_table_position).Command_name & vbCrLf;
                                        'Form2.Text1.Text = Form2.Text1.Text & result_table_array(result_table_position).result_by_game & null_group((result_table_array(result_table_position).counter_match)) & " Итог : " & result_table_array(result_table_position).result_win_lose & "/" & Space & result_table_array(result_table_position).result_score_match & " " & result_table_array(result_table_position).Command_name & vbCrLf
                                        If result_table_array(result_table_position).result_group <> result_table_array(result_table_position + 1).result_group And result_table_array(result_table_position + 1).result_group <> "" Then
                                            Print #3, vbCrLf & result_table_array(result_table_position + 1).result_group & vbCrLf & vbCrLf;
                                            'Form2.Text1.Text = Form2.Text1.Text & vbCrLf & result_table_array(result_table_position + 1).result_group & vbCrLf & vbCrLf
                                        End If
                                        If result_table_array(result_table_position + 1).result_group = "" Then
                                            Print #3, vbCrLf;
                                            'Form2.Text1.Text = Form2.Text1.Text & vbCrLf
                                        End If
                                        result_table_position = result_table_position + 1
                                    Wend
                                    Close #3
                                End If
                                 
                                
            'записываем результаты матчей в файл
            Open App.Path & "\results.txt" For Append As #2
            Print #2, vbCrLf & Time & vbCrLf & "Результаты матча:" & vbCrLf;
            'если 2 пита то пишем только один результат
            'If Combo1.Text <> "" Then
            '    Print #2, Combo1.Text & " " & Label7.Caption & "-" & Label8.Caption & " " & Combo2.Text & vbCrLf;
            'End If

            'Print #2, Combo3.Text & " " & team_1 & " - " & team_2 & " " & Combo4.Text & vbCrLf;
            'Print #2, "Ход матча по геймам:" & vbCrLf & result_info_2 & vbCrLf & team_1 & " - " & team_2 & vbCrLf;
            'Print #2, vbCrLf;
                If result_info_counter = 0 Then
                    Print #2, result_info_2 & overtime_info & "Оставшееся время : " & Label3.Caption & vbCrLf;
                    Else
                    'если вторая пара играет
                    Print #2, result_info & overtime_info & "Оставшееся время : " & Label3.Caption & vbCrLf;
                End If
            
            Close #2
    
            'проверяем лицензию
            If license = 0 Then
                license_counter = license_counter + 1
            End If
            
            'счетчик выхода лицензии
            If license_counter = 2 Then
                PushButton = MsgBox("Демо-режим окончен, активируйте программу.", 48)
                'выходим из программы
                End
            End If
                
            If Command8.BackColor = color_active Then
                Finish_break = -120
                Finish_main_break = -120
            End If
            If Command9.BackColor = color_active Then
                Finish_break = -60
                Finish_main_break = -60
            End If
            If Command10.BackColor = color_active Then
                Finish_break = -30
                Finish_main_break = -30
            End If
            If Command38.BackColor = color_active Then
                Finish_break = finish_break_custom
                Finish_main_break = finish_break_custom
            End If
            Label4.Caption = TimeInLabel(-Finish_break)
            Form3.Label4.Caption = TimeInLabel(-Finish_break)
            Command11.Caption = "Время"
            Command41.Caption = "Overtime"
            Command41.BackColor = color_unactive
            Command41.Visible = False
            a_break = 0
            overtime = 0
            overtime1 = 0
            overtime_save = 0
            Command5.Caption = "Следующие команды"
            'скрываем кнопку продолжить гейм
            Command11.Visible = False
            'зеленим кнопки команд
            Command20.BackColor = &HFF00&
            Command21.BackColor = &HFF00&
            'делаем невидимыми кнопки стар питтайма и время
            Command12.Visible = False
            Command11.Visible = False
            'уведомляем о возможности сделать настройки
            Text8.BackColor = &HFF00&
            Text9.BackColor = &HFF00&
            Text2.BackColor = &HFF00&
            Text14.BackColor = &HFF00&
            If Command1.BackColor = color_active Then
                Finish = -360
                Finish_main_break = -360
            End If
            If Command2.BackColor = color_active Then
                Finish = -480
                Finish_main = -480
            End If
            If Command3.BackColor = color_active Then
                Finish = -600
                Finish_main = -600
            End If
            If Command31.BackColor = color_active Then
                Finish = finish_custom
                Finish_main = finish_custom
            End If
            If Command48.BackColor = color_active Then
                Finish = -600
                Finish_main = -600
            End If
        Label3.Caption = TimeInLabel(-Finish)
        Form3.Label3.Caption = TimeInLabel(-Finish)
        Label4.Caption = TimeInLabel(-Finish_break)
        Form3.Label4.Caption = TimeInLabel(-Finish_break)
        Combo1.Visible = True
        Combo2.Visible = True
        Label9.Caption = ""
        Text5.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        team_1 = 0
        team_2 = 0
        Label1.Caption = team_1
        Form3.Label1.Caption = team_1
        Label2.Caption = team_2
        Form3.Label2.Caption = team_2
        Combo1.Text = ""
        Combo2.Text = ""
        Combo3.Text = ""
        Form3.Text1.Text = Combo3.Text
        Combo4.Text = ""
        Form3.Text2.Text = Combo4.Text
        Label7.Caption = team_1
        Label8.Caption = team_2
        label1_save = team_1
        label2_save = team_2
        'скидываем цвет кнопок timeout для новых команд
        Command6.BackColor = color_unactive_timeout
        Form3.Command2.BackColor = Command6.BackColor
        Command7.BackColor = color_unactive_timeout
        Form3.Command1.BackColor = Command7.BackColor
        Command27.BackColor = color_unactive_timeout
        Command28.BackColor = color_unactive_timeout
        Command29.BackColor = color_unactive_timeout
        Command30.BackColor = color_unactive_timeout
        'сбрасываем инфу об сбигранных матчах для записи в файл
        result_info = ""
        result_info_2 = ""
        result_info_counter = 0
         
    
        'Open App.Path & "\confirm.txt" For Input As #1
        'Do Until EOF(1)
        'Line Input #1, MyText
        'Line Input #1, MyText
        'If MyText = "FINAL" Then
        '        day_end = k + 1
        '    End If
        'Loop
        'Close #1
    
    
        'выводим название команд из расписания
        k = k + 1
        
        'проверяем окончание дня
        'If day_end = k Then
        '
        '        MsgBox ("Турнир окончен. Спасибо что были с нами.")
        '        End
        'End If
        
        'проверяем отборы или нет и даем возможность заполнить названия команд
        Open App.Path & "\schedule.txt" For Input As #1
        Do Until EOF(1)
        Line Input #1, MyText
        
        If MyText = "*" Then
'        If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Or MyText = "FINAL" Then
                'period = "Четверти"
                'Line Input #1, MyText
                Line Input #1, MyText
                If MyText = k Then
                   
                   'проверяем окончание дня
                   Open App.Path & "\confirm.txt" For Input As #11
                   Do Until EOF(11)
                   Line Input #11, MyText
                   Line Input #11, MyText
                   If MyText = "FINAL" Then
                        MsgBox ("Турнир окончен. Спасибо что были с нами.")
                        Label10.Caption = Board_team(t_1main, t_1pit)
                        'Sleep (300)
                        'Label10.Caption = Board_team(t_1main, t_1pit)
                        Sleep (300)
                        Label10.Caption = Board(t_1main, t_1pit)
                   End
                   End If
                   Loop
                   Close #11
                   
                    'печатаем в файл признак того что данные в расписании не обновлены
                    Open App.Path & "\confirm.txt" For Output As #10
                    Print #10, "No";
                    Close #10
                    game = 1
                    Do Until game = 0
                    MsgBox ("Пожалуйста введите команды прошедшие дальше в файл расписания")
                    Open App.Path & "\confirm.txt" For Input As #10
                        Line Input #10, MyText
                        'если признак обновления расписания,то перестаем спрашивать и идем дальше
                        If MyText = "Yes" Then
                            game = 0
                        End If
                    Close #10
                    Loop
                    
                    Qualifying = 1
                    'Form2.Text1.FontSize = Int((Form2.Height \ 15) \ ((Form1.Text15.Text + Form1.Text16.Text * 3 + 9) * 0.75))
                    'Form2.Text1.Text = period & vbCrLf & vbCrLf
                    result_table_position = 0
                    'делаем кол-во проходов по кол-ву команд
                    While result_table_position <= team_counter
                        result_table_array(result_table_position).result_by_game = ""
                        result_table_array(result_table_position).result_win_lose = 0
                        result_table_array(result_table_position).result_score_match = 0
                        result_table_position = result_table_position + 1
                    Wend
                End If
        End If
        Loop
        Close #1
           
        'проверяем на окончание дня
        'Open App.Path & "\schedule.txt" For Input As #1
        'Do Until EOF(1)
        'Line Input #1, MyText
        'If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Or MyText = "FINAL" Then
        '    If MyText = "FINAL" Then
        '        day_end = k + 1
        '    End If
            
            'period = MyText
            'Open App.Path & "\result_table.txt" For Append As #3
            'Print #3, period & vbCrLf & vbCrLf;
            'Close #3
        'End If
        'Loop
        'Close #1
        
        'проверяем финал
        'Open App.Path & "\schedule.txt" For Input As #1
        'Do Until EOF(1)
        'Line Input #1, MyText
        'If MyText = "FINAL" Then
        '        period = MyText
        '        Line Input #1, MyText
        '        Line Input #1, MyText
        '        If MyText = k Then
        '            MsgBox ("Пожалуйста введите команды прошедшие дальше в файл расписания")
        '            Open App.Path & "\result_table.txt" For Append As #3
        '            Print #3, period & vbCrLf & vbCrLf;
        '            Close #3
        '            day_end = k + 1
        '            Qualifying = 1
        '            Form2.Text1.FontSize = Int((Form2.Height \ 15) \ ((Form1.Text15.Text + Form1.Text16.Text * 3 + 9) * 0.75))
        '            Form2.Text1.Text = period & vbCrLf & vbCrLf
        '            result_table_position = 0
                    'делаем кол-во проходов по кол-ву команд
        '            While result_table_position <= team_counter
        '                result_table_array(result_table_position).result_by_game = ""
        '                result_table_array(result_table_position).result_win_lose = 0
        '                result_table_array(result_table_position).result_score_match = 0
        '                result_table_position = result_table_position + 1
        '            Wend
        '        End If
        'End If
        'Loop
        'Close #1
        
        Open App.Path & "\schedule.txt" For Input As #1
        Do Until EOF(1)
        Line Input #1, MyText
        
        'пишем какая четверть в файл результата
        If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Or MyText = "FINAL" Then
            period = MyText
        End If
        
            If MyText = k Then
                Open App.Path & "\result_table.txt" For Append As #3
                Print #3, period & vbCrLf & vbCrLf;
                Close #3
                Line Input #1, MyText
                Combo3.Text = MyText
                Form3.Text1.Text = Combo3.Text
                Command16.Caption = "Гейм команды" & vbCrLf & Combo3.Text
                Command7.Caption = "ТimeOut" & vbCrLf & Combo3.Text
                Line Input #1, MyText
                Combo4.Text = MyText
                Form3.Text2.Text = Combo4.Text
                Command14.Caption = "Гейм команды" & vbCrLf & Combo4.Text
                Command6.Caption = "ТimeOut" & vbCrLf & Combo4.Text
                Line Input #1, MyText
                If MyText <> "" Then
                    Combo1.Text = MyText
                '    Else
                '    i1 = 1
                End If
                Line Input #1, MyText
                If MyText <> "" Then
                    Combo2.Text = MyText
                '    Else
                '    i1 = 1
                End If
            End If
        Loop
        Close #1
    
        '    If Combo1.Text = "" And Combo2.Text = "" Then
        '        i1 = i1 - 1
        '    End If

    End If

    'End If
    
    
    If i1 = 4 Or i1 = 3 Then
        Command12.Visible = True
        'Command11.Visible = True
        Text8.BackColor = &HFFFFFF
        Text9.BackColor = &HFFFFFF
        Text2.BackColor = &HFFFFFF
        Text14.BackColor = &HFFFFFF
        If Command8.BackColor = color_active Then
            Finish_break = -120
            Finish_main_break = -120
        End If
        If Command9.BackColor = color_active Then
            Finish_break = -60
            Finish_main_break = -60
        End If
        If Command10.BackColor = color_active Then
            Finish_break = -30
            Finish_main_break = -30
        End If
        If Command38.BackColor = color_active Then
            Finish_break = finish_break_custom
            Finish_main_break = finish_break_custom
        End If
        
        'Label4.Caption = TimeInLabel(-Finish_break)
    
        If Command1.BackColor = color_active Then
            Finish = -360
            Finish_main_break = -360
        End If
        If Command29.BackColor = color_active Then
            Finish = -480
            Finish_main = -480
        End If
        If Command3.BackColor = color_active Then
            Finish = -600
            Finish_main = -600
        End If
        
        If Command31.BackColor = color_active Then
            Finish = finish_custom
            Finish_main = finish_custom
        End If
        
            If Command48.BackColor = color_active Then
                Finish = -600
                Finish_main = -600
            End If
        
        Label3.Caption = TimeInLabel(-Finish)
        Label6.Caption = Label3.Caption
        Form3.Label3.Caption = TimeInLabel(-Finish)
        Label4.Caption = TimeInLabel(-Finish_break)
        Form3.Label4.Caption = TimeInLabel(-Finish_break)
        Command11.Caption = "Время"
        Command41.Caption = "Overtime"
        Command41.BackColor = color_unactive
        Command41.Visible = False
        i1 = 0
        If Combo1.Text = "" Or Combo2.Text = "" Then
            i1 = 1
        End If
        Command12.Visible = True
        
        Label10.Caption = Board_team(t_1main, t_1pit)
        'Sleep (300)
        'Label10.Caption = Board_team(t_1main, t_1pit)

    End If
End Sub

Private Sub Command51_Click()
'делаем видимой вторую форму
'    Form2.Show
End Sub

Private Sub Command52_Click()
'делаем видимой третью форму
Form3.Label1.Caption = Label1.Caption
Form3.Label2.Caption = Label2.Caption
Form3.Label3.Caption = Label3.Caption
Form3.Label4.Caption = Label4.Caption
Form3.Text1.Text = Combo3.Text
Form3.Text2.Text = Combo4.Text
Form3.Show
End Sub


Private Sub Command53_Click()
If Command53.BackColor = color_unactive Then
        'блокируем изменение порта
        Combo7.Locked = True
        Combo9.Locked = True
    Command53.BackColor = &HFF&
    Command53.Caption = "Разблокируем"
Else
Command53.BackColor = color_unactive
        'разблокируем изменение порта
        Combo7.Locked = False
        Combo9.Locked = False
        Command53.Caption = "Блокируем порт"
        'переопрашиваем порты
        If SComm1.PortOpen = True Then
            SComm1.PortOpen = False
        End If
        If SComm2.PortOpen = True Then
            SComm2.PortOpen = False
        End If
        Combo7.Clear
        Combo7.Text = "Выберите порт подключения приёмника"
        Combo9.Clear
        Combo9.Text = "Табло в поле"
        Call PopulateList
End If
End Sub


Private Sub Command54_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label4.Caption >= "0.00,00" Then
    Finish_break = Finish_break - 60
    If Finish_break > 0 Then
    Finish_break = 0
    End If
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
    End If
End If
End If
End Sub

Private Sub Command55_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Label4.Caption >= "0.00,00" Then
    Finish_break = Finish_break + 60
    If Finish_break > 0 Then
    Finish_break = 0
    End If
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
    End If
End If
End If
End Sub

'при загрузке формы
Sub Form_Initialize()

' проверяем на существование наши файлы
If Dir$(App.Path & "\schedule.txt") = "" Then
MsgBox ("Файл schedule.txt отсутствует." & vbCrLf & "Пожалуйста запустите schedule.exe для создания файла расписания.")
End
End If
If Dir$(App.Path & "\sounds\10_man.wav") = "" Then
MsgBox ("Файл sounds\10_man.wav отсутствует." & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\10_woman.wav") = "" Then
MsgBox ("Файл sounds\10_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\30_man.wav") = "" Then
MsgBox ("Файл sounds\30_man.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\30_woman.wav") = "" Then
MsgBox ("Файл sounds\30_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\60_man.wav") = "" Then
MsgBox ("Файл sounds\60_man.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\60_woman.wav") = "" Then
MsgBox ("Файл sounds\60_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\beep_seconds.wav") = "" Then
MsgBox ("Файл sounds\beep_seconds.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\buzzer.wav") = "" Then
MsgBox ("Файл sounds\buzzer.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\dtmw.wav") = "" Then
MsgBox ("Файл sounds\dtmw.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\dtmw_stop.wav") = "" Then
MsgBox ("Файл sounds\dtmw_stop.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\field_in_game_man.wav") = "" Then
MsgBox ("Файл sounds\field_in_game_man.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\field_in_game_woman.wav") = "" Then
MsgBox ("Файл sounds\field_in_game_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\minute_man.wav") = "" Then
MsgBox ("Файл sounds\minute_man.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\minute_woman.wav") = "" Then
MsgBox ("Файл sounds\minute_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\time_man.wav") = "" Then
MsgBox ("Файл sounds\time_man.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If
If Dir$(App.Path & "\sounds\time_woman.wav") = "" Then
MsgBox ("Файл sounds\time_woman.wav отсутствует. " & vbCrLf & "Пожалуйста установите программу корректно.")
End
End If

Qualifying = 0
correct_shedule = 0
Open App.Path & "\schedule.txt" For Input As #1
Do Until EOF(1)
Line Input #1, MyText
If MyText = "Teams playing schedule" Then
    correct_shedule = 1
    Line Input #1, MyText
    Line Input #1, MyText
    If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Or MyText = "FINAL" Then
        Line Input #1, MyText
        Qualifying = 1
    End If
    k = MyText
End If
Loop
Close #1
If correct_shedule = 0 Then
    MsgBox ("Файл расписания неккоректно заполнен")
    End
End If

'добавляем пустую строку в файл расписания для защиты от кривых рук
Open App.Path & "\schedule.txt" For Append As #1
Print #1, vbCrLf & vbCrLf;
Close #1


Space = Chr(32)

'счетчик кол-ва групп
group_counter = 0

'добавляем пустой идентификатор в кнопочный список
'Combo8.AddItem ""
'Combo10.AddItem ""
'Combo11.AddItem ""
'Combo12.AddItem ""


'спрашиваем название мероприятия
competition_info = InputBox("Введите название проводимого мероприятия : ")

day_end = 0
team_counter = 0
array_group(0) = "Group A"
array_group(1) = "Group B"
array_group(2) = "Group C"
array_group(3) = "Group D"
array_group(4) = "Group E"
array_group(5) = "Group F"
array_group(6) = "Group G"
array_group(7) = "Group H"
result_table_N = 0
result_table_index = 0
result_table_position = 0
i1 = 0
'признак для записи результатов геймов в файл
result_info_counter = 0
result_info = ""
result_info_2 = ""
color_unactive = &HE0E0E0
color_active = &HFFC0C0    '&HC0C0C0
color_unactive_timeout = &HFF00&
color_active_timeout = &HFF&
        Command27.BackColor = color_unactive_timeout
        Command28.BackColor = color_unactive_timeout
        Command29.BackColor = color_unactive_timeout
        Command30.BackColor = color_unactive_timeout
'кнопка теста не активна
Command44.BackColor = color_unactive
'лицензии нет
license = 0

'вносим команды в список
'вносим все команды в ниспадающий список
'k = 1
Open App.Path & "\schedule.txt" For Input As #1
MyText = 0
Do Until MyText = "Teams playing schedule"
Line Input #1, MyText
If MyText <> "" Then
    If MyText <> "Teams playing schedule" Then
        If MyText <> "A list of teams playing" Then
          If MyText = "Group A" Or MyText = "Group B" Or MyText = "Group C" Or MyText = "Group D" Or MyText = "Group E" Or MyText = "Group F" Or MyText = "Group G" Or MyText = "Group H" Then
             '100 запись хранит признак к какой группе принадлежит команда
                result_table_array(99).result_group = MyText
                group_counter = group_counter + 1
                Else
                Combo1.AddItem MyText
                Combo2.AddItem MyText
                Combo3.AddItem MyText
                Combo4.AddItem MyText
                result_table_array(result_table_position).Command_name = MyText
                result_table_array(result_table_position).result_by_game = Space & Space & Space & Space
                result_table_array(result_table_position).result_group = result_table_array(99).result_group
                result_table_position = result_table_position + 1
                'выясняем количество играющих команд
                team_counter = team_counter + 1
           End If
         End If
    End If
End If
Loop
Close #1
    
'выясняем кол-во команд и групп для второй формы
Text15.Text = team_counter
Text16.Text = group_counter

'считаем кол-во матчей у каждой команды
q = 0
null_group(0) = ""
While q <= team_counter
    null_group(q + 1) = null_group(q) + " 0/0 |"
    If Dir$(App.Path & "\schedule.txt") <> "" Then
        Open App.Path & "\schedule.txt" For Input As #1
        result_table_array(q).counter_match = 0
        MyText = 0
        Do Until EOF(1)
        Line Input #1, MyText
        If result_table_array(q).Command_name = MyText Then
           result_table_array(q).counter_match = result_table_array(q).counter_match + 1
        End If
        Loop
        Close #1
        result_table_array(q).counter_match = result_table_array(q).counter_match - 1
        'если команда только в списке расписания то делаем защиту от выхода из массива
        If result_table_array(q).counter_match = 0 Then
            result_table_array(q).counter_match = 1
        End If
    End If
    q = q + 1
Wend

'тестируем вывод в другую форму
'Form2.Text1.Text = competition_info & vbCrLf & vbCrLf
'Form2.Text1.Text = Form2.Text1.Text & "Отборочные игры" & vbCrLf & vbCrLf & "Group A" & vbCrLf & vbCrLf
'q = 0
'While result_table_array(q).Command_name <> ""
'    Form2.Text1.Text = Form2.Text1.Text & result_table_array(q).result_by_game & null_group(result_table_array(q).counter_match) & " Итог : " & result_table_array(q).result_win_lose & "/" & Space & result_table_array(q).result_score_match & " " & result_table_array(q).Command_name & vbCrLf
'    If result_table_array(q).result_group <> result_table_array(q + 1).result_group And result_table_array(q + 1).result_group <> "" Then
'       Form2.Text1.Text = Form2.Text1.Text & vbCrLf & result_table_array(q + 1).result_group & vbCrLf & vbCrLf
'    End If
'    If result_table_array(q + 1).result_group = "" Then
'        Form2.Text1.Text = Form2.Text1.Text & vbCrLf
'    End If
'    q = q + 1
'Wend

'проверяем отборы или нет и даем возможность заполнить названия команд
'        Open App.Path & "\schedule.txt" For Input As #1
'        Do Until EOF(1)
'        Line Input #1, MyText
'        If MyText = "*" Then
'        If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Then
                'Line Input #1, MyText
'                Line Input #1, MyText
'                If MyText = k Then
                    'Open App.Path & "\result_table.txt" For Append As #3
                    'Print #3, period & vbCrLf & vbCrLf;
                    'Close #3
'                    Qualifying = 1
'                    result_table_position = 0
                    'делаем кол-во проходов по кол-ву команд
'                   While result_table_position <= team_counter
'                        result_table_array(result_table_position).result_by_game = ""
'                        result_table_array(result_table_position).result_win_lose = 0
'                        result_table_array(result_table_position).result_score_match = 0
'                        result_table_position = result_table_position + 1
'                    Wend
'                End If
'        End If
'        Loop
'        Close #1
           
        
        'проверяем финал
 '       Open App.Path & "\schedule.txt" For Input As #1
 '       Do Until EOF(1)
 '       Line Input #1, MyText
 '       If MyText = "FINAL" Then
 '               period = MyText
 '               Line Input #1, MyText
 '               Line Input #1, MyText
 '               If MyText = k Then
 '                   Open App.Path & "\result_table.txt" For Append As #3
 '                   Print #3, period & vbCrLf & vbCrLf;
 '                   Close #3
 '                   day_end = k + 1
 '                   Qualifying = 1
 '                   result_table_position = 0
                    'делаем кол-во проходов по кол-ву команд
 '                   While result_table_position <= team_counter
 '                       result_table_array(result_table_position).result_by_game = ""
 '                       result_table_array(result_table_position).result_win_lose = 0
 '                       result_table_array(result_table_position).result_score_match = 0
 '                       result_table_position = result_table_position + 1
 '                   Wend
 '               End If
 '       End If
 '       Loop
 '       Close #1

'добавляем пустую команду
Combo1.AddItem ""
Combo2.AddItem ""
Combo3.AddItem ""
Combo4.AddItem ""

'вносим косанды на 1/2 матч
Open App.Path & "\schedule.txt" For Input As #1
Do Until EOF(1)
Line Input #1, MyText
    
        'пишем какая четверть в файл результата
        If MyText = "1\16" Or MyText = "1\8" Or MyText = "1\4" Or MyText = "1\2" Or MyText = "FINAL" Then
            period = MyText
        End If
        
   
    If MyText = k Then
                Open App.Path & "\result_table.txt" For Append As #3
                Print #3, period & vbCrLf & vbCrLf;
                Close #3
    Line Input #1, MyText
    Combo3.Text = MyText
    Form3.Text1.Text = Combo3.Text
    Command16.Caption = "Гейм команды" & vbCrLf & Combo3.Text
    Command7.Caption = "ТimeOut" & vbCrLf & Combo3.Text
    Line Input #1, MyText
    Combo4.Text = MyText
    Form3.Text2.Text = Combo4.Text
    Command14.Caption = "Гейм команды" & vbCrLf & Combo4.Text
    Command6.Caption = "ТimeOut" & vbCrLf & Combo4.Text
    Line Input #1, MyText
                If MyText <> "" Then
                    Combo1.Text = MyText
                    Else
                    Combo1.Text = MyText
                    i1 = 1
                End If
                Line Input #1, MyText
                If MyText <> "" Then
                    Combo2.Text = MyText
                    Else
                    Combo2.Text = MyText
                    i1 = 1
                End If

'Line Input #1, MyText
'    Combo1.Text = MyText
'    Line Input #1, MyText
'    Combo2.Text = MyText
    End If
Loop
Close #1

'счетчик пробных запусков без лицензии
license_counter = 0

'играем до 3х очков
score_end = 3

'обнуляем входной буфер от com порта
Text7.Text = ""

'задаем звучать даже в свернутом состоянии
Bufferdesc.lFlags = DSBCAPS_GLOBALFOCUS Or DSBCAPS_STATIC
second_counter = 0
    
    ' настройки для работы с портом
    ' Fire Rx Event Every 3 Bytes(3)
    'SComm1.RThreshold = 3
    SComm1.RThreshold = 2
 
    ' When Inputting Data, Input 11 Bytes at a time(2)
    'посылка 3 байта
    'SComm1.InputLen = 2
    SComm1.InputLen = 3
 
    ' 115200 Baud, No Parity, 8 Data Bits, 1 Stop Bit
    SComm1.Settings = "115200,N,8,1"
    SComm2.Settings = "9600,N,8,1"
    
    'скрываем кнопку продолжить гейм
    Command11.Visible = False
    'красим время овертайма
    Command45.BackColor = color_active
    
    'по умолчанию озвучиваем секунды
    Command34.BackColor = color_active
    'по умолчанию озвучка женская
    Command36.BackColor = color_active
    'время overtime 5 мин
    overtime_time = -300
    'присваиваем переменной а значение "0"
    a = 0
    a_break = 0
    a_2 = 0
    a_break_2 = 0
    'задаем перерыв между матчами 2 минуты
    finish_break_custom = -120
    'задаем время матча 8 мин и break 2 мин
    Finish = -480
    Finish_main = -480
    Finish_main_save = -480
    Finish_save = Finish
    finish_custom = -480
    Label3.Caption = TimeInLabel(-Finish)
    Form3.Label3.Caption = TimeInLabel(-Finish)
    Label6.Caption = TimeInLabel(-Finish)
    Command38.BackColor = color_active
    Command2.BackColor = color_active
    Finish_break = -120
    Finish_main_break = -120
    Label4.Caption = TimeInLabel(-Finish_break)
    Form3.Label4.Caption = TimeInLabel(-Finish_break)
    Command8.BackColor = color_active
    Finish_2 = -480
    Finish_main_2 = -480
    Finish_break_2 = -120
    Finish_main_break_2 = -120
    'задаем переменные для смены команд в 4 пита
    label1_save = 0
    label1_save = 0
    label3_save = Finish_main
    'задаём счет команд 0-0 и размер шрифта
    team_1 = 0
    Label1.Caption = team_1
    Form3.Label1.Caption = team_1
    team_2 = 0
    Label2.Caption = team_2
    Form3.Label2.Caption = team_2
    Label7.Caption = 0
    Label8.Caption = 0
    'делаем список com портов
    Call PopulateList
    i2 = 0
    match_point = 0
    Command11.Caption = "Время"
    overtime = 0
    overtime1 = 0
    overtime_save = 0
    Command41.Caption = "Overtime"
    Command41.BackColor = color_unactive
    'уведомляем о возможности сделать настройки
    Text8.BackColor = &HFF00&
    Text9.BackColor = &HFF00&
    Text2.BackColor = &HFF00&
    Text14.BackColor = &HFF00&
    'скрываем кнопку гейм без очка
    Command42.Visible = False
    t1 = 0
    Label10.Caption = Board_team(t_1main, t_1pit)
End Sub

'Увеличиваем счет команды 2 на 1
Private Sub Command14_Click()

Label10.Caption = Board_team(t_1main, t_1pit)

'красим кнопку в серый цвет
Command14.BackColor = &HE0E0E0
If a_break Mod 2 = 0 Then
    If a Mod 2 = 0 Then
        'If overtime = 0 Then
            If Command4.Caption = "Поле в игре" Then
                If Finish1_save <> Finish_main1_save Or Finish <> Finish_main Then
                    team_2 = team_2 + 1
                    Label2.Caption = team_2
                    Label10.Caption = Board(t_1main, t_1pit)
                    Form3.Label2.Caption = team_2
                    If result_info_counter = 0 Then
                        'если первая пара играет
                        result_info_2 = result_info_2 & Combo3.Text & " | " & Label1.Caption & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                        Else
                        'если вторая пара играет
                        result_info = result_info & Combo3.Text & " | " & Label1.Caption & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                    End If
                    
                    'меняем признак играющих команд
                    'если 2 пита, то не меням, если 4 то меняем
                    'If Combo1.Text = "" And Combo2.Text = "" Then
                    '    Else
                    'End If
                    
                    'скрываем кнопку продолжить гейм
                    Command11.Visible = False
                    'Call Sound_dtmw_stop

                    Command4.Caption = "Старт гейма"
                    Command11.Caption = "Старт гейма"
                        If Command10.BackColor = color_active Then
                            Finish_break = -30
                            Finish_main_break = -30
                        End If
                        If Command9.BackColor = color_active Then
                            Finish_break = -60
                            Finish_main_break = -60
                        End If
                        If Command8.BackColor = color_active Then
                            Finish_break = -120
                            Finish_main_break = -120
                        End If
                        'скрываем кнопку гейм без очка
                        Command42.Visible = False
                        'скрываем кнопку овертайм
                            If overtime = 0 Then
                            Command41.Visible = False
                            End If
                        If Label3.Caption = "0.00,00" And team_1 = team_2 Then
                               Command41.Visible = True
                        End If
                        'скрываем кнопку время
                        Command11.Visible = False
                            
                            If i1 = 0 Then
                                    If overtime <> 0 Then
                                        'если идет овертайм
                                        Call Command50_Click 'новый конец матча
                                        'запускаем пит
                                        Call Command12_Click
                                    'если не овертайм
                                    Else
                                        'если время вышло
                                        If Label3.Caption = "0.00,00" Then
                                            'если счет равный
                                            If team_1 = team_2 Then
                                                'переходим в овертайм
                                                Call Command49_Click
                                                'запускаем пит
                                                Call Command12_Click
                                            'если не равный
                                            Else
                                                Call Command50_Click 'новый конец матча
                                                'запускаем пит
                                                Call Command12_Click
                                            End If
                                        'если время не вышло
                                        Else
                                            'счет равный
                                            If team_2 = score_end Then
                                                Call Command50_Click 'новый конец матча
                                                'запускаем пит
                                                Call Command12_Click
                                            'счет не равный
                                            Else
                                                'меняем пару
                                                Call Command26_Click
                                                'запускаем пит
                                                Call Command12_Click
                                            End If
                                        End If
                                    End If
                            Else  'если 2 пита
                                If Label3.Caption <> "0.00,00" Then
                                    If team_2 = score_end Or overtime <> 0 Then
                                        Call Command50_Click 'новый конец матча
                                    Else
                                        Finish_break = -120
                                        Finish_main_break = -120
                                        Call Command12_Click 'запускаем пит
                                    End If
                                Else
                                    If team_1 = team_2 Then
                                        Call Command49_Click 'новый овертайм
                                        Call Command12_Click 'запускаем пит
                                    Else
                                        Call Command50_Click 'новый конец матча
                                    End If
                                End If
                            End If
                 End If
            End If
    End If
End If
End Sub

'Уменьшаем счет команды 2 на 1
Private Sub Command15_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
team_2 = team_2 - 1
If team_2 < 0 Then team_2 = 0
Label2.Caption = team_2
Form3.Label2.Caption = team_2
End If
End If
End Sub

'Увеличиваем счет команды 1 на 1
Private Sub Command16_Click()

Label10.Caption = Board_team(t_1main, t_1pit)

'красим кнопку в серый цвет
Command16.BackColor = &HE0E0E0
If a_break Mod 2 = 0 Then
    If a Mod 2 = 0 Then
       ' If overtime = 0 Then
            If Command4.Caption = "Поле в игре" Then
                If Finish1_save <> Finish_main1_save Or Finish <> Finish_main Then
                    team_1 = team_1 + 1
                    Label1.Caption = team_1
                    Label10.Caption = Board(t_1main, t_1pit)
                    Form3.Label1.Caption = team_1
                    If result_info_counter = 0 Then
                        'если первая пара играет
                        result_info_2 = result_info_2 & Combo3.Text & " | " & team_1 & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                        Else
                        'если вторая пара играет
                        result_info = result_info & Combo3.Text & " | " & team_1 & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                    End If
                    
                    'меняем признак играющих команд
                    'если 2 пита, то не меням, если 4 то меняем
                    'If Combo1.Text = "" And Combo2.Text = "" Then
                    '    Else
                    'End If
                    
                    'скрываем кнопку продолжить гейм
                    Command11.Visible = False
                    'Call Sound_dtmw_stop

                    Command4.Caption = "Старт гейма"
                    Command11.Caption = "Старт гейма"
                        If Command10.BackColor = color_active Then
                            Finish_break = -30
                            Finish_main_break = -30
                        End If
                        If Command9.BackColor = color_active Then
                            Finish_break = -60
                            Finish_main_break = -60
                        End If
                        If Command8.BackColor = color_active Then
                            Finish_break = -120
                            Finish_main_break = -120
                        End If
                        'скрываем кнопку гейм без очка
                        Command42.Visible = False
                        'скрываем кнопку овертайм
                            If overtime = 0 Then
                            Command41.Visible = False
                            End If
                            If Label3.Caption = "0.00,00" And team_1 = team_2 Then
                                Command41.Visible = True
                            End If
                        'скрываем кнопку время
                        Command11.Visible = False
                            
                            If i1 = 0 Then
                                    If overtime <> 0 Then
                                        'если идет овертайм
                                        Call Command50_Click 'новый конец матча
                                        'запускаем пит
                                        Call Command12_Click
                                    'если не овертайм
                                    Else
                                        'если время вышло
                                        If Label3.Caption = "0.00,00" Then
                                            'если счет равный
                                            If team_1 = team_2 Then
                                                'переходим в овертайм
                                                Call Command49_Click
                                                'запускаем пит
                                                Call Command12_Click
                                            'если не равный
                                            Else
                                                Call Command50_Click 'новый конец матча
                                                'запускаем пит
                                                Call Command12_Click
                                            End If
                                        'если время не вышло
                                        Else
                                            'счет равный
                                            If team_1 = score_end Then
                                                Call Command50_Click 'новый конец матча
                                                'запускаем пит
                                                Call Command12_Click
                                            'счет не равный
                                            Else
                                                'меняем пару
                                                Call Command26_Click
                                                'запускаем пит
                                                Call Command12_Click
                                            End If
                                        End If
                                    End If
                            Else  'если 2 пита
                                If Label3.Caption <> "0.00,00" Then
                                    If team_1 = score_end Or overtime <> 0 Then
                                        Call Command50_Click 'новый конец матча
                                    Else
                                        Finish_break = -120
                                        Finish_main_break = -120
                                        Call Command12_Click 'запускаем пит
                                    End If
                                Else
                                    If team_1 = team_2 Then
                                        Call Command49_Click 'новый овертайм
                                        Call Command12_Click 'запускаем пит
                                    Else
                                        Call Command50_Click 'новый конец матча
                                    End If
                                End If
                            End If
                     End If
            End If
    End If
End If
End Sub

'Уменьшаем счет команды 1 на 1
Private Sub Command17_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
team_1 = team_1 - 1
If team_1 < 0 Then team_1 = 0
Label1.Caption = team_1
Form3.Label1.Caption = team_1
End If
End If
End Sub

'кнопка сирены команды 1
Private Sub Command20_Click()
'меняем цвет на красный
If a_break Mod 2 = 0 Then
    If a Mod 2 > 0 Then
        Command20.BackColor = &HFF&
        'гудим
        Call Sound_dtmw_stop
        Sleep (100)
        Call Sound_time
        'показываем кнопку окончание матча
        Command5.Visible = True
        'показываем кнопку продолжить гейм
        Command11.Visible = True
        'показываем кнопку гейм без очка
        Command42.Visible = True
        'если меньше 60 секунд осталось делаем видимой кнопку overtime
        If (-Start + Finish + Timer) > -60 Then
            Command41.Visible = True
            'если не овертайм то делаем кнопку серой
            'If overtime = 0 Then
            'делаем активной и сбрасываем признак овертайма чтобы можно было заново его запустить
            'overtime = 0
            Command41.BackColor = color_unactive
            'End If
        End If
        'останавливаем таймер
        If a Mod 2 <> 0 Then
             Command4.Caption = "Поле в игре"
             Call Command4_Click
        End If
    End If
End If
End Sub

'кнопка сирены команды 2
Private Sub Command21_Click()
'меняем цвет на красный
If a_break Mod 2 = 0 Then
    If a Mod 2 > 0 Then
        Command21.BackColor = &HFF&
        'гудим
        Call Sound_dtmw_stop
        Sleep (100)
        Call Sound_time
        'показываем кнопку окончание матча
        Command5.Visible = True
        'показываем кнопку продолжить гейм
        Command11.Visible = True
        'показываем кнопку гейм без очка
        Command42.Visible = True
        'если меньше 60 секунд осталось делаем видимой кнопку overtime
        If (-Start + Finish + Timer) > -60 Then
            Command41.Visible = True
            'делаем активной и сбрасываем признак овертайма чтобы можно было заново его запустить
            'overtime = 0
            'If overtime = 0 Then
            Command41.BackColor = color_unactive
            'End If
        End If
        'останавливаем таймер
        If a Mod 2 <> 0 Then
            Call Command4_Click
        End If
        Command4.Caption = "Поле в игре"
    End If
End If
End Sub

'кнопка сирены
Private Sub Command19_Click()
Call Sound_buzzer
End Sub

'кнопка выбора времени матча 6 минут
Private Sub Command1_Click()

If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish = -360
Finish_main = -360
Finish_save = -360
Finish_main_save = -360
Command1.BackColor = color_active
Command2.BackColor = color_unactive
Command3.BackColor = color_unactive
Command31.BackColor = color_unactive
Command48.BackColor = color_unactive
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-Finish)
'счет окончания матча
score_end = 2
End If
End If
End If
End If
End Sub


'кнопка выбора времени матча 8 мин
Private Sub Command2_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish = -480
Finish_main = -480
Finish_save = -480
Finish_main_save = -480
Command2.BackColor = color_active
Command1.BackColor = color_unactive
Command3.BackColor = color_unactive
Command31.BackColor = color_unactive
Command48.BackColor = color_unactive
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-Finish)
'счет окончания матча
score_end = 3
End If
End If
End If
End If
End Sub


'кнопка выбора времени матча 10 мин
Private Sub Command3_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish = -600
Finish_main = -600
Finish_save = -600
Finish_main_save = -600
Command3.BackColor = color_active
Command1.BackColor = color_unactive
Command2.BackColor = color_unactive
Command31.BackColor = color_unactive
Command48.BackColor = color_unactive
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-Finish)
'счет окончания матча
score_end = 4
End If
End If
End If
End If
End Sub

'кнопка выбора времени матча Custom
Private Sub Command31_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Command31.BackColor = color_active
Command1.BackColor = color_unactive
Command2.BackColor = color_unactive
Command3.BackColor = color_unactive
Command48.BackColor = color_unactive
Command31.Caption = -finish_custom \ 60 & " мин" & "\unlim game"
Finish = finish_custom
Finish_main = Finish
Finish_save = Finish
Finish_main_save = Finish
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-Finish)
'счет окончания матча
score_end = 100
End If
End If
End If
End If
End Sub

'кнопка изменение custom +1 мин
Private Sub Command32_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
finish_custom = finish_custom - 60
Command31.Caption = -finish_custom \ 60 & " мин" & "\unlim game"
Finish = finish_custom
Finish_main = finish_custom
Finish_save = finish_custom
Finish_main_save = finish_custom
Label3.Caption = TimeInLabel(-finish_custom)
Form3.Label3.Caption = TimeInLabel(-finish_custom)
Label6.Caption = TimeInLabel(-finish_custom)
Command31.BackColor = color_active
Command3.BackColor = color_unactive
Command1.BackColor = color_unactive
Command2.BackColor = color_unactive
Command48.BackColor = color_unactive
'счет окончания матча
score_end = 5
End If
End If
End If
End If
End Sub

'кнопка изменение custom -1 мин
Private Sub Command33_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
If finish_custom < -60 Then
finish_custom = finish_custom + 60
Command31.Caption = -finish_custom \ 60 & " мин" & "\unlim game"
Finish = finish_custom
Finish_main = Finish
Finish_save = Finish
Finish_main_save = Finish
Label3.Caption = TimeInLabel(-Finish)
Form3.Label3.Caption = TimeInLabel(-Finish)
Label6.Caption = TimeInLabel(-finish_custom)
Command31.BackColor = color_active
Command3.BackColor = color_unactive
Command1.BackColor = color_unactive
Command2.BackColor = color_unactive
Command48.BackColor = color_unactive
'счет окончания матча
score_end = 5
End If
End If
End If
End If
End If
End Sub

'кнопка TimeOut 2-й команды
Private Sub Command6_Click()
If Int((Start - Timer - Finish_break) * 100) > 1000 Then
If a_break Mod 2 <> 0 Then
If Command6.BackColor = color_unactive_timeout Then
Command6.BackColor = color_active_timeout
Form3.Command2.BackColor = Command6.BackColor
Finish_break = Finish_break - 60
Call Sound_dtmw
Label4.Caption = TimeInLabel(-Finish_break)
Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If
End If
End If
End Sub

'кнопка TimeOut 1-й команды
Private Sub Command7_Click()
If Int((Start - Timer - Finish_break) * 100) > 1000 Then
If a_break Mod 2 <> 0 Then
If Command7.BackColor = color_unactive_timeout Then
Command7.BackColor = color_active_timeout
Form3.Command1.BackColor = Command7.BackColor
Finish_break = Finish_break - 60
Call Sound_dtmw
Label4.Caption = TimeInLabel(-Finish_break)
Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If
End If
End If
End Sub

'кнопка выбора времени пита 2 мин
Private Sub Command8_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish_break = -120
Finish_main_break = -120
Command8.BackColor = color_active
Command9.BackColor = color_unactive
Command10.BackColor = color_unactive
Command38.BackColor = color_unactive
If Command38.BackColor = color_unactive Then
Label4.Caption = TimeInLabel(-Finish_break)
Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If
End If
End If
End If
End If
End Sub

'кнопка выбора времени пита 1 мин
Private Sub Command9_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish_break = -60
Finish_main_break = -60
Command9.BackColor = color_active
Command8.BackColor = color_unactive
Command10.BackColor = color_unactive
Command38.BackColor = color_unactive
If Command38.BackColor = color_unactive Then
Label4.Caption = TimeInLabel(-Finish_break)
Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If
End If
End If
End If
End If
End Sub

'кнопка выбора времени пита 30 сек
Private Sub Command10_Click()
If a_break Mod 2 = 0 Then
If a Mod 2 = 0 Then
If Finish1_save = Finish_main1_save And Finish = Finish_main Then
If a_break = 0 Then
Finish_break = -30
Finish_main_break = -30
Command10.BackColor = color_active
Command8.BackColor = color_unactive
Command9.BackColor = color_unactive
Command38.BackColor = color_unactive
If Command38.BackColor = color_unactive Then
Label4.Caption = TimeInLabel(-Finish_break)
Form3.Label4.Caption = TimeInLabel(-Finish_break)
End If
End If
End If
End If
End If
End Sub

'кнопка "Старт - стоп"
Sub Command4_Click()
Label10.Caption = Board_team(t_1main, t_1pit)
If a_break Mod 2 = 0 Then
    'увеличиваем ко-во нажатий на кнопку на 1
    a = a + 1
    'если значение нечетное, включаем секундомер
    If a Mod 2 <> 0 Then
        'скрываем кнопку окончание матча
        Command5.Visible = False
        'скрываем кнопку гейм без очка
        Command42.Visible = False
        Switch = True
        Start = Timer
        'Меняем надпись на кнопке на "Время"
        Command11.Caption = "Время"
        Command11.BackColor = color_active
        Command4.Caption = "Время"
        'зеленим кнопки команд
        Command20.BackColor = &HFF00&
        Command21.BackColor = &HFF00&
        'гудим
        'Call Sound_buzzer
        'Finish = Finish - 0.01
        While Switch
            'выводим в метку текущее показание секундомера
            t1 = Start - Timer - Finish
                
            t_1main = t1
            t_1pit = -Finish_break
                
            If Int(t1 * 100) Mod 100 = 0 Then
                Label10.Caption = Board(t_1main, t_1pit)
            End If
                
                
                'сохраняемся каждые 30сек
                'If Int(t1 * 100) Mod 3000 = 0 Then
                '    Open App.Path & "\backup.txt" For Output As #21
                '    Print #21, Combo1.Text & vbCrLf; 'название 1й команды в пите
                '    Print #21, Label7.Caption & vbCrLf; 'счет 1й команды в пите
                '    Print #21, Label9.Caption & vbCrLf; 'время команд в пите
                '    Print #21, Label8.Caption & vbCrLf; 'счет 2й команды в пите
                '    Print #21, Combo2.Text & vbCrLf; 'название 2й команды в пите
                '    Print #21, Combo3.Text & vbCrLf; 'название 1й команды в поле
                '    Print #21, Combo4.Text & vbCrLf; 'название 2й команды в поле
                '    Print #21, Label1.Caption & vbCrLf; 'счет 1й команды в поле
                '    Print #21, Label2.Caption & vbCrLf; 'счет 2й команды в поле
                '    Print #21, Command7.BackColor & vbCrLf; 'признак timeout 1й команды в поле
                '    Print #21, Command6.BackColor & vbCrLf; 'признак timeout 2й команды в поле
                '    Print #21, Finish & vbCrLf;
                '    Print #21, Finish_main & vbCrLf;
                '    Print #21, overtime & vbCrLf;
                '
                '    Print #21, t1 & vbCrLf;
                '    Close #21
                'End If
            
            Label3.Caption = TimeInLabel(Start - Timer - Finish)
            Form3.Label3.Caption = Label3.Caption
            'гудим когда 60 сек dtwm
            'If 6105 > Int(t1 * 100) And Int(t1 * 100) > 6100 Then
            'Call Sound_dtmw
            'End If
            
            
            If 30000 > Int(t1 * 100) And Int(t1 * 100) > 6100 And Int(t1 * 100) Mod 1000 = 0 Then
                If license = 0 Then
                    PushButton = MsgBox("DEMO VERSION", 48)
                End If
            End If
            
            If 6000 > Int(t1 * 100) And Int(t1 * 100) > 5995 Then
            'If Label3.Caption = "0.59,00" Then
            Call Sound_60
            End If
            If t1 < 0.01 Then
            'If Label3.Caption = "0.00,01" Then
                'a = a + 1
                Finish = Finish_main
                Label3.Caption = "0.00,00"
                Form3.Label3.Caption = "0.00,00"
                Switch = False
                Call Sound_dtmw_stop
                Sleep 100
                Call Sound_time
                
                'показываем кнопку окончание матча
                Command5.Visible = True
                'если меньше 60 секунд осталось делаем видимой кнопку overtime
                If overtime = 0 Then
                Command41.BackColor = color_unactive
                End If
                'Command41.Visible = True
                If team_1 = team_2 Then
                    Command41.Visible = False
                End If
                
                Call Command4_Click
                
                'если счет равный запускаем овертайм
                'If team_1 = team_2 Then
                'Call Command41_Click
                'End If
            End If
            DoEvents
          'sleep(1000)
        Wend
    Else
        Finish = -Start + Finish + Timer
        'если количество нажатий на кнопку четное - выключаем секундомер
        Switch = False
        'меняем надпись на кнопке на "Пуск"
        Command42.Visible = True
        Command4.Caption = "Поле в игре"
        Command11.BackColor = color_unactive
        Command11.Caption = "Поле в игре"
        'показываем кнопку окончание матча
        Command5.Visible = True
        'гудим
    End If
End If
End Sub

'кнопка "Старт - стоп" break
Sub Command12_Click()

Label10.Caption = Board_team(t_1main, t_1pit)

Command41.Caption = "Overtime"
'List1.AddItem "1-" & Int((Start - Timer - Finish_break) * 100)
If Command11.Caption = "Поле в игре" Then
Else
If Label3.Caption <> "0.00,00" Then
If a Mod 2 = 0 Then
    'увеличиваем ко-во нажатий на кнопку на 1
    a_break = a_break + 1
    'если значение нечетное, включаем секундомер
    'показываем кнопку Старт Пит-Стоп
    Command12.Visible = True
    If a_break Mod 2 <> 0 Then
        'уведомляем о невозможности сделать настройки
        Text8.BackColor = &HFFFFFF
        Text9.BackColor = &HFFFFFF
        Text2.BackColor = &HFFFFFF
        Text14.BackColor = &HFFFFFF
        Call Sound_dtmw
        Switch = True
        'Меняем надпись на кнопке на "Стоп"
        Command12.Caption = "Стоп"
        Command5.Caption = "Окончание матча"
        Command11.BackColor = color_active
        'зеленим кнопки команд
        Command20.BackColor = &HFF00&
        Command21.BackColor = &HFF00&
        'выводим название команд в кнопки присвоения гейма
        Command16.Caption = "Гейм команды" & vbCrLf & Combo3.Text
        Command14.Caption = "Гейм команды" & vbCrLf & Combo4.Text
        'Finish_break = Finish_break - 0.01
        Start = Timer
        While Switch
            'выводим в метку текущее показание секундомера
            t1 = Start - Timer - Finish_break
            
            t_1pit = t1
            t_1main = -Finish
            
            Label4.Caption = TimeInLabel(Start - Timer - Finish_break)
            Form3.Label4.Caption = Label4.Caption
            'If 9005 > Int(t1 * 100) And Int(t1 * 100) > 9000 Then
            'Call Sound_dtmw
            'End If
            
            If 6000 > Int(t1 * 100) And Int(t1 * 100) > 1100 And Int(t1 * 100) Mod 500 = 0 Then
                If license = 0 Then
                    PushButton = MsgBox("DEMO VERSION", 48)
                End If
            End If
            
            If Int(t1 * 100) Mod 100 = 0 Then
                Label10.Caption = Board(t_1main, t_1pit)
            End If
            
            
            
            If 1100 > Int(t1 * 100) And Int(t1 * 100) > 1095 Then
            Call Sound_dtmw
            End If
            If 1000 > Int(t1 * 100) And Int(t1 * 100) > 995 Then
            Call Sound_10
            End If
            If 3100 > Int(t1 * 100) And Int(t1 * 100) > 3095 Then
            Call Sound_dtmw
            End If
            If 3000 > Int(t1 * 100) And Int(t1 * 100) > 2995 Then
            Call Sound_30
            End If
            If 6100 > Int(t1 * 100) And Int(t1 * 100) > 6095 Then
            Call Sound_dtmw
            End If
            If 6000 > Int(t1 * 100) And Int(t1 * 100) > 5995 Then
            Call Sound_minute
            End If
             
            If 955 > Int(t1 * 100) And Int(t1 * 100) > 950 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 900 > Int(t1 * 100) And Int(t1 * 100) > 895 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 855 > Int(t1 * 100) And Int(t1 * 100) > 850 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 800 > Int(t1 * 100) And Int(t1 * 100) > 795 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 755 > Int(t1 * 100) And Int(t1 * 100) > 750 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 700 > Int(t1 * 100) And Int(t1 * 100) > 695 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 655 > Int(t1 * 100) And Int(t1 * 100) > 650 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 600 > Int(t1 * 100) And Int(t1 * 100) > 595 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 555 > Int(t1 * 100) And Int(t1 * 100) > 550 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 500 > Int(t1 * 100) And Int(t1 * 100) > 495 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 455 > Int(t1 * 100) And Int(t1 * 100) > 450 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 400 > Int(t1 * 100) And Int(t1 * 100) > 395 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 355 > Int(t1 * 100) And Int(t1 * 100) > 350 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 300 > Int(t1 * 100) And Int(t1 * 100) > 295 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 255 > Int(t1 * 100) And Int(t1 * 100) > 250 Then
            second_counter = 0
            End If
            
            If second_counter = 0 Then
            If 200 > Int(t1 * 100) And Int(t1 * 100) > 195 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 155 > Int(t1 * 100) And Int(t1 * 100) > 150 Then
            second_counter = 0
            End If
                        
            If second_counter = 0 Then
            If 100 > Int(t1 * 100) And Int(t1 * 100) > 95 Then
            If Command34.BackColor = color_active Then
            Call Sound_seconds
            second_counter = second_counter + 1
            End If
            End If
            End If
            
            If 55 > Int(t1 * 100) And Int(t1 * 100) > 50 Then
            second_counter = 0
            End If
                     
            If t1 < 0.001 Then
            second_counter = 0
            If Combo1.Text = "" And Combo2.Text = "" Then
            Combo1.Visible = False
            Combo2.Visible = False
            Label9.Caption = "2 пита"
            Text5.Visible = False
            Label7.Visible = False
            Label8.Visible = False
            Command5.Visible = False
            'кнопка теста не активна
            Command44.BackColor = color_unactive
            
            If Combo1.Text <> "" Or Combo2.Text <> "" Then
                    Else
                    i1 = 1
            End If
            
            End If

                Call Sound_buzzer
                Switch = False
                a_break = a_break + 1
                If Command10.BackColor = color_active Then
                Finish_break = -30
                Finish_main_break = -30
                End If
                If Command9.BackColor = color_active Then
                Finish_break = -60
                Finish_main_break = -60
                End If
                If Command8.BackColor = color_active Then
                Finish_break = -120
                Finish_main_break = -120
                End If
                Finish_break = Finish_main_break
                Command12.Caption = "Старт"
                Command4.Caption = "Поле в игре"
                If Combo1.Text = "" And Combo2.Text = "" Then
                Finish_break = -120
                Finish_main_break = -120
                End If
                Label4.Caption = TimeInLabel(-Finish_break)
                Form3.Label4.Caption = TimeInLabel(-Finish_break)
                Call Command4_Click
            End If
            DoEvents
        Wend
    Else
        Finish_break = -Start + Finish_break + Timer
        Call Sound_dtmw
        'если количество нажатий на кнопку четное - выключаем секундомер
        Switch = False
        'меняем надпись на кнопке на "Пуск"
        Command12.Caption = "Старт Пит-Стоп"
    End If
End If
End If
End If
End Sub

'обработка нажатия на кнопку "окончание матча"
Sub Command5_Click()
If a_break Mod 2 = 0 Then
    If a Mod 2 = 0 Then
        If i1 < 2 Then
            PushButton = MsgBox("Подтверждаете окончание матча " & Combo3.Text & " и " & Combo4.Text & "?", 292)
            If PushButton = 6 Then
                If Label3.Caption = "0.00,00" And team_1 <> team_2 Then
                Else
                    
                    PushButton = MsgBox("Техническое поражение" & Combo3.Text & "?", 292)
                    If PushButton = 6 Then
                        team_1 = 0
                        team_2 = score_end

                            If result_info_counter = 0 Then
                                        'если первая пара играет
                                        result_info_2 = result_info_2 & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                                        Else
                                        'если вторая пара играет
                                        result_info = result_info & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                            End If
                        Else
                        PushButton = MsgBox("Техническое поражение" & Combo4.Text & "?", 292)
                        If PushButton = 6 Then
                            team_2 = 0
                            team_1 = score_end
                            
                            If result_info_counter = 0 Then
                                        'если первая пара играет
                                        result_info_2 = result_info_2 & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                                        Else
                                        'если вторая пара играет
                                        result_info = result_info & Combo3.Text & " | " & team_1 & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                            End If
                        
                            Else
                    
                            PushButton = MsgBox("Гейм присуждается:" & Combo3.Text & "?", 292)
                            If PushButton = 6 Then
                                   team_1 = team_1 + 1
                                   Label1.Caption = team_1
                                   Form3.Label1.Caption = team_1
                                    If result_info_counter = 0 Then
                                        'если первая пара играет
                                        result_info_2 = result_info_2 & Combo3.Text & " | " & team_1 & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                                        Else
                                        'если вторая пара играет
                                        result_info = result_info & Combo3.Text & " | " & team_1 & " - " & Label2.Caption & " | " & Combo4.Text & vbCrLf
                                    End If
                            
                            Else
                                    PushButton = MsgBox("Гейм присуждается:" & Combo4.Text & "?", 292)
                                   If PushButton = 6 Then
                                       team_2 = team_2 + 1
                                        Label2.Caption = team_2
                                        Form3.Label2.Caption = team_2
                                        If result_info_counter = 0 Then
                                            'если первая пара играет
                                            result_info_2 = result_info_2 & Combo3.Text & " | " & Label1.Caption & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                                            Else
                                            'если вторая пара играет
                                            result_info = result_info & Combo3.Text & " | " & Label1.Caption & " - " & team_2 & " | " & Combo4.Text & vbCrLf
                                        End If
                            
                                   End If
                            End If
                        End If
                    End If
                End If
                               
                Call Command50_Click
                If i1 < 2 Then
                    Call Command12_Click
                End If
            End If
        Else
        Call Command50_Click
        Call Command12_Click
        End If
    End If
End If
End Sub

'функция секундомера
Function TimeInLabel(t As Single)
    Dim Value As Integer
    Value = Int(t / 60)
    TimeInLabel = CStr(Value) + "."
    t = t - Value * CSng(60)
    Value = Int(t)
    TimeInLabel = TimeInLabel + Format(Value, "00") + ","
    t = t - Value
    Value = Int(t * 100)
    TimeInLabel = TimeInLabel + Format(Value, "00")
End Function

'функция вывод информации для табло команды и счет
Function Board_team(t12_main As Single, t12_pit As Single)
    'режем название до 8 символов и меняем шрифт для табло
    k1_board = translate(Left(Combo3.Text & "        ", 8))
    k2_board = translate(Left(Combo4.Text & "        ", 8))
    'смотрим счет
    k1_score_board = Label1.Caption
    k2_score_board = Label2.Caption
    'отправляем посылки подряд
    If Combo9.Text = "" Or Combo9.Text = "Табло в поле" Then
    Else
    Board_team = "K1" & k1_board & k1_score_board & vbCrLf
    SComm2.Output = Board_team
    Board_team = "K2" & k2_board & k2_score_board & vbCrLf
    SComm2.Output = Board_team
    End If
End Function


'функция вывод информации для табло
Function Board(t12_main As Single, t12_pit As Single)
    'смотрим основное время
    Value = Int(t12_main / 60)
    If Value < 10 Then
        'добавляем ноль если знак один
        t1_board = "0" & CStr(Value) + ":"
    Else
        t1_board = CStr(Value) + ":"
    End If
    t12_main = t12_main - Value * CSng(60)
    Value = Int(t12_main)
    t1_board = t1_board + Format(Value, "00")
    'смотрим время пита
        Value = Int(t12_pit / 60)
    If Value < 10 Then
        'добавляем ноль если знак один
        t2_board = "0" & CStr(Value) + ":"
    Else
        t2_board = CStr(Value) + ":"
    End If
    t12_pit = t12_pit - Value * CSng(60)
    Value = Int(t12_pit)
    t2_board = t2_board + Format(Value, "00")
    'отправляем посылки подряд
    If Combo9.Text = "" Or Combo9.Text = "Табло в поле" Then
    Else
    Board = "T1" & t1_board & vbCrLf
    SComm2.Output = Board
    Board = "T2" & t2_board & vbCrLf
    SComm2.Output = Board
    End If
End Function

'функция подготовки в русской и английской кодировке для табло
Function translate(board_translate As String)
Text17.Text = board_translate
                'меняем английский на заглавные
                Text17.Text = Replace(Text17.Text, "a", "A")
                Text17.Text = Replace(Text17.Text, "b", "B")
                Text17.Text = Replace(Text17.Text, "c", "C")
                Text17.Text = Replace(Text17.Text, "d", "D")
                Text17.Text = Replace(Text17.Text, "e", "E")
                Text17.Text = Replace(Text17.Text, "f", "F")
                Text17.Text = Replace(Text17.Text, "g", "G")
                Text17.Text = Replace(Text17.Text, "h", "H")
                Text17.Text = Replace(Text17.Text, "i", "I")
                Text17.Text = Replace(Text17.Text, "j", "J")
                Text17.Text = Replace(Text17.Text, "k", "K")
                Text17.Text = Replace(Text17.Text, "l", "L")
                Text17.Text = Replace(Text17.Text, "m", "M")
                Text17.Text = Replace(Text17.Text, "n", "N")
                Text17.Text = Replace(Text17.Text, "o", "O")
                Text17.Text = Replace(Text17.Text, "p", "P")
                Text17.Text = Replace(Text17.Text, "q", "Q")
                Text17.Text = Replace(Text17.Text, "r", "R")
                Text17.Text = Replace(Text17.Text, "s", "S")
                Text17.Text = Replace(Text17.Text, "t", "T")
                Text17.Text = Replace(Text17.Text, "u", "U")
                Text17.Text = Replace(Text17.Text, "v", "V")
                Text17.Text = Replace(Text17.Text, "w", "W")
                Text17.Text = Replace(Text17.Text, "x", "X")
                Text17.Text = Replace(Text17.Text, "y", "Y")
                Text17.Text = Replace(Text17.Text, "z", "Z")
                'меняем все используемые символы на *
                Text17.Text = Replace(Text17.Text, "*", "@")
                Text17.Text = Replace(Text17.Text, "~", "@")
                Text17.Text = Replace(Text17.Text, "\", "@")
                Text17.Text = Replace(Text17.Text, ">", "@")
                Text17.Text = Replace(Text17.Text, "<", "@")
                Text17.Text = Replace(Text17.Text, "|", "@")
                Text17.Text = Replace(Text17.Text, "=", "@")
                Text17.Text = Replace(Text17.Text, "$", "@")
                'менсем русские на английский
                Text17.Text = Replace(Text17.Text, "А", "A")
                Text17.Text = Replace(Text17.Text, "Б", "n")
                Text17.Text = Replace(Text17.Text, "В", "B")
                Text17.Text = Replace(Text17.Text, "Г", "g")
                Text17.Text = Replace(Text17.Text, "Д", "d")
                Text17.Text = Replace(Text17.Text, "Е", "E")
                Text17.Text = Replace(Text17.Text, "Ё", "*")
                Text17.Text = Replace(Text17.Text, "Ж", "j")
                Text17.Text = Replace(Text17.Text, "З", "z")
                Text17.Text = Replace(Text17.Text, "И", "i")
                Text17.Text = Replace(Text17.Text, "Й", "~")
                Text17.Text = Replace(Text17.Text, "К", "K")
                Text17.Text = Replace(Text17.Text, "Л", "l")
                Text17.Text = Replace(Text17.Text, "М", "M")
                Text17.Text = Replace(Text17.Text, "Н", "H")
                Text17.Text = Replace(Text17.Text, "О", "O")
                Text17.Text = Replace(Text17.Text, "П", "r")
                Text17.Text = Replace(Text17.Text, "Р", "\")
                Text17.Text = Replace(Text17.Text, "С", "C")
                Text17.Text = Replace(Text17.Text, "Т", "T")
                Text17.Text = Replace(Text17.Text, "У", "y")
                Text17.Text = Replace(Text17.Text, "Ф", "f")
                Text17.Text = Replace(Text17.Text, "Х", "X")
                Text17.Text = Replace(Text17.Text, "Ц", "s")
                Text17.Text = Replace(Text17.Text, "Ч", "u")
                Text17.Text = Replace(Text17.Text, "Ш", "v")
                Text17.Text = Replace(Text17.Text, "Щ", "w")
                Text17.Text = Replace(Text17.Text, "Ъ", ">")
                Text17.Text = Replace(Text17.Text, "Ь", "<")
                Text17.Text = Replace(Text17.Text, "Ы", "|")
                Text17.Text = Replace(Text17.Text, "Э", "=")
                Text17.Text = Replace(Text17.Text, "Ю", "$")
                Text17.Text = Replace(Text17.Text, "Я", "q")
                Text17.Text = Replace(Text17.Text, "а", "A")
                Text17.Text = Replace(Text17.Text, "б", "n")
                Text17.Text = Replace(Text17.Text, "в", "B")
                Text17.Text = Replace(Text17.Text, "г", "g")
                Text17.Text = Replace(Text17.Text, "д", "d")
                Text17.Text = Replace(Text17.Text, "е", "E")
                Text17.Text = Replace(Text17.Text, "ё", "*")
                Text17.Text = Replace(Text17.Text, "ж", "j")
                Text17.Text = Replace(Text17.Text, "з", "z")
                Text17.Text = Replace(Text17.Text, "и", "i")
                Text17.Text = Replace(Text17.Text, "й", "~")
                Text17.Text = Replace(Text17.Text, "к", "K")
                Text17.Text = Replace(Text17.Text, "л", "l")
                Text17.Text = Replace(Text17.Text, "м", "M")
                Text17.Text = Replace(Text17.Text, "н", "H")
                Text17.Text = Replace(Text17.Text, "о", "O")
                Text17.Text = Replace(Text17.Text, "п", "r")
                Text17.Text = Replace(Text17.Text, "р", "\")
                Text17.Text = Replace(Text17.Text, "с", "C")
                Text17.Text = Replace(Text17.Text, "т", "T")
                Text17.Text = Replace(Text17.Text, "у", "y")
                Text17.Text = Replace(Text17.Text, "ф", "f")
                Text17.Text = Replace(Text17.Text, "х", "X")
                Text17.Text = Replace(Text17.Text, "ц", "s")
                Text17.Text = Replace(Text17.Text, "ч", "u")
                Text17.Text = Replace(Text17.Text, "ш", "v")
                Text17.Text = Replace(Text17.Text, "щ", "w")
                Text17.Text = Replace(Text17.Text, "ъ", ">")
                Text17.Text = Replace(Text17.Text, "ь", "<")
                Text17.Text = Replace(Text17.Text, "ы", "|")
                Text17.Text = Replace(Text17.Text, "э", "=")
                Text17.Text = Replace(Text17.Text, "ю", "$")
                Text17.Text = Replace(Text17.Text, "я", "q")
    'пишем результат
    translate = Text17.Text
End Function

'ищем доступные com порты
Private Sub PopulateList()
      
      '// In this example we'll test for ports up to 32 - you can test up to 255 if you want.
       For i = 1 To 32
              SComm1.CommPort = i
              If SComm1.CommName = "" Then
                     '// This Com Port does not exist at all
                  Else
                     '// This port exists so show the DeviceName in the list
                     'Combo7.AddItem SComm1.CommName  ''временно отключаем для демоверсии
                     'Combo9.AddItem SComm1.CommName  ''временно отключаем для демоверсии
                     '// Also store i so that when the user selects one we'll know which port to open
                     'Combo7.ItemData(Combo7.NewIndex) = i ''временно отключаем для демоверсии
                     'Combo9.ItemData(Combo9.NewIndex) = i ''временно отключаем для демоверсии
              End If
       Next i
End Sub

'обрабатываем com порт
Private Sub SComm1_OnComm()
    If (SComm1.CommEvent = comEvReceive) Then
            Buffer = SComm1.Input
            Text7.Text = Buffer
            
        'если нажат тест то запоминаем идентификатор комплекта
        If Command44.BackColor = &HFF& Then
            'oldstyle   If InStr(3, Text7.Text, "8A") > 0 Or InStr(3, Text7.Text, "8B") > 0 Or InStr(3, Text7.Text, "B5") > 0 Or Text7.Text = "8A" Or Text7.Text = "B5" Then
                      
            'если приходит посылка от кнопки
            'проверяем что это посылка - A банер, B Пит, С пульт
            
            'записываем кнопки банера
            If Left(Text7.Text, 1) = "A" Then
                    'запоминаем идентификатор комплекта
                    button_count = 0
                    'признак наличия этого идентификатора в списке
                    button_yes = 0
                    Do Until (button_count > Combo8.ListCount)
                        If Combo8.List(button_count) = Right(Text7.Text, 2) Then
                            'если совпадение найдено то взводим флаг
                            button_yes = 1
                        End If
                        button_count = button_count + 1
                    Loop
                        If button_yes = 1 Then
                            Else
                            'записываем значение для возможности выбора через меню
                            Combo8.AddItem Right(Text7.Text, 2)
                            Combo10.AddItem Right(Text7.Text, 2)
                            button_yes = 0
                        End If
                    Call Sound_dtmw_stop
                    SComm1.PortOpen = False
                    Sleep (1000)
                    SComm1.PortOpen = True
                    license = 1
            End If
            
            'записываем кнопки пита
            If Left(Text7.Text, 1) = "B" Then
                    'запоминаем идентификатор комплекта
                    button_count = 0
                    'признак наличия этого идентификатора в списке
                    button_yes = 0
                    Do Until (button_count > Combo11.ListCount)
                        If Combo11.List(button_count) = Right(Text7.Text, 2) Then
                            'если совпадение найдено то взводим флаг
                            button_yes = 1
                        End If
                        button_count = button_count + 1
                    Loop
                        If button_yes = 1 Then
                            Else
                            'записываем значение для возможности выбора через меню
                            Combo11.AddItem Right(Text7.Text, 2)
                            Combo12.AddItem Right(Text7.Text, 2)
                            button_yes = 0
                        End If
                    Call Sound_seconds
                    SComm1.PortOpen = False
                    Sleep (1000)
                    SComm1.PortOpen = True
                    license = 1
            End If
            
            'записываем пульт старшего судьи поля
            If Left(Text7.Text, 1) = "C" Then
                    'запоминаем идентификатор комплекта
                    button_count = 0
                    'признак наличия этого идентификатора в списке
                    button_yes = 0
                    Do Until (button_count > Combo13.ListCount)
                        If Combo13.List(button_count) = Right(Text7.Text, 2) Then
                            'если совпадение найдено то взводим флаг
                            button_yes = 1
                        End If
                        button_count = button_count + 1
                    Loop
                        If button_yes = 1 Then
                            Else
                            'записываем значение для возможности выбора через меню
                            Combo13.AddItem Right(Text7.Text, 2)
                            button_yes = 0
                        End If
                    Call Sound_seconds
                    SComm1.PortOpen = False
                    Sleep (1000)
                    SComm1.PortOpen = True
                    license = 1
            End If
            
        End If
        'закончили прописывание кнопок
       
        'действие на нажатие кнопки на банере A
        If Text7.Text = "A" & Combo8.Text Then 'первая кнопка\её код '8A
            license = 1
            Call Command20_Click
        End If
        
        'действие на нажатие кнопки на банере B
        If Text7.Text = "A" & Combo10.Text Then 'вторая кнопка\её код 'B5
            license = 1
            Call Command21_Click
        End If
        
        'действие на нажатие кнопки в пите A
        If Text7.Text = "B" & Combo11.Text Then 'вторая кнопка\её код 'B5
            license = 1
            'таймаут
            Call Command7_Click
            'полотенце
            If a_break Mod 2 = 0 Then
                If a Mod 2 > 0 Then
                'останавливаем таймер
                Call Command20_Click
                'красим кнопки банеров в красный чтобы понять что полотенце
                Command20.BackColor = &HFF&
                Command21.BackColor = &HFF&
                'красим кто нажал полотенце
                Command16.BackColor = &HFF&
                'скрываем кнопку продолжить гейм
                Command11.Visible = False
                'скрываем кнопку гейм без очка
                Command42.Visible = False
                
                End If
            End If
        End If
        
        'действие на нажатие кнопки в пите B
        If Text7.Text = "B" & Combo12.Text Then 'вторая кнопка\её код 'B5
            license = 1
            'таймаут
            Call Command6_Click
            'полотенце
            If a_break Mod 2 = 0 Then
                If a Mod 2 > 0 Then
                'останавливаем таймер
                Call Command21_Click
                'красим кнопки банеров в красный чтобы понять что полотенце
                Command20.BackColor = &HFF&
                Command21.BackColor = &HFF&
                'красим кто нажал полотенце
                Command14.BackColor = &HFF&
                'скрываем кнопку продолжить гейм
                Command11.Visible = False
                'скрываем кнопку гейм без очка
                Command42.Visible = False
                
                End If
            End If
        End If
        
        Text7.Text = ""
      End If
End Sub

'выбираем ком порт приемника кнопок для работы
' временно отключаем для демо версии, открыть заремленые строки
Private Sub combo7_Click()
       '// In this example we'll test for ports up to 32 - you can test up to 255 if you want.
       ' If SComm1.PortOpen = True Then
       '     SComm1.PortOpen = False
       ' End If
       ' If Combo7.Text <> Combo9.Text Then
       '     SComm1.CommPort = Combo7.ItemData(Combo7.ListIndex)
       '     SComm1.PortOpen = True
       ' Else
       '  Combo7.Text = "Выберите порт подключения приёмника"
       ' End If
        'блокируем изменение порта
End Sub

'выбираем ком порт табло для работы
Private Sub combo9_Click()
       ''// In this example we'll test for ports up to 32 - you can test up to 255 if you want.
        'If SComm2.PortOpen = True Then
        '    SComm2.PortOpen = False
        'End If
        'If Combo9.Text <> Combo7.Text Then
        '    SComm2.CommPort = Combo9.ItemData(Combo9.ListIndex)
        '    SComm2.PortOpen = True
        'Else
        ' Combo9.Text = "Табло в поле"
        'End If
        'блокируем изменение порта
End Sub
