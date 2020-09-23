VERSION 5.00
Begin VB.Form FileDownload 
   Caption         =   "Downloads ANY file to your computer from anywhere"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox HTMLSource 
      Height          =   2715
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1830
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Height          =   525
      Left            =   4260
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox MyLocalFile 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Text            =   "C:\microsoft_home_page.html"
      Top             =   540
      Width           =   4545
   End
   Begin VB.TextBox MyURL 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Text            =   "http://www.microsoft.com"
      Top             =   210
      Width           =   4545
   End
   Begin VB.Label Label3 
      Caption         =   "Source of File:"
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   1470
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Save As:"
      Height          =   285
      Left            =   210
      TabIndex        =   2
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "URL of File:"
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "FileDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MsgBox "I will now attempt to download the file. If the file is large, the program may freeze, but the program will unfreeze after the file has finished downloading.", vbInformation, "Ready? Read and the Click OK"

Call DownloadFile(MyURL, MyLocalFile) ' Downloads..

' Opens the source onto the screen. Optional
On Error Resume Next
Open MyLocalFile For Input As #1
HTMLSource = Input(LOF(1), 1)
Close #1

MsgBox "Download complete." & vbNewLine & vbNewLine & "To open your file, go here: " & MyLocalFile, vbInformation, "Downloaded"


End Sub
