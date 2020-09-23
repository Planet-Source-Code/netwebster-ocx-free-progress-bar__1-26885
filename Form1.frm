VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "No OCX Progress Bar"
   ClientHeight    =   1440
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   3195
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Show Me"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      FillColor       =   &H00C00000&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblRecord 
      AutoSize        =   -1  'True
      Caption         =   "Record"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'* This program for the Percentage Progress Bar using a picture with no OCX. ***
'* You will need the UpdateProgress Function and a picture box on your    ******
'* form.  The sub below is an example for counting files.  Of course,     ******
'* you will need to adjust this function to suit your needs. I worked    *******
'* on the UpdateProgress Function to it's present state as it was     **********
'* originally compiled to measure time left to download - taken from VB5    ****
'* Developers Handbook.                                                   ******
'* By: K. Juryea                                                       ******
'* On: 12/02/00                                                           ******
'*******************************************************************************

Public Sub ProgressBar()
    
    '***** Declare Variables *****
    Dim CurrentPercent As Long
    Dim TotalRecs, Count As Long
    
    TotalRecs = 5000
    Count = 1 'to prevent division by 0
    Do While Count <= TotalRecs
    Text1.Refresh
    Text1.Text = Count
        If TotalRecs <= 99999 Then
            CurrentPercent = (Count / TotalRecs) * 100
        Else
            CurrentPercent = (Count / TotalRecs) * 50 'if there are more then 99999-adjust
        End If
        UpdateProgress pb, CurrentPercent
        Count = Count + 1
    Loop

End Sub


Public Function UpdateProgress(pb As Control, ByVal Percent)
    '***** Declare Variables *****
    Dim PercentNum As String

    If Not pb.AutoRedraw Then 'AutoRedraw must be true
        pb.AutoRedraw = -1
    End If

    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    PercentNum = Format$(Percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(PercentNum) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(PercentNum)) / 2
    pb.Print PercentNum
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh
    
End Function

Private Sub cmdTest_Click()
    Call ProgressBar
End Sub
