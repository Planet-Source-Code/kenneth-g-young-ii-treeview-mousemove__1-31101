VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function addTreeList()

    Dim tempNode As Node
    Dim counter As Integer
    Dim sTemp As String
    Dim tmpFile As String
    
    tmpFile = App.Path & "\lang.txt"
    TV1.Nodes.Clear
    
    Set tempNode = TV1.Nodes.Add(, , "L", "Languages")
    
    Open tmpFile For Input As #1
    
    For counter = 1 To 7
        Line Input #1, sTemp
        
        Set tempNode = TV1.Nodes.Add("L", tvwChild, "L" & counter, sTemp)
    Next
    tempNode.EnsureVisible
    
    Close #1
    
    TV1.Style = tvwTreelinesText
    TV1.BorderStyle = ccFixedSingle
    TV1.ZOrder 0
End Function

Private Sub Form_Load()
    addTreeList
End Sub

Private Sub TV1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    
Dim oNode As Node
Set oNode = TV1.HitTest(x, y)
    
If Not oNode Is Nothing Then
    oNode.Selected = True
    txt1.Text = oNode.Text
End If
    

End Sub
