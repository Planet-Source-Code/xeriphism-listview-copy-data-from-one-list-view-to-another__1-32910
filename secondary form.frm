VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   Caption         =   "Deliv"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Order"
      Height          =   855
      Left            =   6720
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   855
      Left            =   6720
      TabIndex        =   1
      Top             =   6240
      Width           =   1815
   End
   Begin ComctlLib.ListView List2 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   13361
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End ' how to exit project.
End Sub

Private Sub Command2_Click()
Unload Me ' how to unload a form.
End Sub

Private Sub Form_Load()
Dim clmx As ColumnHeader ' still have to define clmx as going to redef a lv..
Dim listitem As listitem ' and listitem...
Dim i As Integer         ' my counter variable i
If fromorder = True Then ' ok, the tricky part, if you clicked comm2...
i = 1 ' counter initialised..
Set clmx = Form2.List2.ColumnHeaders.Add(, , Form1.lvwList1.ColumnHeaders(1), List2.Width / 3) ' this just copies...
Set clmx = Form2.List2.ColumnHeaders.Add(, , Form1.lvwList1.ColumnHeaders(2), List2.Width / 3)
Set clmx = Form2.List2.ColumnHeaders.Add(, , Form1.lvwList1.ColumnHeaders(3), List2.Width / 3)
'Set clmx = Form2.List2.ColumnHeaders.Add(, , Form1.lvwList1.ColumnHeaders(4), List2.Width / 3) ' dropped one of the clmxs!! hah

While i < Form1.lvwList1.ListItems.Count + 1 ' and while it's not(counter that is) = to the total count of items from lv1...
Set listitem = Form2.List2.ListItems.Add(, , Form1.lvwList1.ListItems(i).Text)  ' add a listitem
listitem.SubItems(1) = Form1.lvwList1.ListItems(i).SubItems(1)                  ' add a subitem
listitem.SubItems(2) = Form1.lvwList1.ListItems(i).SubItems(2)                  ' add another subitem..
i = i + 1 ' increment counter
Wend
End If

If fromorder = False Then 'otherwise...
i = 1 ' reinitialise counter..
Set clmx = List2.ColumnHeaders.Add(, , "BITTEN!", List2.Width / 2) 'initialise lv2 as a seperate diff idea altogether
Set clmx = List2.ColumnHeaders.Add(, , "SHY!", List2.Width / 2)    '

While i < Form1.lvwList1.ListItems.Count ' going to add this data as many times as it appears in the other table :)
Set listitem = List2.ListItems.Add(, , "once") ' and be once bitten, twice shy...
listitem.SubItems(1) = "twice"
' as saddened as i am.. the next two lines failed me...
'Set listitem = Form2.List2.ListItems.Add(, , Form1.lvwList1.ListItems(i).Text)   '*** < NB you can't do this successfully! :(
'listitem.SubItems(1) = Form1.lvwList1.ListItems(i).SubItems(1)                   '*** < --- NOR THIS!!!
i = i + 1
Wend
End If
End Sub
