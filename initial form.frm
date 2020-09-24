VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Order"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exit to Deliv no copy"
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Listview to DELIV"
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin ComctlLib.ListView lvwList1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11880
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Dim listitem As listitem ' initialise listitem..
strall = strall + 1
' with listview you have to set the first item seperately
' the rest are sub items ie

''''''''''''''''''''''
' say this is ur row... ok
'_____________________
'| 1  | s1 | s2 | s3 |
'|____|____|____|____|
' 1 = item
' s1 = subitem 1
' s2 = subitem 2
' s3 = subitem 3
' this is how you have to code for it... so...

Set listitem = lvwList1.ListItems.Add(, , strall)
    ' sets up item... ie 1
    
    listitem.SubItems(1) = "BlERGh"
    ' sets up subitem 1
    
    listitem.SubItems(2) = "hGrelB"
    ' sets up subitem 2

End Sub

Private Sub Command2_Click()
fromorder = True ' ok doing something a bit tricky here... read more on form 2's onload

'                         how to load a new form to the screen...
'                                       ||       || NB!
'                                       \/ _     \/
Load Form2   ' loads form 2 into memory...  \_  you need both to do this...
Form2.Show 1 ' shows form 2 on the screen  _/

End Sub


Private Sub Command3_Click()
End ' command to actually quit the program.. don't do that unload bull when u want to actually end ur program.
End Sub

Private Sub Command4_Click()
fromorder = False ' as in command 2 doing something a bit tricky, refer to form 2's onload

'                         how to load a new form to the screen...
'                                       ||       || NB!
'                                       \/ _     \/
Load Form2   ' loads form 2 into memory...  \_  you need both to do this...
Form2.Show 1 ' shows form 2 on the screen  _/
End Sub

Private Sub Form_Load()
'*'*'*'*'*'*'*'*'*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' How utterly useless i hear you say...                      '
' however, i dare you to make an app that can                '
' port listview to listview with differing                   '
' columnheaders.  That would be the way to                   '
' surpass my excellent little bit of code                    '
' all this does basically is...                              '
' copy the listview data from one box to                     '
' another listview on another form.                          '
'*'*'***'**''''*'*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
''*''***'*'**''*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
'*'*'***'*'*'*'*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'''
' the only part that is almost clever is...                  '
' ok so you want to have differing columnheaders?            '
' well enter the same data                                   '
' then reset the widths of the cols you do not want to see   '
' to zero                                                    '
' hence avoided this 'problem' i have.                       '
'*'*'*'*'*'*'*'*'*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         p.s listen to 'predominance - godflesh'            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*'*'*'*'*'*'*'*'*'*'*'*'*'*''*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'

'ok first thing to be aware of with a listview is setting up the columnheaders
'you can do it pretty much anywhere...
'follow this syntax only replace lvwlist1.columnheaders.add with X.columnheaders.add
'where x is the name of any listview on your form, assuming u may have multiples...
' now, the oversight in this code is that it doesn't properly copy the listview'
' as for some reason you can only copy data into the same type of listview with the same columnheaders
' but notice what is written above.. when you set the width... set it to zero for something you don't want to see
' sure it means both listviews have almost identical code BUT... one is just referencing the other...
' syntax below for this....
' set variable you set for columnheader = listview.columnheaders.add(,,"title",width)
Dim clmx As ColumnHeader ' i use clmx as variable.....
Set clmx = lvwList1.ColumnHeaders.Add(, , "hah", lvwList1.Width / 3) ' then i add hah
Set clmx = lvwList1.ColumnHeaders.Add(, , "B", lvwList1.Width / 3)   ' and B
Set clmx = lvwList1.ColumnHeaders.Add(, , "A", lvwList1.Width / 3)   ' and A
' being a bit tricky, see if u can see what i do....
End Sub
