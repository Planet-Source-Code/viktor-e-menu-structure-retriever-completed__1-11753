VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu structure retriever"
   ClientHeight    =   5085
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox addmenut 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2970
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.TextBox newmenuc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7223
      _Version        =   393217
      Indentation     =   35
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texture from Corel PhotoPaint"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   150
      Left            =   4570
      TabIndex        =   12
      ToolTipText     =   "This sound courtesy of Paramount (:-)"
      Top             =   4905
      Width           =   1635
   End
   Begin VB.Image injos 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5520
      Picture         =   "Form1.frx":233BA
      Top             =   2340
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      Height          =   255
      Left            =   2970
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFFF&
      Height          =   255
      Left            =   4710
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label remmenit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remove this item"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4740
      TabIndex        =   10
      Top             =   1815
      Width           =   1245
   End
   Begin VB.Label addmenit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add a menu item"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   3010
      TabIndex        =   9
      Top             =   1815
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0E0FF&
      Height          =   375
      Left            =   3120
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label infol 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   3240
      TabIndex        =   8
      Top             =   2340
      Width           =   2535
   End
   Begin VB.Label tipmenl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label dameniub 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Populate"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Click here to retrieve this form's menu structure"
      Top             =   195
      Width           =   855
   End
   Begin VB.Image schmimag 
      Height          =   280
      Left            =   240
      Top             =   170
      Width           =   855
   End
   Begin VB.Label chmenul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change the caption of this item to:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label numeniul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label men 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   195
      Width           =   1335
   End
   Begin VB.Menu mprog 
      Caption         =   "&Program"
      Begin VB.Menu mpinf 
         Caption         =   "Informations"
         Begin VB.Menu mpiau 
            Caption         =   "Author"
         End
         Begin VB.Menu mpico 
            Caption         =   "Company"
            Begin VB.Menu mcicp 
               Caption         =   "Profile"
            End
            Begin VB.Menu mcicc 
               Caption         =   "Contact"
            End
         End
      End
      Begin VB.Menu mothe 
         Caption         =   "Other"
      End
      Begin VB.Menu mpaba 
         Caption         =   "Abandon"
      End
   End
   Begin VB.Menu medit 
      Caption         =   "&Edit"
      Begin VB.Menu mecop 
         Caption         =   "Copy"
      End
      Begin VB.Menu meste 
         Caption         =   "Eraser"
         Begin VB.Menu mescu 
            Caption         =   "Delete current"
         End
         Begin VB.Menu mesto 
            Caption         =   "Delete all"
         End
      End
      Begin VB.Menu meinl 
         Caption         =   "Replace"
      End
      Begin VB.Menu mecau 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu moper 
      Caption         =   "&Operations"
      Begin VB.Menu mosal 
         Caption         =   "Save"
      End
      Begin VB.Menu moinc 
         Caption         =   "Load"
      End
      Begin VB.Menu moinh 
         Caption         =   "Close"
         Begin VB.Menu mosdo 
            Caption         =   "Current document"
         End
         Begin VB.Menu mosto 
            Caption         =   "All opened files"
            Begin VB.Menu mocdd 
               Caption         =   "All documents"
            End
            Begin VB.Menu mocdi 
               Caption         =   "All images"
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Menu structure retriever
'Created and tested on a Win98SE system with VB6 Enterprise

'by Victor Stefanescu
'ActiveX developer
'gimelhai@mailcity.com
'09.2000

'Thanks and appreciations go to Corel for the texture    (:-)
'and to Paramount for Star Trek sounds                   (:-), too
'By the way, to all of you VB programmers: let's make -ONLY- COOL interfaces to our programs !...
'...and persuade Microsoft to provide with transparent background all their controls (:-)


Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" _
(ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'constants for menu item-referred APIs:
Const MF_APPEND = &H100& 'to append a menu item
Const MF_REMOVE = &H1000& 'to remove a menu item
Const MF_STRING = &H0& 'specifies the type of a menu item
Const MF_SEPARATOR = &H800& 'insert a separator menu item
Const MF_BYPOSITION = &H400& 'item specified by 0-level relative position
'constants for PlaySound:
Const SND_FILENAME = &H20000 'play a file
Const SND_ASYNC = &H1 'play the wave asynchronously

'God, how I hate to declare variables !
Dim MeniulForm, MeniuriPrincipale, SubItemMeniu, SubItemMeniu2, SubItemMeniu3, Meniu, Submeniu, SubSubMeniu, TotalItemuri, TotalSubItemuri, PozitieMenuItem As Long
Dim Cinta As Boolean, infos, gasitin As String

Private Sub Form_Load()
'prepair some interface elements:
Shape1.Top = infol.Top - 16: Shape1.Height = infol.Height + 20
Shape1.Left = infol.Left - 8: Shape1.Width = infol.Width + 24
'fill the infol label with some text:
infos = "This is an example of how to" & vbCrLf
infos = infos & "retrieve the menu structure of a" & vbCrLf
infos = infos & "Form using several API functions." & vbCrLf
infos = infos & "The code beyond isn't too" & vbCrLf
infos = infos & "dangerous, assuming one uses" & vbCrLf
infos = infos & "logic and common sense along" & vbCrLf
infos = infos & "with simple mathematics." & vbCrLf
infos = infos & "You may freely use this code" & vbCrLf
infos = infos & "and improve it to spell out" & vbCrLf
infos = infos & "more intricated menu structures." & vbCrLf
infos = infos & vbCrLf
infos = infos & "09.2000 Victor Stefanescu"
infol.Caption = infos
infol.ZOrder 0: injos.ZOrder 0
End Sub

Private Sub dameniub_Click()
'the "Populate" label
If dameniub.Enabled = True Then
TotalItemuri = 0
TotalSubItemuri = 0
MeniuriPrincipale = 0

MeniulForm = GetMenu(Me.hwnd)
MeniuriPrincipale = GetMenuItemCount(MeniulForm)

For Meniu = 0 To MeniuriPrincipale - 1
    SubItemMeniu = GetSubMenu(MeniulForm, Meniu)
    TotalItemuri = TotalItemuri + GetMenuItemCount(SubItemMeniu)
Next Meniu

men.Caption = MeniuriPrincipale & " main menus" & vbCrLf

DaNumele
dameniub.Enabled = False
End If
End Sub

Private Sub DaNumele()
'the procedure which populates the treeview according to current menu structure
Dim NumeObtinut, NodMeniu, NodSubmeniu As Long, NodSubSubMeniu As Long, NumeMeniu As String * 20, NumeSubMeniu As String * 20, NumeSubSubMeniu As String * 20, Nod As Node
Dim ComandaNivel0, ComandaNivel1, ComandaNivel2 As Long
Dim TotalSubSubItemuri, NumeObtinut4 As Long
tv.Nodes.Clear
addmenut.Visible = False: addmenit.Visible = True

'insert in treeview nodes identified by main menus' names and positions (Program, Edit, Operations)
For NodMeniu = 0 To MeniuriPrincipale - 1
    SubItemMeniu = GetSubMenu(MeniulForm, 0)
    NumeObtinut = GetMenuString(MeniulForm, NodMeniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
    Set Nod = tv.Nodes.Add(, , "0-" & "main menu " & NodMeniu, NumeMeniu)
Next NodMeniu

'insert in treeview nodes identified by main menus commands' names and positions (Program, Edit, Operations)
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    'now populating the main menu treeview nodes with these items:
    For Meniu = 0 To TotalItemuri - 1
        NumeObtinut = GetMenuString(ComandaNivel0, Meniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
        Set Nod = tv.Nodes.Add(NodMeniu + 1, tvwChild, NodMeniu & "-" & Meniu, NumeMeniu)
    Next Meniu
Next NodMeniu

'are there some submenus? if so, populate the treeview with them:
'start parsing form menus all over again
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    For NodSubmeniu = 0 To TotalItemuri - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
        If TotalSubItemuri > 0 Then
            gasitin = NodMeniu & "-" & NodSubmeniu '=key to identify treeview nodes
            'now populating the main menu treeview nodes with the submenus:
            For Submeniu = 0 To TotalSubItemuri - 1
                NumeObtinut = GetMenuString(ComandaNivel1, Submeniu, NumeSubMeniu, Len(NumeSubMeniu), MF_BYPOSITION)
                Set Nod = tv.Nodes.Add(gasitin, tvwChild, NodMeniu & "-" & NodSubmeniu & "-" & Submeniu, NumeSubMeniu)
            Next Submeniu
        End If
    Next NodSubmeniu
Next NodMeniu

'are there some sub-submenus? if so, populate the treeview with them:
'start parsing form menus all over again; tough work, indeed...
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    For NodSubmeniu = 0 To TotalItemuri - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
        For NodSubSubMeniu = 0 To TotalSubItemuri - 1
            'extracting the handle of each sub-submenu:
            ComandaNivel2 = GetSubMenu(ComandaNivel1, NodSubSubMeniu)
            'extracting the number of items this sub-submenu contains:
            TotalSubSubItemuri = GetMenuItemCount(ComandaNivel2)
            If TotalSubSubItemuri > 0 Then
                gasitin = NodMeniu & "-" & NodSubmeniu & "-" & NodSubSubMeniu '=key atribuita fiecarui nod
                'now populating the submenu nodes with the sub-subitems:
                For SubSubMeniu = 0 To TotalSubSubItemuri - 1
                    NumeObtinut4 = GetMenuString(ComandaNivel2, SubSubMeniu, NumeSubSubMeniu, Len(NumeSubSubMeniu), MF_BYPOSITION)
                    Set Nod = tv.Nodes.Add(gasitin, tvwChild, gasitin & "-" & SubSubMeniu, NumeSubSubMeniu)
                Next SubSubMeniu
            End If
        Next NodSubSubMeniu
    Next NodSubmeniu
Next NodMeniu

For Each Nod In tv.Nodes
    If Nod.Children > 0 Then Nod.Expanded = True
Next Nod
tv.Nodes(1).Selected = True
End Sub

Private Sub mpiau_Click()
memyself = MsgBox("  Menu structure retriever" & vbCrLf & "Victor Stefanescu 09.2000" & vbCrLf & "   gimelhai@mailcity.com", vbInformation, "Menu retriever")
End Sub

Private Sub newmenuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
If newmenuc.Text <> "" Then
Call ModificaMeniul
End If
End If
End Sub

Private Sub ModificaMeniul()
'procedure for modifying a menu item caption
Dim lngID, hMenu, hSubMenu, Submeniu, TotalItemuri0, TotalItemuri1 As Long
Dim NumeObtinut, NumeObtinut0, NumeObtinut1, NumeObtinut2, NumeObtinut3, NodMeniu, NodSubmeniu, NodSubSubMeniu As Long, NumeMeniu As String * 20, NumeSubMeniu As String * 20, NumeSubSubMeniu As String * 20, Nod As Node
Dim ComandaNivel0, ComandaNivel1, ComandaNivel2 As Long
Dim TotalSubSubItemuri As Long

'the first command level:
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri0 = GetMenuItemCount(ComandaNivel0)
    'search for the selected node name in the first command level:
    For Meniu = 0 To TotalItemuri0 - 1
        NumeObtinut0 = GetMenuString(ComandaNivel0, Meniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
        If Left$(NumeMeniu, NumeObtinut0) = Left$(numeniul.Caption, NumeObtinut0) Then 'they match; rename:
            hMenu = GetMenu(Form1.hwnd)
            'extracting the handle of each submenu
            ComandaNivel1 = GetSubMenu(hMenu, NodMeniu)
            'first retain the ID of the item to refer to:
            lngID = GetMenuItemID(ComandaNivel1, Meniu)
            Call ModifyMenu(hMenu, lngID, MF_STRING, lngID, newmenuc.Text)
            DaNumele 'repopulate the treeview to reflect item text change
            'interface refresh:
            newmenuc.Text = "": numeniul.Caption = "": chmenul.Caption = ""
            newmenuc.Visible = False: chmenul.Visible = True
            Exit Sub
        End If
    Next Meniu
Next NodMeniu


'the second command level:
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri0 = GetMenuItemCount(ComandaNivel0)
    'search for the selected node name in the first command level:
    For NodSubmeniu = 0 To TotalItemuri0 - 1
        'extracting the handle of each submenu
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
        If TotalSubItemuri > 0 Then
            For Submeniu = 0 To TotalSubItemuri - 1
                NumeObtinut1 = GetMenuString(ComandaNivel1, Submeniu, NumeSubMeniu, Len(NumeSubMeniu), MF_BYPOSITION)
                If Left$(NumeSubMeniu, NumeObtinut1) = Left$(numeniul.Caption, NumeObtinut1) Then 'they match; rename:
                    'first retain the ID of the item to refer to:
                    lngID = GetMenuItemID(ComandaNivel1, Submeniu)
                    Call ModifyMenu(ComandaNivel1, lngID, MF_STRING, lngID, newmenuc.Text)
                    DaNumele 'repopulate the treeview to reflect item text change
                    'interface refresh:
                    newmenuc.Text = "": numeniul.Caption = "": chmenul.Caption = ""
                    newmenuc.Visible = False: chmenul.Visible = True
                    Exit Sub
                End If
            Next Submeniu
        End If
    Next NodSubmeniu
Next NodMeniu


'the third command level:
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri0 = GetMenuItemCount(ComandaNivel0)
    'search for the selected node name in the second command level:
    For NodSubmeniu = 0 To TotalItemuri0 - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number fo items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
        For NodSubSubMeniu = 0 To TotalSubItemuri - 1
            'extracting the handle of each sub-submenu:
            ComandaNivel2 = GetSubMenu(ComandaNivel1, NodSubSubMeniu)
            'extracting the number of items this sub-submenu contains:
            TotalSubSubItemuri = GetMenuItemCount(ComandaNivel2)
            If TotalSubSubItemuri > 0 Then
                For Submeniu = 0 To TotalSubSubItemuri - 1
                    NumeObtinut2 = GetMenuString(ComandaNivel2, Submeniu, NumeSubMeniu, Len(NumeSubMeniu), MF_BYPOSITION)
                    If Left$(NumeSubMeniu, NumeObtinut2) = Left$(numeniul.Caption, NumeObtinut2) Then
                        'first retain the ID of the item to refer to:
                        lngID = GetMenuItemID(ComandaNivel2, Submeniu)
                        Call ModifyMenu(ComandaNivel2, lngID, MF_STRING, lngID, newmenuc.Text)
                        DaNumele 'repopulate the treeview to reflect item text change
                        'interface refresh:
                        newmenuc.Text = "": numeniul.Caption = "": chmenul.Caption = ""
                        newmenuc.Visible = False: chmenul.Visible = True
                        Exit Sub
                    End If
                Next Submeniu
            End If
        Next NodSubSubMeniu
    Next NodSubmeniu
Next NodMeniu
End Sub

Private Sub addmenit_Click() '"Add a menu item" label
'decided to add a command to the selected (sub)menu
If (tv.Nodes.Count = 0 Or numeniul.Caption = "") Then Exit Sub
If addmenit.Enabled = True Then
addmenit.Visible = False
addmenut.Text = "": addmenut.Visible = True: addmenut.SetFocus
End If
End Sub
Private Sub addmenut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If addmenut.Text <> "" Then
        Call AdaugaMeniu
    End If
End If
End Sub
Private Sub AdaugaMeniu()
'procedure to add a command to the selected (sub)menu
Dim NumeMeniu As String * 20, NumeSubMeniu As String * 20, NumeSubSubMeniu As String * 20
Dim UndeAdd As Long

'add to 0-level, main menus (Program, Edit, Operations):
'following the known procedure...
For NodMeniu = 0 To MeniuriPrincipale - 1
    SubItemMeniu = GetSubMenu(MeniulForm, NodMeniu)
    TotalItemuri = GetMenuItemCount(SubItemMeniu)
    For Meniu = 0 To TotalItemuri - 1
        NumeObtinut = GetMenuString(MeniulForm, NodMeniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
        If Left$(NumeMeniu, NumeObtinut) = Left$(numeniul.Caption, NumeObtinut) Then 'they match; add to this
            'first identify the location to refer to:
            UndeAdd = GetSubMenu(SubItemMeniu, Meniu)
            'add now:
            Call AppendMenu(SubItemMeniu, MF_STRING, ByVal 0&, addmenut.Text)
            'refresh the treeview:
            DaNumele
            Exit Sub
        End If
    Next Meniu
Next NodMeniu

'add to 1st level of submenus (Informations etc.)
'following the known procedure...
For NodMeniu = 0 To MeniuriPrincipale - 1
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    'now populating the main menu treeview nodes with the items:
    For Meniu = 0 To TotalItemuri - 1
        NumeObtinut = GetMenuString(ComandaNivel0, Meniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
        If Left$(NumeMeniu, NumeObtinut) = Left$(numeniul.Caption, NumeObtinut) Then 'they match; add to this
            'first identify the location to refer to:
            UndeAdd = GetSubMenu(ComandaNivel0, Meniu)
            Call AppendMenu(UndeAdd, MF_STRING, ByVal 0&, addmenut.Text)
            'MsgBox "Adding to" & NumeMeniu 'just a control message
            DaNumele
            Exit Sub
        End If
    Next Meniu
Next NodMeniu

'add to 2nd level of submenus (Company etc.)
'following the known procedure...
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the number of items this main menu contains:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    For NodSubmeniu = 0 To TotalItemuri - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
        If TotalSubItemuri > 0 Then
            For Submeniu = 0 To TotalSubItemuri - 1
                NumeObtinut2 = GetMenuString(ComandaNivel1, Submeniu, NumeSubMeniu, Len(NumeSubMeniu), MF_BYPOSITION)
                If Left$(NumeSubMeniu, NumeObtinut2) = Left$(numeniul.Caption, NumeObtinut2) Then 'they match; add to this
                    'first identify the location to refer to:
                    UndeAdd = GetSubMenu(ComandaNivel1, Submeniu)
                    'add now:
                    Call AppendMenu(UndeAdd, MF_STRING, ByVal 0&, addmenut.Text)
                    'refresh the treeview:
                    DaNumele
                    Exit Sub
                End If
            Next Submeniu
        End If
    Next NodSubmeniu
Next NodMeniu

DaNumele
End Sub

Private Sub remmenit_Click() '"Remove this item" label
'procedure to remove (not to destroy! - separate APIf: DestroyMenu) a submenu or command
Dim NumeMeniu As String * 20, NumeSubMeniu As String * 20, NumeSubSubMeniu As String * 20

If tv.Nodes.Count = 0 Then Exit Sub 'menu structure not read yet

'search in the 1st submenu level (Program/Informations):
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu) 'the menu to be !MODIFIED!
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    'now populating the treeview nodes with the items:
    For Meniu = 0 To TotalItemuri - 1
        NumeObtinut = GetMenuString(ComandaNivel0, Meniu, NumeMeniu, Len(NumeMeniu), MF_BYPOSITION)
        If Left$(NumeMeniu, NumeObtinut) = Left$(numeniul.Caption, NumeObtinut) Then 'they match; remove !FROM! this
            'removing menu (item): the first parameter is the menu to be !MODIFIED!
            Call RemoveMenu(ComandaNivel0, Meniu, MF_BYPOSITION) 'or (MF_BYPOSITION Or MF_REMOVE)
            'refresh the treeview:
            DaNumele
            Exit Sub
        End If
    Next Meniu
Next NodMeniu

'search in the 2nd submenu level (Program/Informations/Company):
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    For NodSubmeniu = 0 To TotalItemuri - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
            For Submeniu = 0 To TotalSubItemuri - 1
                NumeObtinut2 = GetMenuString(ComandaNivel1, Submeniu, NumeSubMeniu, Len(NumeSubMeniu), MF_BYPOSITION)
                If Left$(NumeSubMeniu, NumeObtinut2) = Left$(numeniul.Caption, NumeObtinut2) Then 'they match; remove !FROM! this
                    'removing menu (item): the first parameter is the menu to be !MODIFIED!
                    Call RemoveMenu(ComandaNivel1, Submeniu, MF_BYPOSITION) 'or (MF_BYPOSITION Or MF_REMOVE)
                    'refresh the treeview:
                    DaNumele
                    Exit Sub
                End If
            Next Submeniu
    Next NodSubmeniu
Next NodMeniu

'search in the 3rd submenu level (only commands in this example menu structure)
For NodMeniu = 0 To MeniuriPrincipale - 1
    'extracting the handle of each main menu:
    ComandaNivel0 = GetSubMenu(MeniulForm, NodMeniu)
    'extracting the number of items this main menu contains:
    TotalItemuri = GetMenuItemCount(ComandaNivel0)
    For NodSubmeniu = 0 To TotalItemuri - 1
        'extracting the handle of each submenu:
        ComandaNivel1 = GetSubMenu(ComandaNivel0, NodSubmeniu)
        'extracting the number of items this submenu contains:
        TotalSubItemuri = GetMenuItemCount(ComandaNivel1)
            For Submeniu = 0 To TotalSubItemuri - 1
                'extracting the handle of each sub-submenu (command):
                ComandaNivel2 = GetSubMenu(ComandaNivel1, Submeniu)
                'extracting the number of items this sub-submenu contains (no items, as this is a last-level command)
                TotalSubSubItemuri = GetMenuItemCount(ComandaNivel2)
                For SubSubMeniu = 0 To TotalSubSubItemuri - 1
                    NumeObtinut3 = GetMenuString(ComandaNivel2, SubSubMeniu, NumeSubSubMeniu, Len(NumeSubSubMeniu), MF_BYPOSITION)
                    If Left$(NumeSubSubMeniu, NumeObtinut3) = Left$(numeniul.Caption, NumeObtinut3) Then 'they match; remove !FROM! this
                        'removing this command: the first parameter is the menu to be !MODIFIED!
                        Call RemoveMenu(ComandaNivel2, SubSubMeniu, MF_BYPOSITION) 'or (MF_BYPOSITION Or MF_REMOVE)
                        'refresh the treeview:
                        DaNumele
                        Exit Sub
                    End If
                Next SubSubMeniu
            Next Submeniu
    Next NodSubmeniu
Next NodMeniu
End Sub

Private Sub chmenul_Click() 'the label with the text of the selected tv node (menu item)
If (tv.Nodes.Count = 0 Or numeniul.Caption = "") Then Exit Sub
If chmenul.Enabled = True Then
chmenul.Visible = False: newmenuc.Visible = True
newmenuc.SetFocus: newmenuc.SelStart = Len(newmenuc.Text)
End If
End Sub

Private Sub tv_Click() 'click a treeview node

On Error GoTo nutvitems 'in case the tv is empty: no Parent nodes, no FullPath, no Children etc; we have no node to refer to:
addmenut.Visible = False: addmenit.Visible = True
numeniul.Caption = tv.SelectedItem.Text
newmenuc.Text = numeniul.Caption
chmenul.Caption = numeniul.Caption
newmenuc.Visible = False: chmenul.Visible = True
addmenut.Visible = False: addmenit.Visible = True
If (tv.SelectedItem.Children > 0) Then
    'cannot modify main menu (or submenu) caption, but can add to non-empty, parent nodes:
    If tv.SelectedItem.Text = tv.SelectedItem.FullPath Then '0-level, main menu
        'cannot remove 0-level, main menus:
        remmenit.Enabled = False
    Else
        'submenus/commands of the 1st level and beyond - of contrary:
        remmenit.Enabled = True
    End If
    chmenul.Enabled = False
    addmenit.Enabled = True
    Else
    'can modify menu item caption, but cannot add to empty nodes (commands):
    chmenul.Enabled = True
    addmenit.Enabled = False
    remmenit.Enabled = True
End If
tipmenl.Caption = "Menu item level: " & tv.SelectedItem.Key
tipmenl.Caption = tipmenl.Caption & vbCrLf & tv.SelectedItem.Children & " submenus/commands"
Exit Sub
nutvitems: 'empty treeview, as stated above
Exit Sub
End Sub
Private Sub tv_KeyPress(KeyAscii As Integer) 'press a key while focus set to tv
addmenut.Visible = False: addmenit.Visible = True
If KeyAscii = 13 Then KeyAscii = 0
'get rid off that annoying beep, similarly to pressing Enter in a TextBox
End Sub
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
'if 'a' pressed, selected node becomes 'Abandon' or other
addmenut.Visible = False: addmenit.Visible = True
Call tv_Click
End Sub
Private Sub schmimag_Click()
'click the image surrounding the dameniub ("Populate") label
addmenut.Visible = False: addmenit.Visible = True
dameniub_Click
End Sub


'// interface adjusting procedures

Private Sub dameniub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
schmimag.BorderStyle = 1: dameniub.ForeColor = &HFFFF&
End Sub
Private Sub schmimag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
schmimag.BorderStyle = 1
dameniub.ForeColor = &HFFFF&
End Sub
Private Sub chmenul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
chmenul.BorderStyle = 1
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
chmenul.BorderStyle = 0
End Sub
Private Sub tipmenl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
chmenul.BorderStyle = 0
Shape2.Visible = False: Shape3.Visible = False
End Sub
Private Sub men_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
schmimag.BorderStyle = 0: dameniub.ForeColor = &HFFFF00
End Sub
Private Sub injos_Click()
'click to drop the informations label
If infol.Height < 2400 Then
infol.Height = 2400: Shape1.Height = 2424
Else
infol.Height = injos.Height + 20: Shape1.Height = injos.Height + 30
End If
End Sub
Private Sub injos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infol.BorderStyle = 1
Shape2.Visible = False: Shape3.Visible = False
If Cinta = False Then
    Cinta = True
    PlaySound App.Path & "\infosnd.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub
Private Sub infol_Click()
'click in the informations label
PopupMenu mpinf
End Sub
Private Sub infol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
infol.BorderStyle = 1
Shape2.Visible = False: Shape3.Visible = False
If Cinta = False Then
    Cinta = True
    PlaySound App.Path & "\infosnd.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub
Private Sub tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
chmenul.BorderStyle = 0: infol.BorderStyle = 0
schmimag.BorderStyle = 0
Shape2.Visible = False: Shape3.Visible = False
End Sub
Private Sub addmenit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
infol.BorderStyle = 0
Shape3.Visible = True: Shape2.Visible = False
If Cinta = False Then
    Cinta = True
    PlaySound App.Path & "\selctsnd.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub
Private Sub remmenit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
infol.BorderStyle = 0
Shape3.Visible = False: Shape2.Visible = True
If Cinta = False Then
    Cinta = True
    PlaySound App.Path & "\selctsnd.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addmenut.Visible = False: addmenit.Visible = True
Label2.ForeColor = &HFFFF00
If Cinta = False Then
    Cinta = True
    PlaySound App.Path & "\emerg.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.Visible = False: Shape3.Visible = False
addmenut.Visible = False: addmenit.Visible = True
Cinta = False
schmimag.BorderStyle = 0: chmenul.BorderStyle = 0: infol.BorderStyle = 0
addmenit.BorderStyle = 0: remmenit.BorderStyle = 0
Label2.ForeColor = &H800000
If dameniub.ForeColor = &HFFFF00 Then
Exit Sub
Else
dameniub.ForeColor = &HFFFF00
End If
End Sub

Private Sub mpaba_Click()
'Program/Abandon
Unload Form1: Set Form1 = Nothing
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End
End Sub
