VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMain"
   ClientHeight    =   5520
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvMain 
      Height          =   4530
      Left            =   135
      TabIndex        =   2
      Top             =   405
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip tsMain 
      Height          =   5010
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   8837
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Editing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Navigation"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Window"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Window Menu"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Project Explorer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tool Box"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Watch Window"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   9870
      TabIndex        =   0
      Top             =   5085
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const LVM_FIRST = &H1000
Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE = -1

Private strCode_Gen_keys() As String
Private strCode_Edit_keys() As String
Private strCode_Nav_keys() As String
Private strCode_Win_keys() As String
Private strCode_Menu_keys() As String
Private strCode_Proj_keys() As String
Private strCode_Tool_keys() As String
Private strCode_Watch_keys() As String

Option Explicit


Private Sub Form_Load()
    LoadArrays 101, 119, strCode_Gen_keys
    LoadArrays 122, 139, strCode_Edit_keys
    LoadArrays 140, 159, strCode_Nav_keys
    LoadArrays 160, 189, strCode_Win_keys
    LoadArrays 190, 229, strCode_Menu_keys
    
    LoadArrays 190, 229, strCode_Proj_keys
    LoadArrays 190, 229, strCode_Tool_keys
    LoadArrays 190, 229, strCode_Watch_keys
    
    LoadListviewFromArray strCode_Gen_keys
End Sub
Private Sub OKButton_Click()
    Connect.Hide
End Sub
Private Function LoadArrays(lngStart As Long, lngEnd As Long, ByRef varData As Variant) As Boolean
    Dim intIDX_Index As Integer
    Dim intIndex As Integer
    Dim strTempLine As String
    Dim strTemp() As String
    
    intIndex = 0
    ReDim varData(lngEnd, 1)
    
    For intIDX_Index = lngStart To lngEnd
        strTempLine = LoadResString(intIDX_Index)
        If strTempLine <> "." Then
            strTemp = Split(strTempLine, ":", , vbTextCompare)
            
            varData(intIndex, 0) = strTemp(0)
            varData(intIndex, 1) = strTemp(1)
            intIndex = intIndex + 1
        End If
    Next intIDX_Index
    LoadArrays = True
End Function
Private Function LoadListviewFromArray(varData As Variant) As Boolean
    Dim lngUpIndex As Long
    Dim lngLoIndex As Long
    Dim lngIndex As Long
    Dim oListItem As ListItem
    Dim strKeys As String
    Dim strDesc As String
    Dim Index As Long
    lngUpIndex = UBound(varData)
    lngLoIndex = 0
    
    With lvMain
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "Key Press"
        .ColumnHeaders.Add , , "To Do This"
        For lngIndex = lngLoIndex To lngUpIndex
            strKeys = varData(lngIndex, 0)
            strDesc = varData(lngIndex, 1)
            If strKeys <> "." Then
                Set oListItem = .ListItems.Add(, , strKeys)
                oListItem.SubItems(1) = strDesc
            End If
        Next lngIndex
    End With
    
    LockWindowUpdate lvMain.hWnd
    SendMessage lvMain.hWnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE
    SendMessage lvMain.hWnd, LVM_SETCOLUMNWIDTH, 1, LVSCW_AUTOSIZE
    LockWindowUpdate 0
    
End Function

Private Sub tsMain_Click()
    Dim strTabCaption As String
    
    strTabCaption = tsMain.SelectedItem.Caption
    Select Case strTabCaption
    Case "Code General"
        LoadListviewFromArray strCode_Gen_keys
    Case "Code Editing"
        LoadListviewFromArray strCode_Edit_keys
    Case "Code Navigation"
        LoadListviewFromArray strCode_Nav_keys
    Case "Code Window"
        LoadListviewFromArray strCode_Win_keys
    Case "Code Window Menu"
        LoadListviewFromArray strCode_Menu_keys
    Case "Project Explorer"
        LoadListviewFromArray strCode_Proj_keys
    Case "Tool Box"
        LoadListviewFromArray strCode_Tool_keys
    Case "Watch Window"
        LoadListviewFromArray strCode_Watch_keys
        
    End Select
End Sub
