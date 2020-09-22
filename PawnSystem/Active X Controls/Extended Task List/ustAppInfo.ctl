VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl usrAppInfo 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2376
      Left            =   48
      TabIndex        =   0
      Top             =   252
      Width           =   3612
      _ExtentX        =   6371
      _ExtentY        =   4191
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "usrAppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=lvFiles,lvFiles,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Default Property Values:
Const m_def_TaskName = "*"
'Property Variables:
Dim m_TaskName As String




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvFiles,lvFiles,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lvFiles.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lvFiles.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvFiles,lvFiles,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ForeColor = lvFiles.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lvFiles.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvFiles,lvFiles,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lvFiles.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lvFiles.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvFiles,lvFiles,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lvFiles.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lvFiles.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lvFiles,lvFiles,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = lvFiles.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    lvFiles.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    UserControl.Refresh
End Sub

Private Sub lvFiles_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Paint()
    PopulateTaskInfo lvFiles, m_TaskName
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lvFiles.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lvFiles.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    lvFiles.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lvFiles.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lvFiles.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_TaskName = PropBag.ReadProperty("TaskName", m_def_TaskName)
End Sub

Private Sub UserControl_Resize()
    lvFiles.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lvFiles.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lvFiles.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", lvFiles.Enabled, True)
    Call PropBag.WriteProperty("Font", lvFiles.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", lvFiles.BorderStyle, 1)
    Call PropBag.WriteProperty("TaskName", m_TaskName, m_def_TaskName)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get TaskName() As String
Attribute TaskName.VB_Description = "Sets the name of the task to retrieve information for, use * for all currently running tasks."
    TaskName = m_TaskName
End Property

Public Property Let TaskName(ByVal New_TaskName As String)
    m_TaskName = New_TaskName
    PropertyChanged "TaskName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TaskName = m_def_TaskName
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SaveToFile(rsFileName As String) As Variant
Attribute SaveToFile.VB_Description = "Saves the infromation to a file."
Dim llFile  As Long
Dim loitem  As ListItem
Dim loCol   As ColumnHeader
Dim llLoop  As Long
Dim llEnd   As Long
Dim llCount As Long
Dim lsItem  As String

    llFile = FreeFile
    
    
    Open rsFileName For Append As llFile
    Print #llFile, "Modules for " & m_TaskName
    Print #llFile, vbCrLf
    
    For Each loCol In lvFiles.ColumnHeaders
        Print #llFile, loCol.Text & vbTab;
    Next
    Print #llFile, vbCrLf
    llEnd = lvFiles.ColumnHeaders.Count
        
    For Each loitem In lvFiles.ListItems
        llCount = 1
        For llLoop = 1 To llEnd
            If llCount = 1 Then
                lsItem = loitem.Text
            Else
                lsItem = loitem.SubItems(llCount - 1)
            End If
            Print #llFile, lsItem & vbTab;
            llCount = llCount + 1
        Next
        Print #llFile, vbCrLf
        
    Next
    
    Print #llFile, vbCrLf
    Print #llFile, vbCrLf
    Close #llFile
    
End Function

