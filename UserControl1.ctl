VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ScaleHeight     =   1725
   ScaleWidth      =   4620
   Begin VB.PictureBox IconList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   0
      Left            =   120
      ScaleHeight     =   1158.478
      ScaleMode       =   0  'User
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.HScrollBar iconscroll 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   5
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblIconListTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_ListIndex = 0
Const m_def_ListCount = 0
Const m_def_MultiSelect = 0
Const m_def_OLEDragMode = 0
Const m_def_Value = 0

'Property Variables:
Dim m_MultiSelect As Integer
Dim m_OLEDragMode As Integer
Dim IconStart As Integer
Dim iStartSpace As Integer
Dim bClearUsed As Boolean

'Event Declarations:
Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Event Click(Index As Integer)  'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick(Index As Integer) 'MappingInfo=UserControl,UserControl,-1,DblClick
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."

Private Sub IconList_Click(Index As Integer)
For i = 0 To IconList.Count - 1
    IconList(i).BorderStyle = 0
Next i
IconList(Index).BorderStyle = 1
RaiseEvent Click(Index)
End Sub

Private Sub IconList_dblClick(Index As Integer)
    RaiseEvent DblClick(Index)
End Sub

Private Sub IconScroll_Change()
IconStart = iconscroll.Value
Update_Icons
End Sub

Private Sub IconScroll_Scroll()
IconStart = iconscroll.Value
Update_Icons
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Update_Icons
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
'On Error GoTo ErrorCap
List1.Clear
If IconList.Count > 2 Then
    Do
        Unload IconList(IconList.Count - 1)
        Unload lblIconListTitle(lblIconListTitle.Count - 1)
    Loop Until IconList.Count = 1
End If
IconList(0).Picture = LoadPicture()
lblIconListTitle(0).Caption = ""
IconList(0).Visible = False
lblIconListTitle(0).Visible = False
bClearUsed = True
DoEvents
Exit Sub
ErrorCap:
If Err.Number = 340 Then
    Resume Next
Else
    MsgBox Err.Number & " " & Err.Description
End If
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = m_OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
    m_OLEDragMode = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_Resize()
iconscroll.Width = UserControl.Width
iconscroll.Top = UserControl.Height - iconscroll.Height
UserControl.Height = 1785
Update_Icons
ViewableObjects
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_OLEDragMode = m_def_OLEDragMode
    m_ListIndex = m_def_ListIndex
    m_ListCount = m_def_ListCount
    m_MultiSelect = m_def_MultiSelect
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_OLEDragMode = PropBag.ReadProperty("OLEDragMode", m_def_OLEDragMode)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
    iconscroll.SmallChange = PropBag.ReadProperty("SmallChange", 1)
    iconscroll.LargeChange = PropBag.ReadProperty("LargeChange", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("OLEDragMode", m_OLEDragMode, m_def_OLEDragMode)
    Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
    Call PropBag.WriteProperty("SmallChange", iconscroll.SmallChange, 1)
    Call PropBag.WriteProperty("LargeChange", iconscroll.LargeChange, 1)
End Sub

Private Sub Update_Icons(Optional bVisible As Boolean = True)
Dim i As Integer
For i = 0 To IconList.Count - 1
    IconList(i).Visible = False
    lblIconListTitle(i).Visible = False
Next i

'On Error GoTo ErrorSpot
Dim bFirst As Boolean, iViewObjects As Integer
bFirst = True
For i = IconStart To List1.ListCount - 1
    If i > 0 Then
        If IconStart > 0 And bFirst = True Then
            IconList(i).Move (IconList(0).Left), 0, 1215, 1095
            lblIconListTitle(i).Move (lblIconListTitle(0).Left), 1095
            bFirst = False
        Else
            IconList(i).Move (IconList(i - 1).Left + IconList(i - 1).Width + 120), 0, 1215, 1095
            lblIconListTitle(i).Move (lblIconListTitle(i - 1).Left + lblIconListTitle(i - 1).Width + 120), 1095
        End If
    End If

    If (IconList(i).Left + IconList(i).Width) > UserControl.Width Then
        Exit For
    Else
        If bVisible Then
            IconList(i).Visible = True
            lblIconListTitle(i).Visible = True
        End If
    End If
Next i

bClearUsed = False
Exit Sub
ErrorSpot:
If Err.Number = 360 Then
    Resume Next
Else
    MsgBox Err.Number & " " & Err.Description
End If
End Sub

Private Function get_SubPath(sPath As String) As String
get_SubPath = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function

Private Sub ViewableObjects(Optional bAddItem As Boolean = False)
Dim iObjects As Integer, iLastVis As Integer
For i = 0 To IconList.Count - 1
    If ((IconList(i).Left + IconList(i).Width) <= UserControl.Width) And (IconList(i).Left > 0) Then
        iObjects = iObjects + 1
        iLastVis = i
    End If
Next i
iconscroll.Max = List1.ListCount - iObjects
IconList(0).Left = (UserControl.Width - (IconList(iLastVis).Width + IconList(iLastVis).Left)) / 2 + 120
lblIconListTitle(0).Left = IconList(0).Left

Update_Icons Not (bAddItem)
End Sub

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant, Optional sText As String)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
If Item = "" Then Exit Sub
List1.AddItem Item, Index
iconscroll.Max = List1.ListCount

i = List1.ListCount - 1
On Error Resume Next
If (i > 0) Then
    Load IconList(i)
    IconList(i).Move (IconList(i - 1).Left + IconList(i - 1).Width + 120), 0, 1215, 1095
    Load lblIconListTitle(i)
    lblIconListTitle(i).Move (lblIconListTitle(i - 1).Left + lblIconListTitle(i - 1).Width + 120), 1095
End If
On Error GoTo 0

Dim picTemp As Control, picImage As Control
Set picTemp = UserControl.Controls.Add("vb.picturebox", "picTemp")
Set picImage = UserControl.Controls.Add("vb.picturebox", "picImage")
picTemp.AutoSize = True
picImage.AutoSize = True
picImage.AutoRedraw = True
picTemp.Picture = LoadPicture(List1.List(i))
picImage.Picture = LoadPicture()

If picTemp.Width > 1215 Or picTemp.Height > 1095 Then
    Dim cRatio As Currency
    picImage.Width = 1215
    picImage.Height = 1095
    If picTemp.Width > picTemp.Height Then
        cRatio = picTemp.Width / picImage.Width
    Else
        cRatio = picTemp.Height / picImage.Height
    End If
    picImage.Width = picTemp.Width / cRatio
    picImage.Height = picTemp.Height / cRatio
    picImage.PaintPicture picTemp.Picture, 0, 0, picImage.Width, picImage.Height
    picImage.Picture = picImage.Image
Else
    picImage.Width = picTemp.Width
    picImage.Height = picTemp.Height
    picImage.Picture = picTemp.Picture
End If
picImage.Refresh
picTemp.Picture = LoadPicture
UserControl.Controls.Remove "picTemp"

IconList(i).Picture = picImage.Picture

picImage.Picture = LoadPicture
UserControl.Controls.Remove "picImage"

IconList(i).ToolTipText = List1.List(i + IconStart)
lblIconListTitle(i).Caption = get_SubPath(List1.List(i))

If (IconList(i).Left + IconList(i).Width) > UserControl.Width Then
    iconscroll.Enabled = True
Else
    iconscroll.Enabled = False
End If
ViewableObjects True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
List1.RemoveItem Index
iconscroll.Max = List1.ListCount
If (IconList(i - 1).Left + IconList(i - 1).Width) > UserControl.Width Then
    iconscroll.Enabled = True
Else
    iconscroll.Enabled = False
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MultiSelect() As Integer
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether a user can make multiple selections in a control."
    MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Integer)
    m_MultiSelect = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iconscroll,iconscroll,-1,SmallChange
Public Property Get SmallChange() As Integer
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."
    SmallChange = iconscroll.SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Integer)
    iconscroll.SmallChange() = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iconscroll,iconscroll,-1,LargeChange
Public Property Get LargeChange() As Integer
Attribute LargeChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area."
    LargeChange = iconscroll.LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    iconscroll.LargeChange() = New_LargeChange
    PropertyChanged "LargeChange"
End Property

