*************************
Office XP-style Toolbar *
Telperion               *
telperionn@yahoo.com    *
*************************
You are free to use and modify this control as you see fit, but I would love to hear about any changes you make to it!

This control is nowhere near perfect and is just to tide people over until Microsoft releases its version.  I'm sure there are many bugs that I'm not aware of; feel free to fix them if you want.  Combine this control with the vbAccelerator Icon Menu Control (www.vbaccelerator.com) and you can make your apps look almost completely like Office XP!

Just compile the OCX and place it in your Windows\System folder and make sure you check it in the Components page of your project.



*********
Usage:  *
*********
Using the OfficeXPToolbar is very easy.
--------------------------------------------------------------------------
To add a button:
OfficeXPToolbar1.AddButton(ButtonType As Integer, Icon, Optional ButtonKey As String, Optional ToolTipText As String, Optional LineColor As OLE_COLOR, Optional FillColor As OLE_COLOR, Optional PressColor As OLE_COLOR)

ButtonType:
0 = Separator
1 = Button

Icon:
Source of Icon for the button.  Can be imagelist index, loaded image, .picture property, etc.

ButtonKey:
Unique key to identify button.  If none is supplied, ButtonKey = Index.

ToolTipText:
Text that will pop up when mouse pointer hovers over toolbar button.

LineColor:
A value specifiying the color of the border of the highlight that will appear for that particular button.  If none is supplied, OfficeXPToolbar1.HighlightBorderColor is used.

FillColor:
A value specifiying the color of the fill of the highlight that will appear for that particular button.  If none is supplied, OfficeXPToolbar1.HighlightBackColor is used.

PressColor:
A value specifiying the color of the fill of the highlight that will appear for that particular button when it is pressed.  If none is supplied, OfficeXPToolbar1.HighlightPressColor is used.
--------------------------------------------------------------------------
To change the status of a button:
OfficeXPToolbar1.SetStatus(Index as Integer, Status as Integer)

Index:
The index of the button to change the status for.  You must keep track of the index when you create the button.  See Notes for more detail.

Status:
0 = Disabled
1 = Enabled
--------------------------------------------------------------------------
To get the status of a button:
TempVar as Integer = OfficeXPToolbar1.GetStatus(Index as Integer)

Index:
The index of the button to fetch the status for.

TempVar:
Variable to return the status.
0 = Disabled
1 = Enabled




*****************************
Limitations of the Control: *
*****************************
-I only bothered to make the control work properly with images that are 16x16 in size.  Although it may work with other sizes, don't count on it.
-You cannot assign indexes to the toolbar buttons; the index increments automatically, and you must keep track of it yourself.  You can, however, assign keys to make identification easier.
-You cannot remove buttons once you have added them.  I just didn't think that it would be worth the effort to include that.
-You cannot create buttons at design time; you MUST create them using code at run time.  It's not as bad as it sounds.
-You must define toolbar buttons in the order in which they will appear on the toolbar.
-You can only set the status of a toolbar button, not a separator.  If you try to use SetStatus on an index that corresponds to a separator, you will return an "index does not exist" error.
-There is no docking or banding supported.



********
Notes: *
********
-Separators have their own index, just like normal buttons.
-When you create buttons, an index is assigned to each one sequentially, starting at 0.  This includes separators.



**********
Example: *
**********
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(1).Picture, "new", "New" 'Index = 0
OfficeXPToolbar1.AddButton 0, ImageList1.ListImages(1).Picture 'Separator: Index = 1, image source is supplied but will not be used.
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(2).Picture, "open", "Open" 'Index = 2
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(3).Picture, "save", "Save" 'Index = 3
OfficeXPToolbar1.AddButton 0, ImageList1.ListImages(1).Picture 'Separator: Index = 4
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(4).Picture, "cut", "Cut" 'Index = 5
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(5).Picture, "copy", "Copy" 'Index = 6
OfficeXPToolbar1.AddButton 1, ImageList1.ListImages(6).Picture, "paste", "Paste" 'Index = 7

OfficeXPToolbar1.SetStatus 5, 0 'Disable the Cut button
OfficeXPToolbar1.SetStatus 6, 0 'Disable the Copy button

OfficeXPToolbar1.SetStatus 5, 1 'Enable the Cut button
OfficeXPToolbar1.SetStatus 6, 1 'Enable the Copy button

Private Sub OfficeXPToolbar1_Click(Index As Integer, ButtonKey As String)
Select Case ButtonKey

Case "cut"
CutText

Case "copy"
CopyText

Case "paste"
PasteText

Case "new"
NewFile

Case "open"
OpenFile

Case "save"
SaveFile

End Select
End Sub