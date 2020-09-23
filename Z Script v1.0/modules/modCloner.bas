Attribute VB_Name = "modCloner"
Public frmCloned(250) As New frmClone
Dim vControlCount As Integer
Public vFormCount As Integer
Public vButtonCount As Integer


Public Function CopyControl(Control As Variant, Visible As Boolean, Top As Integer, Left As Integer, Width As Integer, Height As Integer) 'Just the perimeters for CopyControl() - You'll know what this is though, of course, having written the AOL module

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = Visible
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
End With
End Function

Public Function CopyForm(vForm As Variant, vFormName As String)

Load frmCloned(vFormCount)
With frmCloned(vFormCount)
    .Tag = vFormName
    .Caption = vFormName
End With
frmMain.List1.AddItem ("frmCloned(" & vFormCount & ")")
vFormCount = vFormCount + 1

End Function

Public Function CopyButton(Control As Variant, Caption As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vButtonCount = Control.Count + 1
Load Control(vButtonCount)
With Control(vButtonCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyLabel(Control As Variant, Caption As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyText(Control As Variant, Text As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Text = Text
    .Tag = Name
End With

End Function

Public Function CopyCheck(Control As Variant, Caption As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyRadio(Control As Variant, Caption As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyFrame(Control As Variant, Caption As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyLink(Control As Variant, Link As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Caption = Caption
    .Tag = Name
End With

End Function

Public Function CopyImage(Control As Variant, Image As String, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Picture = Image
    .Tag = Name
End With

End Function

Public Function CopyList(Control As Variant, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Tag = Name
End With

End Function


Public Function CopyCombo(Control As Variant, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Name As String)

vControlCount = Control.Count + 1
Load Control(vControlCount)
With Control(vControlCount)
    .Visible = True
    .Top = Top
    .Left = Left
    .Width = Width
    .Height = Height
    .Tag = Name
End With

End Function
