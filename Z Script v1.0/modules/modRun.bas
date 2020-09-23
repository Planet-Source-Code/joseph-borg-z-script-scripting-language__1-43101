Attribute VB_Name = "modRun"
Dim vRunSub As Boolean
Public vEvent As Boolean

Dim vFormName As String

Dim vTitle As String

Dim vX As Integer
Dim vY As Integer
Dim vH As Integer
Dim vW As Integer

Dim vButCaption As String
Dim vButY As Integer
Dim vButX As Integer
Dim vButW As Integer
Dim vButH As Integer
Dim vButStyle As String

Dim vLblCaption As String
Dim vLblY As Integer
Dim vLblX As Integer
Dim vLblW As Integer
Dim vLblH As Integer

Dim vTxtText As String
Dim vTxtY As Integer
Dim vTxtX As Integer
Dim vTxtW As Integer
Dim vTxtH As Integer

Dim vChkCaption As String
Dim vChkY As Integer
Dim vChkX As Integer
Dim vChkW As Integer
Dim vChkH As Integer

Dim vRadCaption As String
Dim vRadY As Integer
Dim vRadX As Integer
Dim vRadW As Integer
Dim vRadH As Integer

Dim vFraCaption As String
Dim vFraY As Integer
Dim vFraX As Integer
Dim vFraW As Integer
Dim vFraH As Integer

Dim vLnkCaption As String
Dim vLnkY As Integer
Dim vLnkX As Integer
Dim vLnkW As Integer
Dim vLnkH As Integer

Dim vImgFile As String
Dim vImgY As Integer
Dim vImgX As Integer
Dim vImgW As Integer
Dim vImgH As Integer

Dim vLstY As Integer
Dim vLstX As Integer
Dim vLstW As Integer
Dim vLstH As Integer

Dim vCmbY As Integer
Dim vCmbX As Integer
Dim vCmbW As Integer
Dim vCmbH As Integer


Private mclsStyle As clsStyle

Dim vToolWin As Boolean
Dim vContBox As Boolean
Dim vMinBut As Boolean
Dim vMaxBut As Boolean
Dim vSizable As Boolean
Dim vTopMost As Boolean


Public Function RunCode(vCode As String)
On Error Resume Next

vStep = 0
vRunSub = False

j = Split(vCode, vbCrLf)

For k = 0 To UBound(j)
    vStep = vStep + 1
    
    SpaceSplit (j(k))
    vCommand = UCase(vSpc1)
    
    
    ' } (end sub)
    If vCommand = "}" And vRunSub = True Then
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                MsgBox frmCloned(i).Tag & " " & vFormName
                frmCloned(i).Visible = True
            End If
        Next i
        vRunSub = False
    End If
    
    
    ' form <name> { (load form, start sub)
    If vCommand = "FORM" And vRunSub = False Then
        SpaceSplit (j(k))
        vFormName = vSpc2
        CopyForm frmClone, vFormName
        
        vRunSub = True
        
        
        
    Else:
        
        If vCommand = "FORM" And vSpc2 = "" Then MsgBox "Form name can be null, error in line : " & k
        If vCommand = "FORM" And vSpc3 <> "{" Then MsgBox "{ missing in line : " & k
            
    End If
    
    ' title = <title>
    If vCommand = "TITLE" And vRunSub = True Then
        EqualsSplit (j(k))
        vTitle = vEqu2
        
        For i = 0 To vFormCount
             If frmCloned(i).Tag = vFormName Then frmCloned(i).Caption = vTitle
        Next i
        
        
    End If


' style = <ToolWindow>, <ControlBox>, <MinBut>, <MaxBut>, <Sizable>, <OnTop>
    If vCommand = "STYLE" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
              
        vToolWin = vCom1
        vContBox = vCom2
        vMinBut = vCom3
        vMaxBut = vCom4
        vSizable = vCom5
        vTopMost = vCom6
              
        For i = 0 To vFormCount
             If frmCloned(i).Tag = vFormName Then
                Set mclsStyle = New clsStyle
                Set mclsStyle.Client = frmCloned(i)
                mclsStyle.ToolWindow = vToolWin
                mclsStyle.ControlBox = vContBox
                mclsStyle.MinButton = vMinBut
                mclsStyle.MaxButton = vMaxBut
                mclsStyle.Sizable = vSizable
                mclsStyle.TopMost = vTopMost
             End If
        Next i
        
        
    End If

    ' size = x y h w
    If vCommand = "SIZE" And vRunSub = True Then
        EqualsSplit (j(k))
        SpaceSplit (vEqu2)
        
        vX = vSpc2
        vY = vSpc3
        vH = vSpc4
        vW = vSpc5
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
        
                frmCloned(i).Top = vX
                frmCloned(i).Left = vY
                frmCloned(i).Height = vH
                frmCloned(i).Width = vW
            End If
        Next i
        
       
    End If

    If vCommand = "ABOUT" And vRunSub = True Or vCommand = "ABOUT" And vEvent = True Then
        MsgBox "Z Script By Joseph Borg", vbInformation, "About"
    End If
    
    ' button = <caption>, <name>, y x w h, style(standard/flat/thick)
    If vCommand = "BUTTON" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vButCaption = vCom1
        
        SpaceSplit (vCom3)
        vButY = vSpc1
        vButX = vSpc2
        vButW = vSpc3
        vButH = vSpc4
        
        BracketSplit LCase(vCom4)
        
        CommaSplit (vBrack)
        
        vButStyle = vBrack
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyButton frmCloned(i).Command1, vButCaption, vButY, vButX, vButW, vButH, vCom2
                
                GetInitialBTStyle frmCloned(i).Command1(vButtonCount)
        
                If vButStyle = "standard" Then
                    SetInitialBTStyle frmCloned(i).Command1(vButtonCount)
                    mHover = False
                End If
                
                If vButStyle = "thick" Then
                    BTThick frmCloned(i).Command1(vButtonCount)
                    mHover = False

                End If
                
                If vButStyle = "flat" Then
                    
                    BTFlat frmCloned(i).Command1(vButtonCount)
                    mHover = False

                End If
    
    
            End If
        Next i
        
        
        
        
    End If
    
    ' label = <caption>, <name>, y x w h, style(rigth/center/left)
    If vCommand = "LABLE" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vLblCaption = vCom1
        
        SpaceSplit (vCom3)
        vLblY = vSpc1
        vLblX = vSpc2
        vLblW = vSpc3
        vLblH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyLabel frmCloned(i).Label1, vLblCaption, vLblY, vLblX, vLblW, vLblH, vCom2
            End If
        Next i
     End If
     
     ' text = <text>, <name>, y x w h, style(rigth/center/left)
     If vCommand = "TEXT" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vTxtText = vCom1
        
        SpaceSplit (vCom3)
        vTxtY = vSpc1
        vTxtX = vSpc2
        vTxtW = vSpc3
        vTxtH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyText frmCloned(i).Text1, vTxtText, vTxtY, vTxtX, vTxtW, vTxtH, vCom2
            End If
        Next i
     End If
     
     ' check = <caption>, <name>, y x w h, style()
     If vCommand = "CHECK" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vChkCaption = vCom1
        
        SpaceSplit (vCom3)
        vChkY = vSpc1
        vChkX = vSpc2
        vChkW = vSpc3
        vChkH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyCheck frmCloned(i).Check1, vChkCaption, vChkY, vChkX, vChkW, vChkH, vCom2
            End If
        Next i
     End If
     
     ' radio = <caption>, <name>, y x w h, style()
     If vCommand = "RADIO" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vRadCaption = vCom1
        
        SpaceSplit (vCom3)
        vRadY = vSpc1
        vRadX = vSpc2
        vRadW = vSpc3
        vRadH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyRadio frmCloned(i).Option1, vRadCaption, vRadY, vRadX, vRadW, vRadH, vCom2
            End If
        Next i
     End If
     
     ' frame = <caption>, <name>, y x w h, style()
     If vCommand = "FRAME" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vFraCaption = vCom1
        
        SpaceSplit (vCom3)
        vFraY = vSpc1
        vFraX = vSpc2
        vFraW = vSpc3
        vFraH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyFrame frmCloned(i).Frame1, vFraCaption, vFraY, vFraX, vFraW, vFraH, vCom2
            End If
        Next i
     End If
     
     ' link = <link>, <name>, y x w h, style()
    If vCommand = "LINK" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vLnkCaption = vCom1
        
        SpaceSplit (vCom3)
        vLnkY = vSpc1
        vLnkX = vSpc2
        vLnkW = vSpc3
        vLnkH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyLink frmCloned(i).Label2, vLnkCaption, vLnkY, vLnkX, vLnkW, vLnkH, vCom2
            End If
        Next i
     End If

    ' image = <filename>, <name>, x y w h, style()
    If vCommand = "IMAGE" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        vImgFile = vCom1
        
        SpaceSplit (vCom3)
        vImgY = vSpc1
        vImgX = vSpc2
        vImgW = vSpc3
        vImgH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyImage frmCloned(i).Image1, vImgFile, vImgY, vImgX, vImgW, vImgH, vCom2
            End If
        Next i
     End If
     
     ' list = <name>, x y w h, style()
    If vCommand = "List" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        SpaceSplit (vCom2)
        vLstY = vSpc1
        vLstX = vSpc2
        vLstW = vSpc3
        vLstH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyList frmCloned(i).List1, vLstY, vLstX, vLstW, vLstH, vCom1
            End If
        Next i
     End If
     
     ' combo = <name>, x y w h, style()
    If vCommand = "COMBO" And vRunSub = True Then
        EqualsSplit (j(k))
        CommaSplit (vEqu2)
        
        SpaceSplit (vCom2)
        vCmbY = vSpc1
        vCmbX = vSpc2
        vCmbW = vSpc3
        vCmbH = vSpc4
        
        For i = 0 To vFormCount
            If frmCloned(i).Tag = vFormName Then
                CopyCombo frmCloned(i).Combo1, vCmbY, vCmbX, vCmbW, vCmbH, vCom1
            End If
        Next i
     End If
    
    
    ' menu = <text>, <name>
    ' item = <text>, <name>
    ' item = sep, <name>
    If vCommand = "MENU" Or vCommand = "ITEM" Then
        
        MsgBox "The menu and item commands aren't implemented yet", vbInformation
        
    End If
    
    ' msgbox <text>, <type>, <title>
    If vCommand = "MSGBOX" And vRunSub = True Or vCommand = "MSGBOX" And vEvent = True Then
        BracketSplit (j(k))
        CommaSplit (vBrack)
        
        If UCase(vCom2) = "ZCRIT" Or UCase(vCom2) = "ZEXCLA" Or UCase(vCom2) = "ZINFO" Or UCase(vCom2) = "ZDEFAULT" Then
            
            If UCase(vCom2) = "ZCRIT" Then vCom2 = vbCritical
            If UCase(vCom2) = "ZEXCLA" Then vCom2 = vbExclamation
            If UCase(vCom2) = "ZINFO" Then vCom2 = vbInformation
            If UCase(vCom2) = "ZDEFAULT" Then vCom2 = vbOKOnly
      
            MsgBox vCom1, vCom2, vCom3
        
        End If
        
        
        
    End If
    
    ' jos78de.ole.Caption = helloa
    If vCommand Like UCase("*.*.caption*") And vRunSub = True Or vCommand Like UCase("*.*.caption*") And vEvent = True Then
        Dim vObjname As CommandButton
        
        
        EqualsSplit (j(k))
        FullstopSplit (vEqu1)
        vFormName = vStp1
        'vObjname =
        vTitle = vEqu2
        
        For i = 0 To vFormCount
             If frmCloned(i).Tag = vFormName Then
                          
                For a = 0 To frmCloned(i).Count
                    If frmCloned(i).Command1(a).Tag = vStp2 Then frmCloned(i).Command1(a).Caption = vTitle
                    If frmCloned(i).Label1(a).Tag = vStp2 Then frmCloned(i).Label1(a).Caption = vTitle
                    If frmCloned(i).Check1(a).Tag = vStp2 Then frmCloned(i).Check1(a).Caption = vTitle
                    If frmCloned(i).Option1(a).Tag = vStp2 Then frmCloned(i).Option1(a).Caption = vTitle
                    If frmCloned(i).Frame1(a).Tag = vStp2 Then frmCloned(i).Frame1(a).Caption = vTitle
                    If frmCloned(i).Label2(a).Tag = vStp2 Then frmCloned(i).Label2(a).Caption = vTitle
                Next a
             
             End If
        Next i
        
    End If
    
    
    ' jos78de.Caption = helloa
    If vCommand Like UCase("*.caption*") And vRunSub = True Or vCommand Like UCase("*.caption*") And vEvent = True Then
        EqualsSplit (j(k))
        FullstopSplit (vEqu1)
        vFormName = vStp1
        vTitle = vEqu2
        
        If vStp3 <> "" Then GoTo ready
        
        For i = 0 To vFormCount
             If frmCloned(i).Tag = vFormName Then frmCloned(i).Caption = vTitle
        Next i
        
    End If



ready:

CleanVars



If vRunSub = True Then
    If vCommand = "FORM" Then
        frmMain.TV.Nodes.Add , , vFormName, vFormName
    Else
        frmMain.TV.Nodes.Add vFormName, tvwChild, vFormName & j(k), j(k)
    End If
End If

Next k



End Function

