VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFMsgBox 
   Caption         =   "Usage Error"
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2925
   OleObjectBlob   =   "frmFMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'---------------------------------------------------------------------------------------
' Form      : frmFMsgBox
' Author    : Ken Parker
' Website   : https://github.com/acctpgm
' Purpose   : Display a userform with formatted text
' License   : GNU AFFERO GENERAL PUBLIC LICENSE, version 3
' Warranty  : None - Use is at your own risk. VBA code is intended for developers who
'             are responsible to ensure it is suitable for their intended use.
' Req'd Refs: Module - modFMsgBox has the FMsgBox instance of clsFMsgBox used by the form
'             Class - clsFMsgBox
'
' Usage:
' ~~~~~~
' Show        automatically created the form, which pulls values from the clsFMsgBox
'             class in the form's Initialize event
'
' Revision History:
' Rev       Date(yyyy/mm/dd)    Description
' **************************************************************************************
' 1         2020-05-10          Initial Release
' 2			2020-05-13			Fix width when max form width is less than button width
'---------------------------------------------------------------------------------------

'---------------- Internal constants ---------------------------------------------------

Private Const LineBreak     As String = "<br>"      ' Line break tag
Private Const TabTag        As String = "<tab>"     ' Tab tab

'---------------- Internal values ------------------------------------------------------
Private FMsgBox             As clsFMsgBox           ' Controlling instance of clsFMsgBox

Private WordCount           As Long                 ' Count of words, used to ID Label controls

Private lblWord             As MSForms.Label        ' Current lablel being added to the form

Private MsgTop              As String               ' Top for the next message element
Private MsgLeft             As String               ' Left for the next message element
Private MsgColour           As Long                 ' Font colour
Private MsgHighlight        As Long                 ' Highlight (background) colour
Private MsgBold             As Boolean              ' Bold text
Private MsgUnderline        As Boolean              ' Underline text
Private MsgItalic           As Boolean              ' Italicize text

Private LineHeight          As Single               ' Height of the line

Private ButtonWidth         As Single               ' Width of form button
Private ButtonCount         As Integer              ' Number of buttons displayed
Private CanClose            As Boolean              ' Flag whether red X closes form

Private TabStop             As Single               ' Position after <tabset>
Private SpaceCharWidth      As Single               ' Width of a space character

Private LeftMargin          As Single               ' Left text margin, including indent

Private NextElement         As String               ' Next word to display

Private NumListCounter      As String               ' Counter for numbered list

Private WidthPad            As String               ' Extra width added to label by AutoSize

' ------------------------------------------------------------------------------
' Initialize the form
' ------------------------------------------------------------------------------

Public Sub Dsply(ByVal oFMsgBox As clsFMsgBox)
    
    On Error GoTo on_error
    
    Set FMsgBox = oFMsgBox
    
    lblWarning.Visible = False
    
    WordCount = 0
    
    NumListCounter = 1
    
    MsgBold = False
    MsgUnderline = False
    MsgItalic = False
    
    CalculateWidthPad
    CalculateSpaceWidth

    LineHeight = 0
    
    TabStop = 0
    
    Me.Caption = FMsgBox.FormTitle
    
    MsgColour = FMsgBox.DefaultColour
    
    LeftMargin = FMsgBox.MarginLeft
    
    Me.Width = FMsgBox.FormWidthMin
    
    MsgHighlight = Me.BackColor
    
    MsgFormButtonsEnable
    
    MsgFormText
    
    Me.Height = (Me.Height - Me.InsideHeight) + MsgTop + FMsgBox.MarginBottom + _
        IIf(Not lblWord Is Nothing, lblWord.Height, 0) + FMsgBox.ButtonTopGap + _
        CommandButton1.Height
    
    MsgFormButtonsPosition
    
    PositionForm

    Show
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.Initialize'"
End Sub

' ------------------------------------------------------------------------------
' Show/hide buttons and set button captions according to clsFMsgBox Buttons
' ------------------------------------------------------------------------------

Private Sub MsgFormButtonsEnable()

    On Error GoTo on_error

    ' ------------------------------------------------------------------------------
    ' By default, Buttons 2 and 3 are not visible
    ' ------------------------------------------------------------------------------
    
    CommandButton2.Visible = False
    CommandButton3.Visible = False
    
    ' ------------------------------------------------------------------------------
    ' Enable and caption the buttons according to the Buttons parameter
    '
    ' The "And &HFF" masks off bits that do something other than specifying which
    ' buttons are on the form, such as setting the default button.
    '
    ' Assign the value of the button to the Tag so that, once clicked, the button
    ' can assign the tag value as the class result
    ' ------------------------------------------------------------------------------
    
    Select Case (FMsgBox.FormButtons And &HFF)
    
    Case vbOKOnly:
        CommandButton1.Caption = "OK"
        CommandButton1.Tag = vbOK
        CommandButton1.Cancel = True
        CommandButton1.Visible = True
        
        ButtonCount = 1
    Case vbOKCancel:
        CommandButton1.Caption = "OK"
        CommandButton1.Tag = vbOK
        CommandButton1.Visible = True
    
        CommandButton2.Caption = "Cancel"
        CommandButton2.Tag = vbCancel
        CommandButton2.Cancel = True
        CommandButton2.Visible = True
    
        ButtonCount = 2
    Case vbAbortRetryIgnore:
        CommandButton1.Caption = "Abort"
        CommandButton1.Tag = vbAbort
        CommandButton1.Accelerator = "A"
        CommandButton1.Visible = True
        
        CommandButton2.Caption = "Retry"
        CommandButton2.Tag = vbRetry
        CommandButton2.Accelerator = "R"
        CommandButton2.Visible = True
    
        CommandButton3.Caption = "Ignore"
        CommandButton3.Tag = vbIgnore
        CommandButton3.Accelerator = "I"
        CommandButton3.Visible = True
    
        ButtonCount = 3
    Case vbYesNoCancel:
        CommandButton1.Caption = "Yes"
        CommandButton1.Tag = vbYes
        CommandButton1.Accelerator = "Y"
        CommandButton1.Visible = True
        
        CommandButton2.Caption = "No"
        CommandButton2.Tag = vbNo
        CommandButton2.Accelerator = "N"
        CommandButton2.Visible = True
        
        CommandButton3.Caption = "Cancel"
        CommandButton3.Tag = vbCancel
        CommandButton3.Cancel = True
        CommandButton3.Visible = True
    
        ButtonCount = 3
    Case vbYesNo:
        CommandButton1.Caption = "Yes"
        CommandButton1.Tag = vbYes
        CommandButton1.Accelerator = "Y"
        CommandButton1.Visible = True
        
        CommandButton2.Caption = "No"
        CommandButton2.Tag = vbNo
        CommandButton2.Accelerator = "N"
        CommandButton2.Visible = True
        
        ButtonCount = 2
    Case vbRetryCancel:
        CommandButton1.Caption = "Retry"
        CommandButton1.Tag = vbRetry
        CommandButton1.Accelerator = "R"
        CommandButton1.Visible = True
    
        CommandButton2.Caption = "Cancel"
        CommandButton2.Tag = vbCancel
        CommandButton2.Cancel = True
        CommandButton2.Visible = True
    
        ButtonCount = 2
    End Select
    
    ' ------------------------------------------------------------------------------
    ' Set default button. If not specified, button 1 is the default
    ' ------------------------------------------------------------------------------
    
    If ButtonCount > 2 And ((FMsgBox.FormButtons And vbDefaultButton4) = vbDefaultButton3) Then
        CommandButton3.Default = True
        CommandButton3.SetFocus
    ElseIf ButtonCount > 1 And ((FMsgBox.FormButtons And vbDefaultButton4) = vbDefaultButton2) Then
        CommandButton2.Default = True
        CommandButton2.SetFocus
    Else
        CommandButton1.Default = True
        CommandButton1.SetFocus
    End If
    
    ' ------------------------------------------------------------------------------
    ' Reset initial form width if necessary to accommodate visible buttons
    ' New width may exceed specified maximum width
    ' ------------------------------------------------------------------------------
    
    If Me.Width < (FMsgBox.MarginLeft * 2 + FMsgBox.MarginRight + _
    ((ButtonCount - 1) * FMsgBox.ButtonGapMin) + (ButtonCount * CommandButton1.Width)) Then
        Me.Width = FMsgBox.MarginLeft * 2 + FMsgBox.MarginRight + _
            ((ButtonCount - 1) * FMsgBox.ButtonGapMin) + (ButtonCount * CommandButton1.Width)
        
        If Me.Width > FMsgBox.FormWidthMax Then FMsgBox.FormWidthMax = Me.Width
    End If
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.MsgFormButtonsEnable'"
End Sub

' ------------------------------------------------------------------------------
' Position buttons after text has been processed based on the number of visible
' buttons and the form width and height
' ------------------------------------------------------------------------------

Private Sub MsgFormButtonsPosition()

    On Error GoTo on_error
    
    ' ------------------------------------------------------------------------------
    ' Horizontal adjustment
    ' ------------------------------------------------------------------------------
    
    CommandButton1.Left = (Me.Width - (CommandButton1.Width * ButtonCount) - _
        (FMsgBox.ButtonGapMin * (ButtonCount - 1)) - FMsgBox.MarginRight) / 2

    If ButtonCount > 2 Then
        CommandButton2.Left = CommandButton1.Left + CommandButton1.Width + FMsgBox.ButtonGapMin
        CommandButton3.Left = CommandButton2.Left + CommandButton1.Width + FMsgBox.ButtonGapMin
    ElseIf ButtonCount > 1 Then
        CommandButton2.Left = CommandButton1.Left + CommandButton1.Width + FMsgBox.ButtonGapMin
    End If
    
    ' ------------------------------------------------------------------------------
    ' Vertical adjustment
    ' ------------------------------------------------------------------------------

    CommandButton1.Top = Me.InsideHeight - CommandButton1.Height - FMsgBox.MarginBottom
    CommandButton2.Top = CommandButton1.Top
    CommandButton3.Top = CommandButton1.Top
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.MsgFormButtonsPosition'"
End Sub

' ------------------------------------------------------------------------------
' Break message into elements - words and tags - and send each to processing
' ------------------------------------------------------------------------------

Private Sub MsgFormText()

    On Error GoTo on_error
    
    Dim MsgText As String
    Dim TagEndPos As Long
    
    MsgTop = FMsgBox.MarginTop
    MsgLeft = LeftMargin
    
    ' ------------------------------------------------------------------------------
    ' Convert CrLF sequences to <br>, then convert stand-alone Cr to <br>, then
    ' convert Lf to <br>
    ' Convert tab characters to <tab> tags
    ' ------------------------------------------------------------------------------
    
    MsgText = Replace$(FMsgBox.Msg, vbNewLine, LineBreak)
    MsgText = Replace$(MsgText, vbCr, LineBreak)
    MsgText = Replace$(MsgText, vbLf, LineBreak)
    
    MsgText = Replace$(MsgText, vbTab, TabTag)
    
    ' ------------------------------------------------------------------------------
    ' Loop through all characters in Msg string
    ' ------------------------------------------------------------------------------
    
    Dim ndx As Long
    Dim ElementStart As Long: ElementStart = 1
    
    For ndx = 1 To Len(MsgText)
        ' ------------------------------------------------------------------------------
        ' If the character is a tag start '<', see if there is a tag end '>' and if so
        ' process the tag.
        ' If there's no terminating '>' process the < as an element.
        ' ------------------------------------------------------------------------------
        
        Do While Mid$(MsgText, ndx, 1) = "<"
            NextElement = Mid$(MsgText, ElementStart, ndx - ElementStart)
            
            ProcessMsgWord
            
            TagEndPos = InStr(ndx, MsgText, ">")
            If TagEndPos > 0 Then
                ProcessTag (Mid$(MsgText, ndx, TagEndPos - ndx + 1))
                
                ndx = TagEndPos + 1
            Else
                NextElement = "<"
                
                ProcessMsgWord
                
                ndx = ndx + 1
            End If
                    
            ElementStart = ndx
        Loop
        
        ' ------------------------------------------------------------------------------
        ' Process the segment ended by a space character as a word, the process a
        ' chr$(160) to separate from the next word.
        ' ------------------------------------------------------------------------------
        
        If Mid$(MsgText, ndx, 1) = " " Or ndx = Len(MsgText) Then
        
            If ndx = Len(MsgText) Then
                NextElement = Mid$(MsgText, ElementStart, ndx - ElementStart + 1)
            Else
                NextElement = Mid$(MsgText, ElementStart, ndx - ElementStart)
            End If
            
            ProcessMsgWord
            
            NextElement = Chr$(160)
            ProcessMsgWord
            
            ElementStart = ndx + 1
        End If
    Next ndx
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.MsgFormText'"
End Sub

' ------------------------------------------------------------------------------
' Process the next word in the message
' ------------------------------------------------------------------------------

Private Sub ProcessMsgWord()
    
    On Error GoTo on_error
    
    Dim NeedNewLine As Boolean: NeedNewLine = False
    
    If Len(NextElement) > 0 Then
        Set lblWord = Me.Controls.Add("Forms.Label.1", "W" & WordCount, True)
    
        lblWord.Left = MsgLeft
        lblWord.Top = MsgTop
        
        lblWord.ForeColor = MsgColour
        lblWord.BackColor = MsgHighlight
        lblWord.Font.Size = FMsgBox.FontSize
        lblWord.Parent.Controls("W" & WordCount).Font.Bold = MsgBold
        lblWord.Parent.Controls("W" & WordCount).Font.Underline = MsgUnderline
        lblWord.Parent.Controls("W" & WordCount).Font.Italic = MsgItalic
        
        ' ------------------------------------------------------------------------------
        ' Set the label width to the maximum form width before adding the element text
        ' to avoid having the label wrap if text is longer than the default label.
        ' ------------------------------------------------------------------------------
        
        lblWord.AutoSize = False
        lblWord.Width = FMsgBox.FormWidthMax
        lblWord.Caption = NextElement
        lblWord.AutoSize = True
        
        lblWord.Visible = True
        
        If (MsgLeft + lblWord.Width + LeftMargin + FMsgBox.MarginRight) > Me.Width Then
            If (MsgLeft + lblWord.Width + FMsgBox.MarginLeft + FMsgBox.MarginRight) > FMsgBox.FormWidthMax Then
                Me.Width = FMsgBox.FormWidthMax
                StartNextLine
                NeedNewLine = False
            Else
                Me.Width = MsgLeft + lblWord.Width + FMsgBox.MarginLeft + FMsgBox.MarginRight
            End If
        End If
        
        MsgLeft = MsgLeft + lblWord.Width - WidthPad
        
        If lblWord.Height > LineHeight Then LineHeight = lblWord.Height
        
        WordCount = WordCount + 1
        
    End If
    
    If NeedNewLine And Not lblWord Is Nothing Then
        StartNextLine
    End If
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.ProcessMsgWord() - Element'"
End Sub

' ------------------------------------------------------------------------------
' Start next line by moving the top of the next word down by the height of the
' label and the left of the next word to the left margin
' ------------------------------------------------------------------------------

Private Sub StartNextLine()

    On Error GoTo on_error
    
    lblWord.Left = LeftMargin
    MsgLeft = lblWord.Left
    
    lblWord.Top = lblWord.Top + LineHeight * FMsgBox.LineSpacing
    MsgTop = lblWord.Top
    
    LineHeight = lblWord.Height
    
    If Left$(lblWord.Caption, 1) = " " Then
        lblWord.Caption = Mid$(lblWord.Caption, 2, Len(lblWord.Caption) - 1)
    End If
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.StartNextLine()'"
End Sub

' ------------------------------------------------------------------------------
' Process a formatting tag
' ------------------------------------------------------------------------------

Private Sub ProcessTag(ByVal FormatTag As String)

    On Error GoTo on_error
    
    Select Case LCase$(FormatTag)
    
    Case "<blue>":
        MsgColour = FMsgBox.Blue
    Case "</blue>":
        If MsgColour = FMsgBox.Blue Then MsgColour = FMsgBox.DefaultColour
    
    Case "<red>":
        MsgColour = FMsgBox.Red
    Case "</red>":
        If MsgColour = FMsgBox.Red Then MsgColour = FMsgBox.DefaultColour
    
    Case "<green>":
        MsgColour = FMsgBox.Green
    Case "</green>":
        If MsgColour = FMsgBox.Green Then MsgColour = FMsgBox.DefaultColour
    
    Case "<purple>":
        MsgColour = FMsgBox.Purple
    Case "</purple>":
        If MsgColour = FMsgBox.Purple Then MsgColour = FMsgBox.DefaultColour
    
    Case "<orange>":
        MsgColour = FMsgBox.Orange
    Case "</orange>":
        If MsgColour = FMsgBox.Orange Then MsgColour = FMsgBox.DefaultColour
    
    Case "<high>":
        MsgHighlight = FMsgBox.Highlight
    Case "</high>":
        MsgHighlight = Me.BackColor
    
    Case "<b>":
        MsgBold = True
    Case "</b>":
        MsgBold = False
    Case "<u>":
        MsgUnderline = True
    Case "</u>":
        MsgUnderline = False
    Case "<i>":
        MsgItalic = True
    Case "</i>":
        MsgItalic = False

    Case LineBreak:
        If Not lblWord Is Nothing Then
            MsgTop = MsgTop + lblWord.Height * FMsgBox.LineSpacing
            MsgLeft = LeftMargin
        End If
    
    Case "<tabset>":
        TabStop = MsgLeft
        
    Case "<tabunset>":
        TabStop = 0
    Case "<tab>"
        If TabStop > 0 Then
            MsgLeft = TabStop
        Else
            NextElement = vbTab
            ProcessMsgWord
        End If
        
    Case "<indent>":
        MsgTop = MsgTop + IIf(Not lblWord Is Nothing, lblWord.Height, 0)
        LeftMargin = FMsgBox.MarginLeft + SpaceCharWidth * FMsgBox.IndentSize
        MsgLeft = LeftMargin
    Case "</indent>":
        MsgTop = MsgTop + IIf(Not lblWord Is Nothing, lblWord.Height, 0)
        LeftMargin = FMsgBox.MarginLeft
    
    Case "<li>":
        ListEntry Format(NumListCounter, "0") & "."
        NumListCounter = NumListCounter + 1
    Case "</li>":
        MsgTop = MsgTop + IIf(Not lblWord Is Nothing, lblWord.Height, 0)
        LeftMargin = FMsgBox.MarginLeft
        MsgLeft = LeftMargin

    Case "<ul>":
        ListEntry FMsgBox.BulletChar
    Case "</ul>":
        MsgTop = MsgTop + IIf(Not lblWord Is Nothing, lblWord.Height, 0)
        LeftMargin = FMsgBox.MarginLeft
        MsgLeft = LeftMargin

    End Select
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.ProcessTag()'"
End Sub

' ------------------------------------------------------------------------------
' Start a new list entry
' ListMarker is the number (with any desired punctuation), bullet, etc.
' ------------------------------------------------------------------------------

Private Sub ListEntry(ByVal ListMarker As String)
    
    On Error GoTo on_error
    
    If MsgLeft <> LeftMargin Then
        MsgTop = MsgTop + IIf(Not lblWord Is Nothing, lblWord.Height, 0)
    End If
    LeftMargin = FMsgBox.MarginLeft + SpaceCharWidth * FMsgBox.IndentSize
    MsgLeft = LeftMargin
    
    Set lblWord = Me.Controls.Add("Forms.Label.1", "W" & WordCount, True)

    lblWord.Left = FMsgBox.MarginLeft
    lblWord.Top = MsgTop
    
    lblWord.ForeColor = MsgColour
    lblWord.Font.Size = FMsgBox.FontSize
    lblWord.Parent.Controls("W" & WordCount).Font.Bold = MsgBold
    lblWord.Parent.Controls("W" & WordCount).Font.Underline = MsgUnderline
    lblWord.Parent.Controls("W" & WordCount).Font.Italic = MsgItalic
    
    lblWord.Caption = ListMarker
    
    lblWord.Visible = True
    
    lblWord.AutoSize = True
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.ListEntry()'"
End Sub

' ------------------------------------------------------------------------------
' Handle button clicks
'  - assign the button tag to the class result
'  - unload the userform
' ------------------------------------------------------------------------------

Private Sub CommandButton1_Click()
    
    On Error GoTo on_error
    
    FMsgBox.Result = CommandButton1.Tag
    
    Unload Me
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.CommandButton1_Click()'"
    
    On Error Resume Next
    
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    
    On Error GoTo on_error
    
    FMsgBox.Result = CommandButton2.Tag
    
    Unload Me
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.CommandButton2_Click()'"
End Sub

Private Sub CommandButton3_Click()
    
    On Error GoTo on_error
    
    FMsgBox.Result = CommandButton3.Tag
    
    Unload Me
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.CommandButton3_Click()'"
End Sub

' -------------------------------------------------------------------------
' On close - click on the X in upper-right - return the cancel button
' if it is set, but don't allow close if cancel hasn't been set
' -------------------------------------------------------------------------

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    On Error GoTo on_error
    
    If FMsgBox Is Nothing Then
        Cancel = False
    ElseIf CloseMode = 0 Then
        Cancel = False
            
        If CommandButton1.Cancel Then
            CommandButton1_Click
        ElseIf CommandButton2.Cancel Then
            CommandButton2_Click
        ElseIf CommandButton3.Cancel Then
            CommandButton3_Click
        Else
            Cancel = True
        End If
    End If
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.QueryClose()'"
End Sub

' -------------------------------------------------------------------------
' Use manual positioning to centre the form on calling Excel window
' It doesn't just use CenterOwner because that doesn't work correctly
' with multiple screens
' -------------------------------------------------------------------------

Private Sub PositionForm()
    
    On Error GoTo on_error
    
    Me.StartUpPosition = 0              ' Manual positioning
    
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.PositionForm()'"
End Sub

' -------------------------------------------------------------------------
' Calculate the amount of padding added to a label on AutoSize
' The padding is the difference between a label with a caption of AA and
' the width of two individual labels with captions of A
' -------------------------------------------------------------------------

Private Sub CalculateWidthPad()

    On Error GoTo on_error
    
    Dim lblOneA As MSForms.Label
    Dim lblTwoA As MSForms.Label

    Set lblOneA = Me.Controls.Add("Forms.Label.1", "WithSpace", True)
    lblOneA.Font.Size = FMsgBox.FontSize
    lblOneA.Parent.Controls("WithSpace").Font.Bold = MsgBold
    lblOneA.Parent.Controls("WithSpace").Font.Underline = MsgUnderline
    lblOneA.Parent.Controls("WithSpace").Font.Italic = MsgItalic
    lblOneA.Visible = False
    lblOneA.Caption = "A"
    lblOneA.AutoSize = True

    Set lblTwoA = Me.Controls.Add("Forms.Label.1", "WithoutSpace", True)
    lblTwoA.Font.Size = FMsgBox.FontSize
    lblTwoA.Parent.Controls("WithoutSpace").Font.Bold = MsgBold
    lblTwoA.Parent.Controls("WithoutSpace").Font.Underline = MsgUnderline
    lblTwoA.Parent.Controls("WithoutSpace").Font.Italic = MsgItalic
    lblTwoA.Visible = False
    lblTwoA.Caption = "AA"
    lblTwoA.AutoSize = True

    WidthPad = (lblOneA.Width * 2) - lblTwoA.Width

    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.CalculateWidthAdjust()'"
End Sub

' -------------------------------------------------------------------------
' Calculate and save the width of a (leading) space
' The value is used to adjust the location of <tab> if <tabset> was used
' on a space character, i.e. so the text aligns vertically
' -------------------------------------------------------------------------

Private Sub CalculateSpaceWidth()

    On Error GoTo on_error
    
    If True Then
        Dim lblWithSpace As MSForms.Label
        Dim lblWithoutSpace As MSForms.Label
    
        Set lblWithSpace = Me.Controls.Add("Forms.Label.1", "WithSpace", True)
        lblWithSpace.Font.Size = FMsgBox.FontSize
        lblWithSpace.Parent.Controls("WithSpace").Font.Bold = MsgBold
        lblWithSpace.Parent.Controls("WithSpace").Font.Underline = MsgUnderline
        lblWithSpace.Parent.Controls("WithSpace").Font.Italic = MsgItalic
        lblWithSpace.Visible = False
        lblWithSpace.Caption = " X"
        lblWithSpace.AutoSize = True
    
        Set lblWithoutSpace = Me.Controls.Add("Forms.Label.1", "WithoutSpace", True)
        lblWithoutSpace.Font.Size = FMsgBox.FontSize
        lblWithoutSpace.Parent.Controls("WithoutSpace").Font.Bold = MsgBold
        lblWithoutSpace.Parent.Controls("WithoutSpace").Font.Underline = MsgUnderline
        lblWithoutSpace.Parent.Controls("WithoutSpace").Font.Italic = MsgItalic
        lblWithoutSpace.Visible = False
        lblWithoutSpace.Caption = "X"
        lblWithoutSpace.AutoSize = True
    
        SpaceCharWidth = lblWithSpace.Width - lblWithoutSpace.Width
    Else
        Dim lblSpaceWidth As MSForms.Label
        
        Set lblSpaceWidth = Me.Controls.Add("Forms.Label.1", "WithSpace", True)
        lblSpaceWidth.Font.Size = FMsgBox.FontSize
        lblSpaceWidth.Parent.Controls("WithSpace").Font.Bold = MsgBold
        lblSpaceWidth.Parent.Controls("WithSpace").Font.Underline = MsgUnderline
        lblSpaceWidth.Parent.Controls("WithSpace").Font.Italic = MsgItalic
        lblSpaceWidth.Visible = False
        lblSpaceWidth.Caption = " X"
        lblSpaceWidth.AutoSize = True
        
        SpaceCharWidth = lblSpaceWidth.Width - WidthPad
    End If
    
    Exit Sub
on_error:
    Debug.Print "Error in 'frmFMsgBox.CalculateSpaceWidth'"
End Sub
