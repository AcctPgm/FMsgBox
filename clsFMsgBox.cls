VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Class     : clsFMsgBox
' Author    : Ken Parker
' Website   : https://github.com/acctpgm
' Purpose   : Class module for a MsgBox replacement that allows text to be formatted
'             to a degree
' License   : GNU AFFERO GENERAL PUBLIC LICENSE, version 3
' Warranty  : None - Use is at your own risk. VBA code is intended for developers who
'             are responsible to ensure it is suitable for their intended use.
' Req'd Refs: frmMsgBox (form)
'             No other external refernces are used
'
' Usage:
' ~~~~~~
' a) This will show a message with 'done' in bold and red text, with an OK button.
'       Dim fm as clsFMsgBox
'       Set fm = new clsFMsgBox
'       fm "Congratulatins, you're <b><red>done</red></b>!"
'
' b) This will show a bulletted list, and Yes and No buttons with No as the default
'       Dim fm as clsFMsgBox
'       Dim Response as vbMsgBoxResult
'       set fn = new clsFMsgBox
'       With fn
'           .FormTitle = "Status"
'           .FormButtons = vbYesNo + vbDefaultButton2
'           .Msg = "Meetings are set for:" & _
'               "<ul>Monday</ul>" & _
'               "<ul>Wednesday</ul>" & _
'               "<ul>Friday</ul>" & _
'               "<br>Do you want to accept the invitation?"
'           Response = fn.Dsply
'       End With
'
' Revision History:
' Rev       Date(yyyy/mm/dd)    Description
' **************************************************************************************
' 1         2020-04-19          Initial Release
'---------------------------------------------------------------------------------------
'
' Format tags:
' ~~~~~~~~~~~~
' Apply formatting to the message (prompt) text by embedding tags in the text. Some tags
' are similar to HTML, but this isn't intended to be a full HTML interpreter so they
' generally don't follow.

' Most tags are paired as <start></stop>, with any text between receiving the format.
'
' Colours:
' ~~~~~~~~
' The exact RGB values for the colours may be set through FMsgBox properties
'   <blue></blue>
'   <red></red>
'   <green></green>
'   <orange></orange>
'   <purple></purple>
'
' The default text colour is black. It may be set through FMsgBox properties
'
' Formats:
' ~~~~~~~~
'   <b></b>                     Bold
'   <u></u>                     Underline
'   <i></i>                     Italic
'   <high></high>               Add a background highlight colour to the text
'                               By default the highlight is yellow but may be se
'                               through FMsgBox properties
'
' Line break:
' ~~~~~~~~~~~
'   <br>                        Start a new line
'                               vbLf (chr$(10)), vbCr (13), vbCrLf (13+10), and
'                               vbNewLine (13+10) are all treated as a <br>
'
' Tabs:
' ~~~~~
'   <tab>                       Advance to the next tab stop position
'                               If the position isn't defined, text will be at the next
'                               default tab position based on the text size.
'   <tabset>                    Save the current position as a tab stop. This allows
'                               taxt to be left-aligned as the position without
'                               worrying about the width of preceeding text in the line
'   <tabunset>                  Remove the defined stop, reverting to default positiong
'
' Indents and Lists:
' ~~~~~~~~~~~~~~~~~~
'   <indent></indent>           Indent text by the width of four space charactersa
'                               The size of the indent, in space characters, may be
'                               set through FMsgBox properties
'   <li></li>                   Numbered list item.
'                               The number increments after each entry
'   <ul></ul>                   Bullet list item.
'                               The bullet character may be set through FMsgBox properties
'
'---------------- Public Class Properties ----------------------------------------------
'
' Msg               message to display
' MarginTop
' MarginLeft
'
'---------------- Public Class Methods -------------------------------------------------
'
' Dsply             show the form
' Reset             reset properties to their default values

'---------------- Internal constants----------------------------------------------------
'
Private Const NoButtons         As Integer = -1         ' Indicates no button value

'---------------- Default values -------------------------------------------------------
'
' Default values use the public property name, prefixed with 'd'

Private Const dFormTitle        As String = "Microsoft Excel"

Private Const dMarginTop        As Single = 12
Private Const dMarginLeft       As Single = 12
Private Const dMarginRight      As Single = 12
Private Const dMarginBottom     As Single = 9

Private Const dLineSpacing      As Single = 1

Private Const dFormWidthMin     As Single = 108
Private Const dFormWidthMax     As Single = 800

Private Const dBlack            As Long = 0             ' Black - RGB(0, 0, 0)
Private Const dRed              As Long = 255           ' Red - RGB(255, 0, 0)
Private Const dGreen            As Long = 32768         ' Green - RBG(0, 128, 0)
Private Const dBlue             As Long = 16711680      ' Blue - RGB(0, 0, 255)
Private Const dPurple           As Long = 10498160      ' Purple - RGB(112, 48, 160)
Private Const dOrange           As Long = 1271551       ' Orange - RGB(223, 102, 19)
Private Const dDefaultColour    As Long = dBlack        ' Default default text colour

Private Const dHighlight        As Long = vbYellow

Private Const dFontSize         As Single = 9

Private Const dButtons          As Integer = vbOKOnly   ' show OK button if none specified
Private Const dButtonGapMin     As Single = 9
Private Const dButtonGapMax     As Single = 15
Private Const dButtonTopGap     As Single = 18

Private Const dIndentSize       As Single = 4

Private Const dBulletChar       As String = "•"          ' chr(149)

Private Const dTagS             As String = "<"
Private Const dTagE             As String = ">"

'---------------- Properties -----------------------------------------------------------
'
' Properties use the public property name, prefixed with 'p'

Private pResult         As VbMsgBoxResult   ' Button clickec to close the form

Private pFormTitle      As String           ' Form title
Private pMsg            As String           ' Message to show, with formatting codes

Private pMarginTop      As Single           ' Top of first line of message text
Private pMarginLeft     As Single           ' Left margin of message text
Private pMarginRight    As Single           ' Right margin between text and form edge
Private pMarginBottom   As Single           ' Space between bottom of buttons and bottom of form

Private pLineSpacing    As Single           ' Line spacing as a multiple of the height of a word

Private pFormWidthMin   As Single           ' Starting form width
Private pFormWidthMax   As Single           ' Maximum form width -> text wraps after this width

Private pBlack          As Long             ' Text colour values
Private pRed            As Long
Private pGreen          As Long
Private pBlue           As Long
Private pPurple         As Long
Private pOrange         As Long
Private pDefaultColour  As Long             ' Default text colour

Private pHighlight      As Long             ' Text highlight (background) colour

Private pFontSize       As Single           ' Message font size

Private pButtons        As VbMsgBoxStyle    ' Buttons to show on message box
Private pButtonGapMin   As Single           ' Minimum allowed gap between buttons
Private pButtonGapMax   As Single           ' Maximum allowed gap between buttons
Private pButtonTopGap   As Single           ' Space between bottom of text and top of buttons

Private pIndentSize     As Single           ' Width, in space characters, of an indent

Private pBulletChar     As String           ' character for bullet lists

Private pTagS           As String           ' Tag start character
Private pTagE           As String           ' Tag end character

'---------------- Let/Get Methods ------------------------------------------------------
'
' Let/Get methods use the private property name without the initial 'p'

Friend Property Let Result(ByVal r As VbMsgBoxResult)
    pResult = r
End Property

Friend Property Get Result() As VbMsgBoxResult
    Result = pResult
End Property

Friend Property Let FormTitle(ByVal t As String)
    pFormTitle = t
End Property

Friend Property Get FormTitle() As String
    FormTitle = pFormTitle
End Property

Friend Property Let Msg(ByVal m As String)
    pMsg = m
End Property

Friend Property Get Msg() As String
    Msg = pMsg
End Property

Friend Property Let MarginTop(m As Single)
    pMarginTop = m
End Property

Friend Property Get MarginTop() As Single
    MarginTop = pMarginTop
End Property

Friend Property Let MarginLeft(m As Single)
    pMarginLeft = m
End Property

Friend Property Get MarginLeft() As Single
    MarginLeft = pMarginLeft
End Property

Friend Property Let MarginRight(m As Single)
    pMarginRight = m
End Property

Friend Property Get MarginRight() As Single
    MarginRight = pMarginRight
End Property

Friend Property Let MarginBottom(m As Single)
    pMarginBottom = m
End Property

Friend Property Get MarginBottom() As Single
    MarginBottom = pMarginBottom
End Property

Friend Property Let LineSpacing(l As Single)
    pLineSpacing = l
End Property

Friend Property Get LineSpacing() As Single
    LineSpacing = pLineSpacing
End Property

Friend Property Let FormWidthMin(w As Single)
    pFormWidthMin = w
End Property

Friend Property Get FormWidthMin() As Single
    FormWidthMin = pFormWidthMin
End Property

Friend Property Let FormWidthMax(w As Single)
    pFormWidthMax = w
End Property

Friend Property Get FormWidthMax() As Single
    FormWidthMax = pFormWidthMax
End Property

Friend Property Let Black(ByVal c As Long)
    pBlack = c
End Property

Friend Property Get Black() As Long
    Black = pBlack
End Property

Friend Property Let Red(ByVal c As Long)
    pRed = c
End Property

Friend Property Get Red() As Long
    Red = pRed
End Property

Friend Property Let Green(ByVal c As Long)
    pGreen = c
End Property

Friend Property Get Green() As Long
    Green = pGreen
End Property

Friend Property Let Blue(ByVal c As Long)
    pBlue = c
End Property

Friend Property Get Blue() As Long
    Blue = pBlue
End Property

Friend Property Let Purple(ByVal c As Long)
    pPurple = c
End Property

Friend Property Get Purple() As Long
    Purple = pPurple
End Property

Friend Property Let Orange(ByVal c As Long)
    pOrange = c
End Property

Friend Property Get Orange() As Long
    Orange = pOrange
End Property

Friend Property Let DefaultColour(ByVal c As Long)
    pDefaultColour = c
End Property

Friend Property Get DefaultColour() As Long
    DefaultColour = pDefaultColour
End Property

Friend Property Let Highlight(ByVal c As Long)
    pHighlight = c
End Property

Friend Property Get Highlight() As Long
    Highlight = pHighlight
End Property

Friend Property Let FontSize(ByVal s As Single)
    pFontSize = s
End Property

Friend Property Get FontSize() As Single
    FontSize = pFontSize
End Property

Friend Property Let FormButtons(ByVal b As VbMsgBoxStyle)
    pButtons = b
End Property

Friend Property Get FormButtons() As VbMsgBoxStyle
    FormButtons = pButtons
End Property

Friend Property Let ButtonGapMin(ByVal b As Single)
    pButtonGapMin = b
End Property

Friend Property Get ButtonGapMin() As Single
    ButtonGapMin = pButtonGapMin
End Property

Friend Property Let ButtonGapMax(ByVal b As Single)
    pButtonGapMax = b
End Property

Friend Property Get ButtonGapMax() As Single
    ButtonGapMax = pButtonGapMax
End Property

Friend Property Let ButtonTopGap(ByVal b As Single)
    pButtonTopGap = b
End Property

Friend Property Get ButtonTopGap() As Single
    ButtonTopGap = pButtonTopGap
End Property

Friend Property Let IndentSize(ByVal i As Single)
    pIndentSize = i
End Property

Friend Property Get IndentSize() As Single
    IndentSize = pIndentSize
End Property

Friend Property Let BulletChar(ByVal b As String)
    pBulletChar = b
End Property

Friend Property Get BulletChar() As String
    BulletChar = pBulletChar
End Property

Friend Property Let TagS(ByVal t As String)
    pTagS = t
End Property

Friend Property Get TagS() As String
    TagS = pTagS
End Property

Friend Property Let TagE(ByVal t As String)
    pTagE = t
End Property

Friend Property Get TagE() As String
    TagE = pTagE
End Property

'---------------------------------------------------------------------------------------
'
' Methods

Private Sub Class_Initialize()
    Reset
End Sub

'---------------------------------------------------------------------------------------
' Reset properties to their default values
'---------------------------------------------------------------------------------------

Public Sub Reset()
    FormTitle = dFormTitle
    
    MarginTop = dMarginTop
    MarginLeft = dMarginLeft
    MarginRight = dMarginRight
    MarginBottom = dMarginBottom
    
    LineSpacing = dLineSpacing
    
    FormWidthMin = dFormWidthMin
    FormWidthMax = dFormWidthMax
    
    Black = dBlack
    Red = dRed
    Green = dGreen
    Blue = dBlue
    Purple = dPurple
    Orange = dOrange
    DefaultColour = dDefaultColour
    
    Highlight = dHighlight
    
    FontSize = dFontSize

    FormButtons = dButtons
    ButtonGapMin = dButtonGapMin
    ButtonGapMax = dButtonGapMax
    ButtonTopGap = dButtonTopGap
    
    IndentSize = dIndentSize
    
    BulletChar = dBulletChar
    
    TagS = dTagS
    TagE = dTagE
End Sub

'---------------------------------------------------------------------------------------
' Show the form and return the value of the button used to close the form
'---------------------------------------------------------------------------------------

Public Function Dsply(Optional ByVal Prompt As String = vbNullString, _
                      Optional ByVal Buttons As VbMsgBoxStyle = NoButtons, _
                      Optional ByVal Title As String = vbNullString, _
                      Optional ByVal ResetAfter As Boolean = True) As VbMsgBoxResult
                      
    If Prompt <> vbNullString Then Msg = Prompt
    If Buttons <> NoButtons Then FormButtons = Buttons
    If Title <> vbNullString Then FormTitle = Title
    
    Dim Form As frmFMsgBox
    Set Form = New frmFMsgBox
    
    Form.Dsply Me
    
    Dsply = Result
    
    If ResetAfter Then Reset
End Function

