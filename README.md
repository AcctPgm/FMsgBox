# FMsgBox
Purpose   : Class module for a MsgBox replacement that allows text to be formatted to a degree  
License   : GNU AFFERO GENERAL PUBLIC LICENSE, version 3  
Warranty  : None - Use is at your own risk. VBA code is intended for developers who are 
responsible to ensure it is suitable for their intended use.  
Req'd Refs: frmMsgBox (form)  
            No other external refernces are used  

## Usage:  
a) This will show a message with 'done' in bold and red text, with an OK button.
      Dim fm as clsFMsgBox
      Set fm = new clsFMsgBox
      fm "Congratulatins, you're <b><red>done</red></b>!"

b) This will show a bulletted list, and Yes and No buttons with No as the default
      Dim fm as clsFMsgBox
      Dim Response as vbMsgBoxResult
      set fn = new clsFMsgBox
      With fn
          .FormTitle = "Status"
          .FormButtons = vbYesNo + vbDefaultButton2
          .Msg = "Meetings are set for:" & _
              "<ul>Monday</ul>" & _
              "<ul>Wednesday</ul>" & _
              "<ul>Friday</ul>" & _
              "<br>Do you want to accept the invitation?"
          Response = fn.Dsply
      End With

## Format tags:
Apply formatting to the message (prompt) text by embedding tags in the text. Some tags
are similar to HTML, but this isn't intended to be a full HTML interpreter so they
generally don't follow.

Most tags are paired as <start></stop>, with any text between receiving the format.

### Colours:
The exact RGB values for the colours may be set through FMsgBox properties
  <blue></blue>
  <red></red>
  <green></green>
  <orange></orange>
  <purple></purple>

The default text colour is black. It may be set through FMsgBox properties

### Formats:
  <b></b>                     Bold
  <u></u>                     Underline
  <i></i>                     Italic
  <high></high>               Add a background highlight colour to the text
                              By default the highlight is yellow but may be se
                              through FMsgBox properties

### Line break:
  <br>                        Start a new line
                              vbLf (chr$(10)), vbCr (13), vbCrLf (13+10), and
                              vbNewLine (13+10) are all treated as a <br>

### Tabs:
  <tab>                       Advance to the next tab stop position
                              If the position isn't defined, text will be at the next
                              default tab position based on the text size.
  <tabset>                    Save the current position as a tab stop. This allows
                              taxt to be left-aligned as the position without
                              worrying about the width of preceeding text in the line
  <tabunset>                   Remove the defined stop, reverting to default positiong

Indents and Lists:
~~~~~~~~~~~~~~~~~~
  <indent></indent>           Indent text by the width of four space charactersa
                              The size of the indent, in space characters, may be
                              set through FMsgBox properties
  <li></li>                   Numbered list item.
                              The number increments after each entry
  <ul></ul>                   Bullet list item.
                              The bullet character may be set through FMsgBox properties
