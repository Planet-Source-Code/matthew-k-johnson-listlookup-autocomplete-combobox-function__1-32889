Attribute VB_Name = "modListLookup"
Option Explicit

'********************************************************************************************************************
'
'  Name: ListLookup           (Based on Mike Schaffer's IncrLookup [page 28 of the Visual Basic Source Code Library])
'
'  Last Modified: 3/20/02
'
'  By: Matthew K. Johnson     (mjohnson@austin.rr.com)
'
'  Summary: Performs an incremental matching lookup of a list during key entry on a ComboBox.
'
'  Functionality:  This procedure will allow a Dropdown ComboBox (style 0) or a Simple ComboBox (style 1) to
'    perform a search for a matching list entry as the user types in the data, similar to an Access combo, or
'    the Internet Explorer's autocomplete feature.  Additional options allow the coder to specify that the ComboBox
'    search the list with matching case, or inhibit user data entry similar to a Dropdown List (style 2).
'
'  Information Passed To Procedure:
'    AKey                         Integer                       The ascii keystroke value
'    LimitToList (Optional)       Boolean (Default = False)     Makes combobox act as a Dropdown List
'    CaseSensitive (Optional)     Boolean (Default = False)     Makes searching the list case sensitive
'
'  Returns: <NOTHING>
'
'  References:
'    Screen Object                VB Object
'    ComboBox Object              VB Object
'
'  Usage:  Place this line in the combobox KeyPress event:
'
'    ListLookup KeyAscii[, exprLimitToList[, exprCaseSensitive]]
'
'  Limitations:
'      -  Combobox must be style 0 (Dropdown Combo) or style 1 (Simple Combo) to work properly.
'      -  The delete key functions, but there is no code responding to it's actions at this time.
'      -  The LimitToList will not stop the user from having an empty field.
'
'  Design Notes:
'    Modification to this module was originally started to give backspace functionality and to limit
'    the user to a Dropdown List style ComboBox and progressed from there.  In the process, I added several
'    features, replaced the search routine, and added documentation.
'
'    Because this is a modifaction, and not a planned module, it is not entirely efficient.  I belive that recoding
'    this as a class with better data entry control would result in a better functioning module.  This is planned
'    in the future, but I lack the time right now to spend on this module.
'
'    In the course of the coding, the module lost the ability to require an entry in the combobox (see indicators
'    below).  I cannot identify a reason why the code should not work; if you know the reason why, please let me
'    know.
'
'    The use of the Select statement in the second part of the module is to allow key restriction / additional
'    actions to be added depending on the key input.  Modify to your need.
'
'    The use of this module with a Dropdown combo has been extensively tested; no testing has been performed with
'    a Simple ComboBox.
'
'    With a slight change to the coding of the first section of this module, I believe that this code would work
'    with a listbox also.  I plan to change the code in the future to allow for dual functionaly by replacing the
'    ComboBox with an object.
'
'*******************************************************************************************************************
Public Sub ListLookup(AKey As Integer, Optional ByVal LimitToList As Boolean = False, Optional ByVal CaseSensitive _
 As Boolean = False)
Dim Ctrl As ComboBox                                        'ComboBox to be manipulated
Dim Srch As String                                          'String to be searched for
Dim NLoc As Long                                            'Location of the string in the list if found
Dim ELoc As Long                                            'Old Location of the string in the list
Dim I As Integer                                            'Loop Counter

  
  'Verify that the calling object is a Dropdown or Simple ComboBox, and data is in the list
  
  If TypeOf Screen.ActiveControl Is ComboBox Then           'If active control on screen is a combobox then
    Set Ctrl = Screen.ActiveControl                           'Set Ctrl to the control
    If Ctrl.Style < 0 Or Ctrl.Style > 1 Then                  'If not a Dropdown (0) or Simple (1) ComboBox then
      Exit Sub                                                  'Exit procedure
    End If                                                    'End if
    If Ctrl.ListCount = 0 Then                                'If not items in list for matching then
      Exit Sub                                                  'Exit procedure
    End If                                                    'End if
  Else                                                      'Else
    Exit Sub                                                  'Exit procedure
  End If                                                    'End if
  
  With Ctrl                                                 'Set With to Ctrl
    
    'Determine key pressed to perform appropriate action
    Select Case AKey                                          'Select Key passed to sub
      Case 8                                                    'If backspace was pressed then
        If .SelStart = 0 And .SelLength = 0 Then                  'If at beginning of text and nothing selected
          Exit Sub                                                  'Nothing to do, exit sub
        End If                                                    'End if
        If .SelLength >= Len(.Text) - 1 Then                      'If all text, or all text except begining char selected
          If LimitToList = False Then                               'If Not Limited to list then
            .Text = ""                                                'Set text to nothing
          Else                                                      'Else if limited to list then
            .ListIndex = 0                                            '** NOT WORKING ** Set list index to first item
            .Text = .List(0)                                          '** NOT WORKING ** Text = Text(Location of item)
            .SelStart = 0                                             '** NOT WORKING ** Move cursor to beginning of text
            .SelLength = Len(.Text)                                   '** NOT WORKING ** Select all text
          End If                                                    'End if
          Exit Sub                                                  'Exit procedure
        End If                                                    'End if
        If Len(.Text) = .SelStart + .SelLength Then               'If text is text is selected and begins at far right
           Srch = Left$(.Text, .SelStart - 1)                       'Add the next letter to left of selected text
        Else                                                      'Else if no selected text, or selection is not at end
          If .SelLength = 0 Then                                    'If no text selected
            .Text = Left$(.Text, .SelStart - 1) & Right$(.Text, Len(.Text) - .SelStart)
                                                                      'Delete the character
          Else                                                      'Else is text is selected
            .Text = Left$(.Text, .SelStart) & Right$(.Text, Len(.Text) - .SelStart - .SelLength)
                                                                      'Delete the selected text
          End If                                                    'End if
          Srch = .Text                                              'Set search key to entire text
        End If                                                    'End if
      Case Else                                                 'If any other key pressed
        If Len(.Text) > 1 And .SelLength = 0 And .SelStart <> Len(.Text) Then
                                                                  'If character is inserted into existing string
          Srch = Left$(.Text, .SelStart) & Chr$(AKey) & Right$(.Text, Len(.Text) - .SelStart)
                                                                    'Insert character into middle of search string
        Else                                                      'Else character is inserted at end of string
          Srch = Left$(.Text, .SelStart) & Chr$(AKey)               'Search unselected text and key entered
        End If                                                    'End if
    End Select                                                'End Select
    
    'Search combobox list for matching text
    ELoc = .ListIndex                                         'Get existing list index location
    NLoc = -1                                                 'Set selected item in list to -1
    For I = 0 To .ListCount                                   'For each item in the list
      If CaseSensitive = True Then                              'If case matching required then
        If Srch = Left$(.List(I), Len(Srch)) Then                 'Compare search key with list item
          NLoc = I                                                  'Set listindex to entry
          Exit For                                                  'Exit for
        End If                                                    'End if
      Else                                                      'Else if case matching not required
        If UCase(Srch) = UCase(Left$(.List(I), Len(Srch))) Then   'Compare search key with list item using upper case
          NLoc = I                                                  'Set listindex to entry
          Exit For                                                  'Exit for
        End If                                                    'End if
      End If                                                    'End if
    Next I                                                    'Next i
   
    'Actions to be taken depending on match results
    If NLoc >= 0 Then                                         'If found then
      .ListIndex = NLoc                                         'Set Listindex to Location of item
      .Text = .List(NLoc)                                       'Text = Text(Location of item)
      .SelStart = Len(Srch)                                     'Selection start is end of searched text
      .SelLength = Len(.Text)                                   'Select the remainder of the text
    Else                                                      'Else
      If LimitToList = True Then                                'If text can only be selected from list
        .ListIndex = ELoc                                         'Set Listindex to Location of item
        .Text = .List(ELoc)                                       'Text = Text(Location of item)
        If .SelStart > 0 Then                                     'If cursor not at begining
          .SelStart = Len(Srch) - 1                                 'Selection start is length of search -1
        Else                                                      'Else
          .SelStart = Len(Srch)                                     'Selection start is length of search
        End If                                                    'End if
        .SelLength = Len(.Text)                                   'Select the remainder of the text
      Else                                                      'Else if not restricted to list
        .Text = Srch                                              'Set text to user entry
        .SelStart = Len(Srch)                                     'Move cursor to end of text
      End If                                                   'End if
    End If                                                   'End if
  
  End With                                                 'Set With to Ctrl
  AKey = 0                                                 'Reset the Ascii Key

End Sub



