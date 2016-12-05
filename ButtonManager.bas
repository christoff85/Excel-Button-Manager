Attribute VB_Name = "ButtonManager"
'----------------------------------------------------------------------------------------------------------------
'   Provides abstraction for Excel button object
'   Enables to easily add new buttons to the Excel contextMenu
'
'   Written By Krzysztof Grzeslak 05/11/2015
'
'   Preconditions:
'   *   Excel macro file must include all three cooperating modules: c_Button, c_ButtonCounter and ButtonManager
'   *   Excel macro file This_Workbook Object should include following statements:
'       Workbook_Activate: Call ButtonManager.RemoveButtons, Call ButtonManager.AddButtons
'       Workbook_Deactivate: Call ButtonManager.RemoveButtons
'
'   Usage:
'   *   Functions RowButtonDataArray and CellButtonDataArray should be modified to hardcode the buttons properties
'   *   Module includes only the Cell and Row buttons. Columns can be added by modified parts of the code.
'----------------------------------------------------------------------------------------------------------------

Option Explicit

' Function is a list of hardcoded row menu button data. Array is created from the list and provided as output.
Private Function RowButtonDataArray() As Variant

    ' Template: Array (Macro Name, Button Caption, Button Tag enum, Button Face ID enum)
    Dim tempRowButtonArray() As Variant
    tempRowButtonArray = Array( _
        Array("AddRow", "Insert Row(s)", e_buttonTag_addRow, e_buttonFace_Menu), _
        Array("DeleteRow", "Delete Row(s)", e_buttonTag_deleteRow, e_buttonFace_Menu) _
    ) ' ... before the ')'
    
    RowButtonDataArray = tempRowButtonArray

End Function

' Function is a list of hardcoded cell menu button data. Array is created from the list and provided as output.
Private Function CellButtonDataArray() As Variant

    ' Template: Array (Macro Name, Button Caption, Button Tag enum, Button Face ID enum)
    Dim tempCellButtonArray() As Variant
    tempCellButtonArray = Array( _
        Array("ChangeMode", "Change Mode", e_buttonTag_modeChange, e_buttonFace_Change), _
        Array("Sheetgen", "Generate new Model sheets", e_buttonTag_tpdMode, e_buttonFace_Menu) _
    ) ' ... before the ')'
    
    CellButtonDataArray = tempCellButtonArray

End Function

Public Sub AddButtons(Optional internalProcedure As Boolean = True)
    
    Dim buttonCnt As C_ButtonCounter
    Set buttonCnt = New C_ButtonCounter 'Start new button counter
    
    Dim rowButtons
    rowButtons = RowButtonDataArray 'retrieve all row buttons data and create array
    Call AddButtonGroup(RowButtonDataArray, e_contextMenu_Row, buttonCnt)
    
    Dim cellButtons
    cellButtons = CellButtonDataArray 'retrieve all cell buttons data and create array
    Call AddButtonGroup(CellButtonDataArray, e_contextMenu_Cell, buttonCnt)
    
    Set buttonCnt = Nothing
    
End Sub

Public Sub RemoveButtons(Optional internalProcedure As Boolean = True)
    
    Dim button As C_Button
    Set button = New C_Button
    
    Call button.DeleteAllButtons
    Set button = Nothing
    
End Sub

' Extracts buttons information from array and adds them to chosen contextMenu
Private Sub AddButtonGroup(ByVal buttonArray As Variant, ByVal menu As e_contextMenu, ByRef counter As C_ButtonCounter)

    Dim button As C_Button
    Set button = New C_Button
    
    Dim macroName As String
    Dim buttonCaption As String
    Dim buttonTag As e_buttonTag
    Dim buttonFace As e_buttonFace
    Dim buttonIndex As Integer
    
    For buttonIndex = LBound(buttonArray) To UBound(buttonArray)
        macroName = buttonArray(buttonIndex)(0)
        buttonCaption = buttonArray(buttonIndex)(1)
        buttonTag = buttonArray(buttonIndex)(2)
        buttonFace = buttonArray(buttonIndex)(3)

        With button
            Call .Initialize(macroName, buttonCaption, buttonTag, buttonFace)
            Call .AddToMenu(menu, counter)
        End With
    Next buttonIndex
    
    Call button.CreateGroup(menu, counter)
    
    Set button = Nothing
End Sub

