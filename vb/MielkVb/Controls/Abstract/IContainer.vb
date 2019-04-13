Imports System.Drawing

Public Interface IContainer

    Function GetParent(Optional level As Integer = 0) As IContainer
    Function GetWindow() As Window
    Function GetLevel() As Integer

    Function GetStylesManager() As StylesManager
    'Sub AddControl(ctrl As Control)
    'Sub RemoveControl(ctrl As Control)

    Function GetSizeType(prop As StylePropertyEnum) As ControlSizeTypeEnum
    Function GetHeight() As Single
    Function GetInnerHeight() As Single
    Function GetWidth() As Single
    Function GetInnerWidth() As Single
    'Function CalculateAutoHeight() As Single
    'Function CalculateAutoWidth() As Single

    Function GetPaddingLeft() As Single
    Function GetPaddingRight() As Single
    Function GetPaddingTop() As Single
    Function GetPaddingBottom() As Single

    Sub UpdateHeight()
    Sub UpdateWidth()
    Sub ArrangeControls()

    'Sub UpdateView(Optional propagateDown As Boolean = False)
    'Sub UpdateLayout(Optional propagateDown As Boolean = False)
    'Sub RearrangeControls()
    'Sub ResizeControls()
    'Sub AdjustAfterChildrenSizeChange()

End Interface