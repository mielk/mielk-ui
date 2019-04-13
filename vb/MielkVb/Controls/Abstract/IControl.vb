Imports System.Drawing

Public Interface IControl

    '[Style classes]
    Sub SetStyleClasses(name As String, Optional invalidate As Boolean = True)
    Sub AddStyleClasses(name As String, Optional invalidate As Boolean = True)
    Sub RemoveStyleClasses(name As String, Optional invalidate As Boolean = True)
    Sub SetStyleProperty(propertyType As Long, value As Object, Optional invalidate As Boolean = True)
    Function GetStylesManager() As StylesManager

    '[Meta]
    Function GetId() As String
    Function GetElementTag() As String
    Sub SetParent(parent As IContainer)
    Function GetParent(Optional level As Integer = 0) As IContainer
    Function GetWindow() As Window
    Function GetLevel() As Integer

    '[Size and position]
    Function IsRendered() As Boolean
    Function GetWidth() As Single
    Function GetHeight() As Single
    Function GetSizeType(prop As StylePropertyEnum) As ControlSizeTypeEnum
    Function GetMargin() As Coordinate

    Sub UpdateHeight()
    Sub UpdateWidth()
    Sub UpdateView(prop As StylePropertyEnum)

    Function IsAbsolutePositioned() As Boolean
    Function GetWidthSizeType() As ControlSizeTypeEnum
    Function GetHeightSizeType() As ControlSizeTypeEnum
    Sub LocateInContainer(ByRef top As Single, ByRef right As Single, ByRef bottom As Single, ByRef left As Single)

    'Sub UpdateView(Optional propagateDown As Boolean = False)
    'Sub UpdateLayout(Optional propagateDown As Boolean = False)
    'Sub AddToResizeQueue(prop As StylePropertyEnum)



    ''[Setting inline properties]
    'Sub SetTop(value As VariantType?)
    'Sub SetLeft(value As VariantType?)
    'Sub SetBottom(value As VariantType?)
    'Sub SetRight(value As VariantType?)

    'Sub SetWidth(value As VariantType?)
    'Sub SetHeight(value As VariantType?)

    'Sub SetMargins(value As VariantType?)
    'Sub SetMarginTop(value As VariantType?)
    'Sub SetMarginLeft(value As VariantType?)
    'Sub SetMarginBottom(value As VariantType?)
    'Sub SetMarginRight(value As VariantType?)
    'Sub SetPaddings(value As VariantType?)
    'Sub SetPaddingTop(value As VariantType?)
    'Sub SetPaddingLeft(value As VariantType?)
    'Sub SetPaddingBottom(value As VariantType?)
    'Sub SetPaddingRight(value As VariantType?)

    'Sub SetBackgroundColor(color As Long, Optional opacity As Single = 1)
    'Sub SetBorderColor(color As Long, Optional opacity As Single = 1)
    'Sub SetBorderThickness(value As VariantType?)

    'Sub SetFontSize(value As VariantType?)
    'Sub SetFontFamily(value As VariantType?)
    'Sub SetFontBold(value As VariantType?)
    'Sub SetFontColor(value As VariantType?)

End Interface