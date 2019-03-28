Imports System.Drawing

Public Interface IControl

    '[Style classes]
    Sub AddStyleClass(name As String)
    Sub RemoveStyleClass(name As String)
    Sub SetStyleProperty(propertyType As Long, value As Object)
    Function GetStylesManager() As StylesManager

    '[Meta]
    Function GetId() As String
    Function GetElementTag() As String
    Sub SetParent(parent As IContainer)
    Function GetParent() As IContainer
    Function GetWindow() As Window

    '[Size and position]
    Sub UpdateView(Optional propagateDown As Boolean = False)
    Sub UpdateLayout(Optional propagateDown As Boolean = False)
    'Sub UpdateSize(Optional ByRef anyChanges As Boolean = False)
    'Sub UpdatePosition(Optional ByRef anyChanges As Boolean = False)
    Function GetWidth() As Single
    Function GetHeight() As Single

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