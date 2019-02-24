Imports System.Xml

Public Interface IVbaControl
    Sub LoadFromXml(xmlNode As xmlnode)
    Sub SetParent(value As IVbaContainer)
End Interface