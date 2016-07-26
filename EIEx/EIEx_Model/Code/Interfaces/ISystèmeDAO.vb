
Imports System.Xml.Serialization

Public Interface ISystèmeDAO
    <XmlIgnore>
    Property DateModif As Date
    Sub UnSerialize(NewT As Système)
End Interface

