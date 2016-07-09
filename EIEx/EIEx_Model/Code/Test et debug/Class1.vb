
Imports System.Xml.Serialization
Imports Model

Public Class Class1
    Implements IClasse

    Private _Nom As String
    Property Nom As String
        Get
            Return _Nom
        End Get
        Set(value As String)
            _Nom = value
        End Set
    End Property

    Property INom As String Implements IClasse.Nom

    <XmlIgnore>
    Property Objet As Object Implements IClasse.IObjet2

    <XmlIgnore>
    Property IObjet As IClasse

End Class

Public Class Class2
    Property Nom As String


End Class

Public Interface IClasse
    Property Nom As String
    Property IObjet2 As Object

End Interface
