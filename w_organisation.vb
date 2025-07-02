Imports System.ComponentModel
Public Class w_organisation
    Implements INotifyPropertyChanged
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
           Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal prop As String)
        anychange = True
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Public anychange As Boolean

    Private _name As String = ""
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Value
            OnPropertyChanged("Name")
        End Set
    End Property
    Private _zusatz As String = ""
    Public Property Zusatz() As String
        Get
            Return _zusatz
        End Get
        Set(ByVal Value As String)
            _zusatz = Value
            OnPropertyChanged("Zusatz")
        End Set
    End Property
    Private _typ1 As String = ""
    Public Property Typ1() As String
        Get
            Return _typ1
        End Get
        Set(ByVal Value As String)
            _typ1 = Value
            OnPropertyChanged("Typ1")
        End Set
    End Property

    Private _typ2 As String = ""
    Public Property Typ2() As String
        Get
            Return _typ2
        End Get
        Set(ByVal Value As String)
            _typ2 = Value
            OnPropertyChanged("Typ2")
        End Set
    End Property
    Private _eigentuemer As String = ""
    Public Property Eigentuemer() As String
        Get
            Return _eigentuemer
        End Get
        Set(ByVal Value As String)
            _eigentuemer = Value
            OnPropertyChanged("Eigentuemer")
        End Set
    End Property
    ''' <summary>
    ''' wer hats angelegt
    ''' </summary>
    ''' <remarks></remarks>
    Private _quelle As String = ""
    Public Property Quelle() As String
        Get
            Return _quelle
        End Get
        Set(ByVal Value As String)
            _quelle = Value
            OnPropertyChanged("Quelle")
        End Set
    End Property
    Private _bemerkung As String = ""
    Public Property Bemerkung() As String
        Get
            Return _bemerkung
        End Get
        Set(ByVal Value As String)
            _bemerkung = Value
            OnPropertyChanged("Bemerkung")
        End Set
    End Property
    Public Property Anschriftid() As Integer

    Sub clear()
        anychange = False
        Name = ""
        Zusatz = ""
        Typ1 = ""
        Typ2 = ""
        Eigentuemer = ""
    End Sub
    Overrides Function tostring() As String
        Dim a$ = String.Format("Name: {0}{1}", Name, vbCrLf)
        a$ = String.Format("{0}Zusatz: {1}{2}", a$, Zusatz, vbCrLf)
        a$ = String.Format("{0}Typ1: {1}{2}", a$, Typ1, vbCrLf)
        a$ = String.Format("{0}Typ2: {1}{2}", a$, Typ2, vbCrLf)
        a$ = String.Format("{0}Eigentuemer: {1}{2}", a$, Eigentuemer, vbCrLf)
        Return a$
    End Function
End Class

