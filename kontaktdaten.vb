Imports System.ComponentModel
Public Class Kontaktdaten
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
 Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal prop As String)
        anychange = True
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Public anychange As Boolean
    Private _bemerkung As String
    Public Property Bemerkung() As String
        Get
            Return _bemerkung
        End Get
        Set(ByVal Value As String)
            _bemerkung = Value
            OnPropertyChanged("Bemerkung")
        End Set
    End Property

    Private _gesellFunktion As String = ""
    Public Property GesellFunktion() As String
        Get
            Return _gesellFunktion
        End Get
        Set(ByVal Value As String)
            _gesellFunktion = Value
            OnPropertyChanged("GesellFunktion")
        End Set
    End Property

    Private _org As w_organisation
    Public Property Org() As w_organisation
        Get
            Return _org
        End Get
        Set(ByVal Value As w_organisation)
            _org = Value
            OnPropertyChanged("Org")
        End Set
    End Property

    Private _elektr As w_fonfax
    Public Property elektr() As w_fonfax
        Get
            Return _elektr
        End Get
        Set(ByVal Value As w_fonfax)
            _elektr = Value
            OnPropertyChanged("elektr")
        End Set
    End Property
    Private _anschrift As w_adresse
    Public Property Anschrift() As w_adresse
        Get
            Return _anschrift
        End Get
        Set(ByVal Value As w_adresse)
            _anschrift = Value
            OnPropertyChanged("Anschrift")
        End Set
    End Property
    Public Property KontaktID() As Integer
    ''' <summary>
    ''' wer hats angelegt
    ''' </summary>
    ''' <remarks></remarks>
    Private _quelle As String
    Public Property Quelle() As String
        Get
            Return _quelle
        End Get
        Set(ByVal Value As String)
            _quelle = Value
            OnPropertyChanged("Quelle")
        End Set
    End Property
    Public Property OrgID() As Integer
    Public Property AnschriftID() As Integer
    Public Property BankkontoID() As Integer

    Private _bankkonto As clsBankverbindungSEPA
    'Public Property Bankkonto() As clsBankverbindungSEPA
    '    Get
    '        Return _bankkonto
    '    End Get
    '    Set(ByVal Value As clsBankverbindungSEPA)
    '        _bankkonto = Value
    '        OnPropertyChanged("Bankkonto")
    '    End Set
    'End Property

    Sub New()
        'Bankkonto = New clsBankverbindungSEPA
        Anschrift = New w_adresse
        elektr = New w_fonfax
        Org = New w_organisation
    End Sub

    Sub clear()
        Anschrift.clear()
        elektr.clear()
        Org.clear()
        GesellFunktion = ""
        Bemerkung = ""
        Quelle = ""
    End Sub
    Overrides Function tostring() As String
        Dim a$ = String.Format("GesellFunktion: {0}{1}", GesellFunktion, vbCrLf)
        a$ = a$ & Org.tostring & vbCrLf
        a$ = a$ & Anschrift.tostring & vbCrLf
        a$ = a$ & elektr.tostring & vbCrLf
        a$ = a$ & Bemerkung & vbCrLf & vbCrLf
        Return a$
    End Function
End Class


