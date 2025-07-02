Imports System.ComponentModel
Public Class w_adresse
    Implements INotifyPropertyChanged
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
     Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal prop As String)
        anychange = True
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Public anychange As Boolean
    Private _gemeindename As String
    Public Property Gemeindename() As String
        Get
            Return _gemeindename
        End Get
        Set(ByVal Value As String)
            _gemeindename = Value
            OnPropertyChanged("Gemeindename")
        End Set
    End Property
    Private _strasse As String
    Public Property Strasse() As String
        Get
            Return _strasse
        End Get
        Set(ByVal Value As String)
            _strasse = Value
            OnPropertyChanged("Strasse")
        End Set
    End Property
    Private _hausnr As String
    Public Property Hausnr() As String
        Get
            Return _hausnr
        End Get
        Set(ByVal Value As String)
            _hausnr = Value
            OnPropertyChanged("Hausnr")
        End Set
    End Property
    Private _pLZ As String ' von integer auf string geändert wg ddr und urbanke
    Public Property PLZ() As String
        Get
            Return _pLZ
        End Get
        Set(ByVal Value As String)
            _pLZ = Value
            OnPropertyChanged("PLZ")
        End Set
    End Property
    Private _postfach As String
    Public Property Postfach() As String
        Get
            Return _postfach
        End Get
        Set(ByVal Value As String)
            _postfach = Value
            OnPropertyChanged("Postfach")
        End Set
    End Property
    Private _postfachPLZ As String
    Public Property PostfachPLZ() As String
        Get
            Return _postfachPLZ
        End Get
        Set(ByVal Value As String)
            _postfachPLZ = Value
            OnPropertyChanged("PostfachPLZ")
        End Set
    End Property
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

    Sub clear()
        anychange = False
        Gemeindename = ""
        Strasse = ""
        Hausnr = ""
        PLZ = ""
        Postfach = ""
        Quelle = ""
        Bemerkung = ""
        PostfachPLZ = ""
    End Sub
    Overrides Function tostring() As String
        Dim a$ = String.Format("Gemeinde: {0} {1}{2}", PLZ, Gemeindename, vbCrLf)
        a$ = String.Format("{0}Strasse: {1}{2}", a$, Strasse, vbCrLf)
        a$ = String.Format("{0}Hausnr: {1}{2}", a$, Hausnr, vbCrLf)
        a$ = String.Format("{0}Postfach: {1}{2}", a$, Postfach, vbCrLf)
        Return a$
    End Function

    Public Sub New()
        clear()
    End Sub
End Class
