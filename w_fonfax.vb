Imports System.ComponentModel
Public Class w_fonfax
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
     Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal prop As String)
        anychange = True
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Public anychange As Boolean

    Private _telefon1 As String
    Public Property Telefon1() As String
        Get
            Return _telefon1
        End Get
        Set(ByVal Value As String)
            _telefon1 = Value
            OnPropertyChanged("Telefon1")
        End Set
    End Property
    Private _telefon2 As String
    Public Property Telefon2() As String
        Get
            Return _telefon2
        End Get
        Set(ByVal Value As String)
            _telefon2 = Value
            OnPropertyChanged("Telefon2")
        End Set
    End Property
    Private _fax1 As String
    Public Property Fax1() As String
        Get
            Return _fax1
        End Get
        Set(ByVal Value As String)
            _fax1 = Value
            OnPropertyChanged("Fax1")
        End Set
    End Property
    Private _fax2 As String
    Public Property Fax2() As String
        Get
            Return _fax2
        End Get
        Set(ByVal Value As String)
            _fax2 = Value
            OnPropertyChanged("Fax2")
        End Set
    End Property
    Private _mobilFon As String
    Public Property MobilFon() As String
        Get
            Return _mobilFon
        End Get
        Set(ByVal Value As String)
            _mobilFon = Value
            OnPropertyChanged("MobilFon")
        End Set
    End Property
    Private _email As String
    Public Property Email() As String
        Get
            Return _email
        End Get
        Set(ByVal Value As String)
            _email = Value
            OnPropertyChanged("Email")
        End Set
    End Property
    Private _homepage As String
    Public Property Homepage() As String
        Get
            Return _homepage
        End Get
        Set(ByVal Value As String)
            _homepage = Value
            OnPropertyChanged("Homepage")
        End Set
    End Property
    Public Property Quelle() As String
    Sub clear()
        anychange = False
        Telefon1 = ""
        Telefon2 = ""
        Fax1 = ""
        Fax2 = ""
        MobilFon = ""
        Email = ""
        Homepage = ""
    End Sub
    Overrides Function tostring() As String
        Dim a$ = String.Format("Telefon1: {0}{1}", Telefon1, vbCrLf)
        a$ = String.Format("{0}Telefon2: {1}{2}", a$, Telefon2, vbCrLf)
        a$ = String.Format("{0}Fax1: {1}{2}", a$, Fax1, vbCrLf)
        a$ = String.Format("{0}Fax2: {1}{2}", a$, Fax2, vbCrLf)
        a$ = String.Format("{0}MobilFon: {1}{2}", a$, MobilFon, vbCrLf)
        a$ = String.Format("{0}Email: {1}{2}", a$, Email, vbCrLf)
        a$ = String.Format("{0}Homepage: {1}{2}", a$, Homepage, vbCrLf)
        Return a$
    End Function

    Public Sub New()
        clear()
    End Sub
End Class
