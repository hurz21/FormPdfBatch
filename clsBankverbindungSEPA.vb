Imports System.ComponentModel
Public Class clsBankverbindungSEPA
    Implements INotifyPropertyChanged
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
                             Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(ByVal prop As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Property personenID As Integer
    Property BANKKONTOID As Integer
    'Property istVORLAGE1 As boolean = false


    Property tss As Date
    Property Quelle As String


    Private _istVORLAGE1 As Boolean
    Public Property istVORLAGE1() As Boolean
        Get
            Return _istVORLAGE1
        End Get
        Set(ByVal value As Boolean)
            _istVORLAGE1 = value
            OnPropertyChanged("istVORLAGE1")
        End Set
    End Property


    Private _bankname As String
    Public Property BankName() As String
        Get
            Return _bankname
        End Get
        Set(ByVal Value As String)
            _bankname = Value
            OnPropertyChanged("BankName")
        End Set
    End Property

    Private _BIC As String
    Public Property BIC() As String
        Get
            Return _BIC
        End Get
        Set(ByVal Value As String)
            _BIC = Value
            OnPropertyChanged("BIC")
        End Set
    End Property
    Private _IBAN As String
    Public Property IBAN() As String
        Get
            Return _IBAN
        End Get
        Set(ByVal Value As String)
            _IBAN = Value
            OnPropertyChanged("IBAN")
        End Set
    End Property
    Private _titel As String
    Public Property Titel() As String
        Get
            Return _titel
        End Get
        Set(ByVal Value As String)
            _titel = Value
            OnPropertyChanged("Titel")
        End Set
    End Property

    Sub clear()
        Titel = ""
        IBAN = ""
        BIC = ""
        BankName = ""

        '  personenID=0
        BANKKONTOID = 0
    End Sub
End Class

