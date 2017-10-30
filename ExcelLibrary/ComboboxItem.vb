Public Class ComboboxItem
    Public Property Text() As String
        Get
            Return m_Text
        End Get
        Set(ByVal value As String)
            m_Text = Value
        End Set
    End Property
    Private m_Text As String
    Public Property Value() As Object
        Get
            Return m_Value
        End Get
        Set(ByVal value As Object)
            m_Value = Value
        End Set
    End Property
    Private m_Value As Object

    Public Overrides Function ToString() As String
        Return Text
    End Function
End Class
