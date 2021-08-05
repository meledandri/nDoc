Imports System.IO

Public Class local_db
    Dim _local_file As String = ""
    Dim _auto_save As Boolean = False
    Dim _file_present As Boolean = False

    Dim _data As New List(Of _record)
    Public ReadOnly Property data As List(Of _record)
        Get
            Return _data
        End Get
    End Property

    Public Sub add(ByVal key As String, value As String)
        Dim found As Boolean = False
        Dim found_rec As _record = find(key)
        If Not IsNothing(found_rec) Then
            _data.Remove(found_rec)
        End If
        _data.Add(New _record(key, value))
    End Sub

    Private Function find(Key As String) As _record
        Dim found_rec As _record
        For Each r As _record In _data
            If r.Key = Key Then
                found_rec = r
                Exit For
            End If
        Next
        Return found_rec
    End Function

    Public Sub remove(key As String)
        Dim found_rec As _record = find(key)
        If Not IsNothing(found_rec) Then _data.Remove(found_rec)
    End Sub

    Public Function get_value(key As String) As String
        Dim r As String = Nothing
        Dim found_rec As _record = find(key)
        If Not IsNothing(found_rec) Then r = found_rec.Value
        Return r
    End Function

    Public ReadOnly Property Exist As Boolean
        Get
            Return _file_present
        End Get
    End Property

    Public Sub New(ByVal local_file As String, Optional auto_save As Boolean = False)
        _auto_save = auto_save
        _local_file = local_file
        Try
            If File.Exists(local_file) Then
                load()
                _file_present = True
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub save()
        Dim F As Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim s As IO.Stream
        F = New Runtime.Serialization.Formatters.Binary.BinaryFormatter()
        s = New IO.FileStream(_local_file, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None)
        F.Serialize(s, data)
        s.Close()

    End Sub

    Public Sub load()
        Dim f As Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim s As IO.Stream
        f = New Runtime.Serialization.Formatters.Binary.BinaryFormatter()
        s = New IO.FileStream(_local_file, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
        _data = DirectCast(f.Deserialize(s), Object)
        s.Close()

    End Sub


    <Serializable()> Public Class _record
        Dim _key As String = ""
        Public ReadOnly Property Key As String
            Get
                Return _key
            End Get
        End Property

        Dim _value As String = Nothing
        Public ReadOnly Property Value As String
            Get
                Return _value
            End Get
        End Property

        Public Sub New(ByVal key As String, Optional value As String = Nothing)
            _key = key
            _value = value
        End Sub


    End Class

End Class
