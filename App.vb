Imports System.Data.OracleClient

Public Class App
    Private _TableUser As String
    Private _Client As String
    Private _Name As String

    Private _Conn As OracleConnection
    Private _Cmd As OracleCommand

    Private _IsOpen As Boolean
    Private _UseMailbox As Boolean

    Private sSQL As String = String.Empty
    Private mOdbcReader As OracleDataReader

    Private mRawSql As MESRawSQL001.clsMESRawSQL
    Private mRawSqlReader As VBA.Collection

    Property ClientID() As String
        Get
            Return _Client
        End Get
        Set(ByVal value As String)
            _Client = value
        End Set
    End Property

    Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Public Sub New(ByVal aConnectObj As Object, ByVal aTableUser As String)
        Try
            _TableUser = aTableUser
            If TypeOf (aConnectObj) Is String Then
                If aConnectObj = String.Empty Then
                    Throw New ApplicationException("資料庫連線參數不可以空白")
                End If
                _Conn = New OracleConnection(aConnectObj)

                _Cmd = New OracleCommand
                _Cmd.Connection = _Conn
                _IsOpen = False
                _UseMailbox = False
            Else
                mRawSql = aConnectObj
                _IsOpen = True
                _UseMailbox = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    'Public Sub New(ByVal aConnectString As String)
    '    Try

    '        If aConnectString = String.Empty Then
    '            Throw New ApplicationException("資料庫連線參數不可以空白")
    '        End If
    '        _Conn = New OracleConnection(aConnectString)

    '        _Cmd = New OracleCommand
    '        _Cmd.Connection = _Conn
    '        _IsOpen = False
    '        _UseMailbox = False
    '    Catch ex As Exception
    '        Throw
    '    End Try

    'End Sub



    Public Sub Open()
        Try
            If _UseMailbox = True Then
                Exit Sub
            End If
            If _IsOpen = False Then
                _Conn.Open()
                _IsOpen = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub Close()
        Try
            If _UseMailbox = True Then
                Exit Sub
            End If
            If _IsOpen = True Then
                _Conn.Close()
                _IsOpen = False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 將參數值依照參數名稱寫入FA_APP_STATUS
    ''' </summary>
    ''' <param name="aAttribName">String，參數名稱</param>
    ''' <param name="aAttribValue">String，參數值</param>
    ''' <remarks></remarks>
    Public Sub Write(ByVal aAttribName As String, ByVal aAttribValue As String)
        Try
            If _IsOpen = False Then
                Throw New ApplicationException("資料庫尚未連線")
            End If
            If aAttribName = String.Empty Then
                Throw New ApplicationException("參數名稱不可以空白")
            End If
            If aAttribValue = String.Empty Then
                Throw New ApplicationException("參數值不可以空白")
            End If

            ' RawSQL 物件會將 \ 過濾，如果需要保留，就必須使用 \\
            If _UseMailbox = True Then
                aAttribValue = aAttribValue.Replace("\", "\\")
            End If

            sSQL = "  SELECT * FROM " & _TableUser & ".FA_APP_STATUS "
            sSQL &= " WHERE 1=1 "
            sSQL &= " AND CLIENT_ID='" & _Client & "' "
            sSQL &= " AND APP_NAME='" & _Name & "' "
            sSQL &= " AND ATTRIB_NAME='" & aAttribName & "' "

            Dim bHasRows As Boolean = False
            If _UseMailbox = True Then
                mRawSqlReader = mRawSql.QueryDatabase(sSQL)
                bHasRows = IIf(mRawSqlReader.Count <> 0, True, False)
            Else
                _Cmd.CommandText = sSQL
                mOdbcReader = _Cmd.ExecuteReader()
                bHasRows = mOdbcReader.HasRows
            End If

            If bHasRows Then
                sSQL = "  UPDATE " & _TableUser & ".FA_APP_STATUS SET ATTRIB_VALUE='" & aAttribValue & "' "
                sSQL &= " WHERE 1=1 "
                sSQL &= " AND CLIENT_ID='" & _Client & "' "
                sSQL &= " AND APP_NAME='" & _Name & "' "
                sSQL &= " AND ATTRIB_NAME='" & aAttribName & "' "
            Else
                sSQL = "  INSERT INTO " & _TableUser & ".FA_APP_STATUS ("
                sSQL &= " CLIENT_ID,APP_NAME,ATTRIB_NAME,ATTRIB_VALUE) VALUES ( "
                sSQL &= "'" & _Client & "',"
                sSQL &= "'" & _Name & "',"
                sSQL &= "'" & aAttribName & "',"
                sSQL &= "'" & aAttribValue & "')"
            End If

            If _UseMailbox = True Then
                mRawSql.QueryDatabase(sSQL)
            Else
                _Cmd.CommandText = sSQL
                _Cmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' 取得參數值
    ''' </summary>
    ''' <param name="aAttribName">String，參數名稱</param>
    ''' <returns>String，參數值</returns>
    ''' <remarks>如果參數名稱錯誤會傳回空白字串</remarks>
    Public Function Read(ByVal aAttribName As String) As String
        Dim sAttribValue As String = String.Empty
        Try
            If _IsOpen = False Then
                Throw New ApplicationException("資料庫尚未連線")
            End If
            If aAttribName = String.Empty Then
                Throw New ApplicationException("參數名稱不可以空白")
            End If

            sSQL = "  SELECT ATTRIB_VALUE FROM " & _TableUser & ".FA_APP_STATUS "
            sSQL &= " WHERE 1=1 "
            sSQL &= " AND CLIENT_ID='" & _Client & "' "
            sSQL &= " AND APP_NAME='" & _Name & "' "
            sSQL &= " AND ATTRIB_NAME='" & aAttribName & "' "

            _Cmd.CommandText = sSQL
            mOdbcReader = _Cmd.ExecuteReader

            While mOdbcReader.Read

                sAttribValue = mOdbcReader("ATTRIB_VALUE").ToString

                Exit While
            End While
        Catch ex As Exception
            Throw
        End Try

        Return sAttribValue
    End Function
End Class
