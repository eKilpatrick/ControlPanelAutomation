Imports System.Data.Sql
Imports System.Data.SqlClient
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

Public Class DALControl
    'Connection instructions for Oracle DB
    Dim SERVER_IP As String = "161.134.200.77"
    Dim DATABASE_NAME As String = "sfmjacks"
    Dim DATABASE_USER As String = "SFM_RCH"
    Dim DATABASE_PASSWORD As String = "SFMRichland$1"

    Dim oracleString As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & SERVER_IP & ")(PORT=1521))(CONNECT_DATA=(SERVER=DEDICATED)" &
                                       "(SERVICE_NAME=" & DATABASE_NAME & "))); USER ID = " & DATABASE_USER & ";Password = " & DATABASE_PASSWORD & ";"
    Public SQLConn As New Oracle.ManagedDataAccess.Client.OracleConnection(oracleString)
    Public SQLCommand01 As OracleCommand
    Public SQLDA As OracleDataAdapter
    Public SQLDataset01 As DataSet
    Public Function HasConnection() As Boolean
        Try
            SQLConn.Open()
            SQLConn.Close()
            Return True
        Catch ex As System.Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Sub RunQuery(Query As String)
        Try
            SQLConn.Open()
            SQLCommand01 = New OracleCommand(Query, SQLConn)
            SQLDA = New OracleDataAdapter(SQLCommand01)
            SQLDataset01 = New DataSet
            SQLDA.Fill(SQLDataset01)
            SQLConn.Close()
        Catch ex As System.Exception
            MsgBox("ERROR!" + vbNewLine + ex.Message + vbNewLine + " Query: " & Query, MsgBoxStyle.Critical)
            If SQLConn.State = ConnectionState.Open Then
                SQLConn.Close()
            End If


        End Try
    End Sub
    Public Function STR0_InsertRow(Str0_ParentID As Integer, Str0_Name1 As String, Str0_Name2 As String,
                         Str0_Type As String, Str0_Serie As String, Str0_Applies As Boolean,
                         Str0_Status As String, Str0_ToolTip As String, Str0_Notes As String, Str0_CreateBy As String,
                         Str0_LastUpdateBy As String, Str0_String1 As String, Str0_Description As String,
                         Str0_String2 As String, Str0_String3 As String, Str0_Int1 As Integer, Str0_Int2 As Integer,
                         Str0_Int3 As Integer, Str0_Float1 As Double, Str0_Float2 As Double, Str0_Float3 As Double) As Boolean

        Try
            SQLCommand01 = New OracleCommand("STR0_InsertRow", SQLConn)
            SQLCommand01.CommandType = CommandType.StoredProcedure
            SQLCommand01.Parameters.Add("@Str0_ParentID", SqlDbType.BigInt).Value = Str0_ParentID
            SQLCommand01.Parameters.Add("@Str0_Name1", SqlDbType.NVarChar).Value = Str0_Name1
            SQLCommand01.Parameters.Add("@Str0_Name2", SqlDbType.NVarChar).Value = Str0_Name2
            SQLCommand01.Parameters.Add("@Str0_Type", SqlDbType.NVarChar).Value = Str0_Type
            SQLCommand01.Parameters.Add("@Str0_Serie", SqlDbType.NChar).Value = Str0_Serie
            SQLCommand01.Parameters.Add("@Str0_Applies", SqlDbType.Bit).Value = Str0_Applies
            SQLCommand01.Parameters.Add("@Str0_Status", SqlDbType.NVarChar).Value = Str0_Status
            SQLCommand01.Parameters.Add("@Str0_ToolTip", SqlDbType.NVarChar).Value = Str0_ToolTip
            SQLCommand01.Parameters.Add("@Str0_Notes", SqlDbType.NVarChar).Value = Str0_Notes
            SQLCommand01.Parameters.Add("@Str0_CreateBy", SqlDbType.NVarChar).Value = Str0_CreateBy
            SQLCommand01.Parameters.Add("@Str0_LastUpdate", SqlDbType.DateTime).Value = DateTime.Now
            SQLCommand01.Parameters.Add("@Str0_LastUpdateBy", SqlDbType.NVarChar).Value = Str0_LastUpdateBy
            SQLCommand01.Parameters.Add("@Str0_String1", SqlDbType.NVarChar).Value = Str0_String1
            SQLCommand01.Parameters.Add("@Str0_CreateDate", SqlDbType.SmallDateTime).Value = DateTime.Now
            SQLCommand01.Parameters.Add("@Str0_Description", SqlDbType.NVarChar).Value = Str0_Description
            SQLCommand01.Parameters.Add("@Str0_String2", SqlDbType.NVarChar).Value = Str0_String2
            SQLCommand01.Parameters.Add("@Str0_String3", SqlDbType.NVarChar).Value = Str0_String3
            SQLCommand01.Parameters.Add("@Str0_Int1", SqlDbType.Int).Value = Str0_Int1
            SQLCommand01.Parameters.Add("@Str0_Int2", SqlDbType.Int).Value = Str0_Int2
            SQLCommand01.Parameters.Add("@Str0_Int3", SqlDbType.Int).Value = Str0_Int3
            SQLCommand01.Parameters.Add("@Str0_Float1", SqlDbType.Float).Value = Str0_Float1
            SQLCommand01.Parameters.Add("@Str0_Float2", SqlDbType.Float).Value = Str0_Float2
            SQLCommand01.Parameters.Add("@Str0_Float3", SqlDbType.Float).Value = Str0_Float3


            SQLConn.Open()
            SQLCommand01.ExecuteNonQuery()

            SQLConn.Close()
            Return True
        Catch ex As System.Exception
            MsgBox(ex.Message)
            Return (False)
        End Try





    End Function
    Public Function Ex0_InsertRow(Ext0_Name1 As String, Ext0_Name2 As String, Ext0_Value As String,
                                  Ext0_Description1 As String, Ext0_Description2 As String, Ext0_Notes1 As String,
                                  Ext0_Notes2 As String, Ext0_DataType As String, Ext0_OptionalValues As String,
                                  Ext0_CreateBy As String, Ext0_LastUpdateBy As String, Ext0_Applies As Boolean,
                                  Ext0_Status As Integer, Ext0_Serie As String, Ext0_ValueAux1 As String,
                                  Ext0_ValueAux2 As String, Ext0_User1 As String, Ext0_User2 As String,
                                  Ext0_String1 As String, Ext0_String2 As String, Ext0_String3 As String,
                                  Ext0_Int1 As Integer, Ext0_Int2 As Integer, Ext0_Int3 As Integer, Str0_ID As Integer,
                                  Ext0_Float1 As Double, Ext0_Float2 As Double, Ext0_Float3 As Double) As Boolean

        Try
            SQLCommand01 = New OracleCommand("Ex0_InsertRow", SQLConn)
            SQLCommand01.CommandType = CommandType.StoredProcedure
            SQLCommand01.Parameters.Add("@Ext0_Name1", SqlDbType.NVarChar).Value = Ext0_Name1
            SQLCommand01.Parameters.Add("@Ext0_Name2", SqlDbType.NVarChar).Value = Ext0_Name2
            SQLCommand01.Parameters.Add("@Ext0_Value", SqlDbType.NVarChar).Value = Ext0_Value
            SQLCommand01.Parameters.Add("@Ext0_Description1", SqlDbType.NVarChar).Value = Ext0_Description1
            SQLCommand01.Parameters.Add("@Ext0_Description2", SqlDbType.NVarChar).Value = Ext0_Description2
            SQLCommand01.Parameters.Add("@Ext0_Notes1", SqlDbType.NVarChar).Value = Ext0_Notes1
            SQLCommand01.Parameters.Add("@Ext0_Notes2", SqlDbType.NVarChar).Value = Ext0_Notes2
            SQLCommand01.Parameters.Add("@Ext0_DataType", SqlDbType.NVarChar).Value = Ext0_DataType
            SQLCommand01.Parameters.Add("@Ext0_OptionalValues", SqlDbType.NVarChar).Value = Ext0_OptionalValues
            SQLCommand01.Parameters.Add("@Ext0_CeateDate", SqlDbType.SmallDateTime).Value = DateTime.Now
            SQLCommand01.Parameters.Add("@Ext0_LastUpdate", SqlDbType.SmallDateTime).Value = DateTime.Now
            SQLCommand01.Parameters.Add("@Ext0_CreateBy", SqlDbType.NVarChar).Value = Ext0_CreateBy
            SQLCommand01.Parameters.Add("@Ext0_LastUpdateBy", SqlDbType.NVarChar).Value = Ext0_LastUpdateBy
            SQLCommand01.Parameters.Add("@Ext0_Applies", SqlDbType.Bit).Value = Ext0_Applies
            SQLCommand01.Parameters.Add("@Ext0_Status", SqlDbType.Int).Value = Ext0_Status
            SQLCommand01.Parameters.Add("@Ext0_Serie", SqlDbType.NVarChar).Value = Ext0_Serie
            SQLCommand01.Parameters.Add("@Ext0_Date1", SqlDbType.SmallDateTime).Value = DBNull.Value
            SQLCommand01.Parameters.Add("@Ext0_Date2", SqlDbType.SmallDateTime).Value = DBNull.Value
            SQLCommand01.Parameters.Add("@Ext0_ValueAux1", SqlDbType.NVarChar).Value = Ext0_ValueAux1
            SQLCommand01.Parameters.Add("@Ext0_ValueAux2", SqlDbType.NVarChar).Value = Ext0_ValueAux2
            SQLCommand01.Parameters.Add("@Ext0_User1", SqlDbType.NVarChar).Value = Ext0_User1
            SQLCommand01.Parameters.Add("@Ext0_User2", SqlDbType.NVarChar).Value = Ext0_User2
            SQLCommand01.Parameters.Add("@Ext0_String1", SqlDbType.NVarChar).Value = Ext0_String1
            SQLCommand01.Parameters.Add("@Ext0_String2", SqlDbType.NVarChar).Value = Ext0_String2
            SQLCommand01.Parameters.Add("@Ext0_String3", SqlDbType.NVarChar).Value = Ext0_String3
            SQLCommand01.Parameters.Add("@Ext0_Int1", SqlDbType.Int).Value = Ext0_Int1
            SQLCommand01.Parameters.Add("@Ext0_Int2", SqlDbType.Int).Value = Ext0_Int2
            SQLCommand01.Parameters.Add("@Ext0_Int3", SqlDbType.Int).Value = Ext0_Int3
            SQLCommand01.Parameters.Add("@Str0_ID", SqlDbType.Int).Value = Str0_ID
            SQLCommand01.Parameters.Add("@Ext0_Float1", SqlDbType.Float).Value = Ext0_Float1
            SQLCommand01.Parameters.Add("@Ext0_Float2", SqlDbType.Float).Value = Ext0_Float2
            SQLCommand01.Parameters.Add("@Ext0_Float3", SqlDbType.Float).Value = Ext0_Float3


            SQLConn.Open()
            SQLCommand01.ExecuteNonQuery()

            SQLConn.Close()
            Return True
        Catch ex As System.Exception
            MsgBox(ex.Message)
            Return (False)
        End Try

    End Function


End Class
