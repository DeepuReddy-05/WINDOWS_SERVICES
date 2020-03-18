Imports System.Data.OleDb
Imports System.IO
Imports System.Configuration
Imports System.Threading
Imports System.Net
Imports System.Xml
Imports Newtonsoft.Json
Imports System.Text

Public Class Shopify_Inventory_Update

    Dim dbad As New OleDb.OleDbDataAdapter
    Dim cmd As OleDbCommand
    Dim dbcon As OleDbConnection

    Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
    Protected Shared strLogFilePath As String = AppPath + "\" + "OmegaCube_Shopify_Inventory_Update_Log.txt"
    Private Shared sw As StreamWriter = Nothing
    Dim scheduleTimer As New System.Timers.Timer()
    Public eCheckFreq As Integer
    Public strPreserveLog As String = String.Empty
    Dim nSchemas As Integer = 0

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        Try
            eCheckFreq = Convert.ToInt16(ConfigurationManager.AppSettings("CheckFrequency").ToString())

            strPreserveLog = ConfigurationManager.AppSettings("PreserveLog").ToString()
            If String.IsNullOrEmpty(strPreserveLog) Then
                strPreserveLog = "N"
            End If

            If eCheckFreq = 0 Then
                eCheckFreq = 1
            End If

            LogExceptions(strLogFilePath, Nothing, "Service Started")
            scheduleTimer.Interval = eCheckFreq * 60 * 1000

            AddHandler scheduleTimer.Elapsed, New Timers.ElapsedEventHandler(AddressOf TimerElapsed)
            scheduleTimer.Enabled = True
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub

    Private Sub TimerElapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        Try
            scheduleTimer.Enabled = False
            LogExceptions(strLogFilePath, Nothing, "Started Executed scripts")
            Call Inventory_Update()
            LogExceptions(strLogFilePath, Nothing, "Started Executed scripts END")
            scheduleTimer.Enabled = True
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub


    Public Sub Inventory_Update()
        Try
            nSchemas = Convert.ToInt16(ConfigurationManager.AppSettings("NoOfSchemas").ToString())

            If nSchemas = 0 Then
                nSchemas = 10
            End If

            For i As Integer = 1 To nSchemas

                Dim strServerName As String = String.Empty
                Dim strUserName As String = String.Empty
                Dim strPassWord As String = String.Empty

                strServerName = ConfigurationManager.AppSettings("Server" & i)
                strUserName = ConfigurationManager.AppSettings("User" & i)
                strPassWord = ConfigurationManager.AppSettings("Pwd" & i)

                Try
                    If Not String.IsNullOrEmpty(strServerName) Then
                        Try
                            If Not String.IsNullOrEmpty(strUserName) Then
                                Try
                                    If Not String.IsNullOrEmpty(strPassWord) Then
                                        Try

                                            'Shopify_Inventory_Update
                                            'Start
                                            dbcon = New OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + ConfigurationManager.AppSettings("Server") + ";User Id=" + ConfigurationManager.AppSettings("User") + ";Password=" + ConfigurationManager.AppSettings("Pwd") + ";")
                                            dbad.SelectCommand = New OleDbCommand
                                            dbad.SelectCommand.Connection = dbcon
                                            Dim dsn As New Data.DataSet
                                            dbad.SelectCommand.CommandText = "SELECT ITEM_NO,AVAILABLE,VARIANT_ID FROM IS_INVENTORY WHERE UPDATED_TO_WEB IS NULL AND VARIANT_ID IS NOT NULL and rownum<50"
                                            dsn.Clear()
                                            If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Closed) Then
                                                dbad.SelectCommand.Connection.Open()
                                            End If
                                            dbad.Fill(dsn)
                                            If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Open) Then
                                                dbad.SelectCommand.Connection.Close()
                                            End If
                                            If (dsn.Tables(0).Rows.Count > 0) Then
                                                Try

                                                    dbcon = New OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + strServerName + ";User Id=" + strUserName + ";Password=" + strPassWord + ";")
                                                    dbad.SelectCommand = New OleDbCommand
                                                    dbad.SelectCommand.Connection = dbcon
                                                    Dim dst As New DataSet
                                                    dbad.SelectCommand.CommandText = "SELECT VALUE FROM SYS_TOOL_SETTINGS WHERE KEY='SHOPIFY_INVENTORY_CREDENTIALS'"
                                                    dst.Clear()
                                                    If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Closed) Then
                                                        dbad.SelectCommand.Connection.Open()
                                                    End If
                                                    dbad.Fill(dst)
                                                    If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Open) Then
                                                        dbad.SelectCommand.Connection.Close()
                                                    End If
                                                    If dst.Tables(0).Rows.Count > 0 Then

                                                        Dim TimeVal As String = dst.Tables(0).Rows(0)("VALUE").ToString()
                                                        Dim api_credentials As String() = TimeVal.Split(New Char() {":"c})
                                                        Dim username, passowrd As String
                                                        username = api_credentials(0)
                                                        passowrd = api_credentials(1)

                                                        Dim shopity_url As String = "https://froedge-machine-and-supply-company.myshopify.com"

                                                        Try
                                                            Dim dst1 As New DataSet
                                                            dbad.SelectCommand.CommandText = "SELECT VALUE FROM SYS_TOOL_SETTINGS WHERE KEY='SHOPIFY_CAN_URL'"
                                                            dst1.Clear()
                                                            If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Closed) Then
                                                                dbad.SelectCommand.Connection.Open()
                                                            End If
                                                            dbad.Fill(dst1)
                                                            If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Open) Then
                                                                dbad.SelectCommand.Connection.Close()
                                                            End If
                                                            If dst1.Tables(0).Rows.Count > 0 Then
                                                                shopity_url = dst1.Tables(0).Rows(0)("VALUE").ToString()
                                                            End If
                                                        Catch ex As Exception

                                                        End Try
                                                        For Each dr As DataRow In dsn.Tables(0).Rows
                                                            Try


                                                                Dim VARIANT_ID As String

                                                                If Not (Equals(dr("VARIANT_ID"), System.DBNull.Value)) Then
                                                                    VARIANT_ID = dr("VARIANT_ID")
                                                                Else
                                                                    VARIANT_ID = ""
                                                                End If


                                                                If VARIANT_ID <> "" Then

                                                                    Dim id As String = VARIANT_ID

                                                                    Dim request1 As WebRequest = WebRequest.Create(shopity_url & "/admin/api/2020-01/inventory_levels/set.json")
                                                                    request1.Credentials = New NetworkCredential(username, passowrd)


                                                                    If File.Exists(AppPath + "\" + "Shopify_Inventory_Update.json") Then
                                                                        File.Delete(AppPath + "\" + "Shopify_Inventory_Update.json")
                                                                    End If

                                                                    Dim filePath = AppPath + "\" + "Shopify_Inventory_Update.json"
                                                                    Dim sw As StreamWriter = Nothing
                                                                    sw = New StreamWriter(filePath, True)
                                                                    sw.WriteLine("{")
                                                                    ' sw.WriteLine("""variant"": {")
                                                                    sw.WriteLine("""location_id"":39055692,")
                                                                    sw.WriteLine("""inventory_item_id"":" & id & ",")
                                                                    sw.WriteLine("""available"":" & dr("AVAILABLE"))
                                                                    '  sw.WriteLine("}")
                                                                    sw.WriteLine("}")
                                                                    sw.Flush()
                                                                    sw.Close()


                                                                    Dim sb As New StringBuilder()
                                                                    Using sr As New StreamReader(filePath)
                                                                        Do While sr.Peek() > -1
                                                                            sb.AppendLine(sr.ReadLine())
                                                                        Loop

                                                                    End Using

                                                                    request1.Method = "POST"

                                                                    request1.ContentType = "application/json"


                                                                    Dim lbPostBuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(sb.ToString())
                                                                    request1.ContentLength = lbPostBuffer.Length
                                                                    Dim loPostData As Stream = request1.GetRequestStream()
                                                                    loPostData.Write(lbPostBuffer, 0, lbPostBuffer.Length)
                                                                    loPostData.Close()

                                                                    Dim response1 As HttpWebResponse = Nothing
                                                                    Try
                                                                        response1 = DirectCast(request1.GetResponse(), HttpWebResponse)
                                                                    Catch ex As Exception
                                                                        LogExceptions(strLogFilePath, ex, Nothing)
                                                                        Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                                                                        Exit Sub
                                                                    End Try

                                                                    Dim dataStream1 As Stream = response1.GetResponseStream()
                                                                    Dim reader1 As New StreamReader(dataStream1)
                                                                    Dim responseFromServer1 As String = reader1.ReadToEnd()


                                                                    If response1.StatusCode = 200 And response1.StatusDescription = "OK" Then
                                                                        dbcon = New OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + ConfigurationManager.AppSettings("Server") + ";User Id=" + ConfigurationManager.AppSettings("User") + ";Password=" + ConfigurationManager.AppSettings("Pwd") + ";")
                                                                        dbad.UpdateCommand = New OleDbCommand
                                                                        dbad.UpdateCommand.Connection = dbcon
                                                                        If (dbad.UpdateCommand.Connection.State = Data.ConnectionState.Closed) Then
                                                                            dbad.UpdateCommand.Connection.Open()
                                                                        End If
                                                                        dbad.UpdateCommand.CommandText = "UPDATE IS_INVENTORY SET UPDATED_TO_WEB=SYSDATE WHERE ITEM_NO='" & dr("ITEM_NO") & "'"
                                                                        dbad.UpdateCommand.ExecuteNonQuery()
                                                                        dbad.UpdateCommand.Connection.Close()
                                                                    Else LogExceptions(strLogFilePath,, response1.StatusCode & "-" & response1.StatusDescription)
                                                                        Insert_Sys_Log(response1.StatusCode & "-" & response1.StatusDescription, strServerName, strUserName, strPassWord)
                                                                    End If


                                                                End If
                                                                Thread.Sleep(5000)
                                                            Catch ex As Exception
                                                                Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                                                            End Try
                                                        Next
                                                    End If
                                                Catch ex As Exception
                                                    LogExceptions(strLogFilePath, ex, Nothing)
                                                    Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                                                End Try
                                            End If
                                            'End


                                        Catch ex As Exception
                                            LogExceptions(strLogFilePath, ex, Nothing)
                                            Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                                        End Try
                                    End If
                                Catch ex As Exception
                                    LogExceptions(strLogFilePath, ex, Nothing)
                                    Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                                End Try
                            End If
                        Catch ex As Exception
                            LogExceptions(strLogFilePath, ex, Nothing)
                            Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                        End Try
                    End If
                Catch ex As Exception
                    LogExceptions(strLogFilePath, ex, Nothing)
                    Insert_Sys_Log(ex.ToString(), strServerName, strUserName, strPassWord)
                End Try
        Next
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, Nothing)
        End Try
    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        Try
            Thread.Sleep(1000)
            scheduleTimer.Enabled = False
            LogExceptions(strLogFilePath, Nothing, "Service Stopped")
        Catch ex As Exception

        End Try
    End Sub

    Public Sub LogExceptions(ByVal filePath As String, Optional ByVal ex As Exception = Nothing, Optional ByVal msg As String = Nothing)
        If File.Exists(filePath) Then
            If strPreserveLog = "N" Then
                File.Delete(filePath)
            End If
        End If
        If False = File.Exists(filePath) Then
            Dim fs As New FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            fs.Close()
        End If
        WriteExceptionLog(filePath, ex, msg)
    End Sub

    Private Sub WriteExceptionLog(ByVal strPathName As String, Optional ByVal objException As Exception = Nothing, Optional ByVal msg As String = Nothing)
        If (Not objException Is Nothing) AndAlso (Not String.IsNullOrEmpty(msg)) Then
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine("Source     :" & objException.Source.ToString().Trim())
            sw.WriteLine("Method     : " & objException.TargetSite.Name.ToString())
            sw.WriteLine("Date       : " & DateTime.Now.ToLongTimeString())
            sw.WriteLine("Time       : " & DateTime.Now.ToShortDateString())
            sw.WriteLine("Error      : " & objException.Message.ToString().Trim())
            sw.WriteLine("Stack Trace: " & objException.StackTrace.ToString().Trim())
            sw.WriteLine("^^-------------------------------------------------------------------^^")
            sw.WriteLine(msg)
            sw.WriteLine("^^-------------------------------------------------------------------^^")
        ElseIf String.IsNullOrEmpty(msg) Then
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine("Source     :" & objException.Source.ToString().Trim())
            sw.WriteLine("Method     : " & objException.TargetSite.Name.ToString())
            sw.WriteLine("Date       : " & DateTime.Now.ToLongTimeString())
            sw.WriteLine("Time       : " & DateTime.Now.ToShortDateString())
            sw.WriteLine("Error      : " & objException.Message.ToString().Trim())
            sw.WriteLine("Stack Trace: " & objException.StackTrace.ToString().Trim())
            sw.WriteLine("^^-------------------------------------------------------------------^^")
        Else
            sw = New StreamWriter(strPathName, True)
            sw.WriteLine(msg)
        End If
        sw.Flush()
        sw.Close()
    End Sub

    Public Sub Insert_Sys_Log(ByVal message As String, ByVal strServerName1 As String, ByVal strUserName1 As String, ByVal strPassWord1 As String)
        Try
            dbcon = New OleDbConnection("Provider=OraOLEDB.Oracle;Data Source=" + strServerName1 + ";User Id=" + strUserName1 + ";Password=" + strPassWord1 + ";")
            Dim sterr1, sterr2, sterr3, sterr4, sterr As String
            sterr = Replace(message, "'", "''")
            If (Len(sterr) > 4000) Then
                sterr1 = Mid(sterr, 1, 4000)
                If (Len(sterr) > 8000) Then
                    sterr2 = Mid(sterr, 4000, 8000)
                    If (Len(sterr) > 12000) Then
                        sterr3 = Mid(sterr, 8000, 12000)
                        If (Len(sterr) > 16000) Then
                            sterr4 = Mid(sterr, 12000, 16000)
                        Else
                            sterr4 = Mid(sterr, 12000, Len(sterr))
                        End If
                    Else
                        sterr3 = Mid(sterr, 8000, Len(sterr))
                        sterr4 = ""
                    End If
                Else
                    sterr2 = Mid(sterr, 4000, Len(sterr))
                    sterr3 = ""
                    sterr3 = ""
                    sterr4 = ""
                End If
            Else
                sterr1 = sterr
                sterr2 = ""
                sterr3 = ""
                sterr4 = ""
            End If
            dbad.InsertCommand = New OleDbCommand
            dbad.InsertCommand.Connection = dbcon
            dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select nvl(max(to_number(line_no)),0)+1 from SYS_ACTIVATE_STATUS_LOG),'EDGE','SERVICE','Shopify_Inventory_Update','" & sterr1 & "','N',sysdate,'" & sterr2 & "','" & sterr3 & "','" & sterr4 & "')"
            If (dbad.InsertCommand.Connection.State = ConnectionState.Closed) Then
                dbad.InsertCommand.Connection.Open()
            End If
            dbad.InsertCommand.ExecuteNonQuery()
            If (dbad.InsertCommand.Connection.State = ConnectionState.Open) Then
                dbad.InsertCommand.Connection.Close()
            End If
        Catch ex As Exception
            LogExceptions(strLogFilePath, ex, "Insert_Sys_Log")
        End Try
    End Sub

End Class
