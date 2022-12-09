Module DatabaseQueries

    Public Sub CheckSerialNumber(serialNum As String, ByRef drawingNum As String)
        Try
            Dim DalCount As New DALControl
            Dim queryCount As String = "SELECT COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_RECENTORDER WHERE SERIALNUM = '" & serialNum & "'"
            DalCount.RunQuery(queryCount)
            If DalCount.SQLDataset01.Tables(0).Rows(0).Item(0) > 0 Then
                Dim dalDN As New DALControl
                Dim queryDN As String = "SELECT DRAWINGNUM FROM SFM_RCH.Y_PANEL_BUILD_RECENTORDER"
                dalDN.RunQuery(queryDN)
                drawingNum = dalDN.SQLDataset01.Tables(0).Rows(0).Item(0)
                Dim dalComplete As New DALControl
                Dim queryComplete As String = "SELECT COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' AND COMPLETED = 0 AND PART_NUMBER <> '72480649001:1'"
                dalComplete.RunQuery(queryComplete)
                If (dalComplete.SQLDataset01.Tables(0).Rows(0).Item(0) = 0) Then
                    Dim dalUpdate As New DALControl
                    Dim queryUpdate As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 0 WHERE DRAWINGNUM = '" & drawingNum & "'"
                    dalUpdate.RunQuery(queryUpdate)
                End If
            Else
                Dim dalSNpo As New DALControl
                Dim querySNpo As String = "SELECT COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_PASTORDERS WHERE SERIALNUM = '" & serialNum & "'"
                dalSNpo.RunQuery(querySNpo)
                If dalSNpo.SQLDataset01.Tables(0).Rows(0).Item(0) = 0 Then
                    MsgBox("That order has not been extracted yet")
                    Dim message, title
                    message = "What is the order's drawing number?"
                    title = "Extraction Check"
                    drawingNum = InputBox(message, title)
                    MainForm.txtDN.Text = drawingNum
                    CheckDrawingNumber(serialNum, drawingNum)
                Else
                    Dim dalDraw As New DALControl
                    Dim queryDraw As String = "SELECT DRAWINGNUM FROM Y_PANEL_BUILD_PASTORDERS WHERE SERIALNUM = '" & serialNum & "'"
                    dalDraw.RunQuery(queryDraw)
                    drawingNum = dalDraw.SQLDataset01.Tables(0).Rows(0).Item(0)
                    Dim dalUpdate As New DALControl
                    Dim queryUpdate As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 0 WHERE DRAWINGNUM = '" & drawingNum & "'"
                    dalUpdate.RunQuery(queryUpdate)
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Sub CheckDrawingNumber(serialNum As String, ByRef drawingNum As String)
        Try
            Dim dalDraw As New DALControl
            Dim queryDraw As String = "SELECT COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_PASTORDERS WHERE DRAWINGNUM = '" & drawingNum & "'"
            dalDraw.RunQuery(queryDraw)
            If dalDraw.SQLDataset01.Tables(0).Rows(0).Item(0) = 0 Then
                'That order has not been extracted yet
                ExtractInventorData(drawingNum)
                FuseAssemblyExtract(drawingNum)
            Else
                Dim dalDraw2 As New DALControl
                Dim queryDraw2 As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 0 WHERE DRAWINGNUM = '" & drawingNum & "'"
                dalDraw2.RunQuery(queryDraw2)
            End If
            Dim dalAdd As New DALControl
            Dim queryAdd As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_PASTORDERS (SERIALNUM, DRAWINGNUM) VALUES ('" & serialNum & "', '" & drawingNum & "')"
            dalAdd.RunQuery(queryAdd)
        Catch ex As Exception

        End Try
    End Sub

    Public Function GetRecentOrder()
        Dim dal As New DALControl
        Dim query As String = "SELECT SERIALNUM, DRAWINGNUM FROM SFM_RCH.Y_PANEL_BUILD_RECENTORDER"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Sub WriteActiveOrder(sSN As String, drawingNum As String)
        Dim dal4 As New DALControl
        Dim Query4 = "UPDATE SFM_RCH.Y_PANEL_BUILD_RECENTORDER SET SERIALNUM = '" & sSN & "'"
        dal4.RunQuery(Query4)
        Dim dal5 As New DALControl
        Dim query5 = "UPDATE SFM_RCH.Y_PANEL_BUILD_RECENTORDER SET DRAWINGNUM = '" & drawingNum & "'"
        dal5.RunQuery(query5)
        Dim dal6 As New DALControl
        Dim query6 = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET DRAWINGNUM = '" & drawingNum & "' WHERE PART_NUMBER = '72480649001:1'"
        dal6.RunQuery(query6)
        Dim dal7 As New DALControl
        Dim query7 = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 0 WHERE PART_NUMBER = '72480649001:1'"
        dal7.RunQuery(query7)
    End Sub

    Public Function GetPanel(drawingNum As String) As String
        'Selects the panel from Y_PANEL_POSITIONS then marks it as complete
        Try
            Dim dalSelect As New DALControl
            Dim SelectPanel As String = "SELECT PART_NUMBER FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' AND X1 = 0 AND Y1 = 0 AND Z1 = 0 AND (COMPLETED = 0 OR COMPLETED IS NULL)"
            dalSelect.RunQuery(SelectPanel)
            Dim sPanel As String = dalSelect.SQLDataset01.Tables(0).Rows(0).Item(0)
            Return sPanel
        Catch ex As Exception
            MsgBox("Error: Drawing Number entered was not found.")
            Application.Restart()
            Return Nothing
        End Try
    End Function

    Public Function CheckCount(drawingNum As String) As Boolean

        Dim dalLoc As New DALControl
        Dim LocQuery = "SELECT COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' AND COMPLETED = 0"

        dalLoc.RunQuery(LocQuery)
        Dim partCount As Integer = CInt(dalLoc.SQLDataset01.Tables(0).Rows(0).Item(0))
        MainForm.txtLog.Text &= "Number of parts remaining: " & partCount & Environment.NewLine
        If partCount > 1 Then '0 Then  'temporary. Change when panel completion gets updated
            CheckCount = True
        Else
            CheckCount = False
        End If
        Return CheckCount
    End Function

    Public Function GetFirstItem(drawingNum As String, sPanel As String) As String
        Dim dalFirst As New DALControl
        Dim FirstQuery = "SELECT PART_NUMBER FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' AND PART_NUMBER <> '" & sPanel & "' AND (COMPLETED = 0 OR COMPLETED IS NULL) ORDER BY X1, Y1"
        dalFirst.RunQuery(FirstQuery)
        Dim sPN As String
        Try
            sPN = dalFirst.SQLDataset01.Tables(0).Rows(0).Item(0)
        Catch
            sPN = Nothing
        End Try
        Return sPN

    End Function

    Public Function GetLocation(drawingNum As String, sPartNum As String) As DataRow

        Dim dalLoc As New DALControl
        Dim LocQuery = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM LIKE '%" & drawingNum & "%' AND PART_NUMBER LIKE '%" & sPartNum & "%' AND Z1 = '0.1196' AND (COMPLETED = 0 OR COMPLETED IS NULL)"
        dalLoc.RunQuery(LocQuery)
        GetLocation = dalLoc.SQLDataset01.Tables(0).Rows(0)

    End Function

    Public Sub SetIncomplete(drawingNum As String, sPartNum As String)
        Dim dal As New DALControl
        Dim query As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 0 WHERE DRAWINGNUM LIKE '%" & drawingNum & "%' AND PART_NUMBER = '" & sPartNum & "'"
        dal.RunQuery(query)
    End Sub

    Public Function GetBinCoordinates(sPartNumber As String) As DataRow
        Dim subString As String
        subString = Left(sPartNumber, 8)
        Dim dal As New DALControl
        Dim Query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_STORAGE_LOCATION WHERE UNIT = 'TEST' AND PART_NUMBER LIKE '%" & subString & "%' AND FILL != 0"
        dal.RunQuery(Query)
        Dim returnRow As DataRow
        Try
            returnRow = dal.SQLDataset01.Tables(0).Rows(0)
        Catch ex As Exception
            returnRow = Nothing
        End Try
        Return returnRow
    End Function

    Public Function GetOrigin(Name As String) As DataRow

        Dim dalLoc As New DALControl
        Dim LocQuery = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_ORIGINS WHERE NAME LIKE '%" & Name & "%'"
        dalLoc.RunQuery(LocQuery)
        GetOrigin = dalLoc.SQLDataset01.Tables(0).Rows(0)

    End Function

    Public Function GetPartSize(splitPN As String) As DataRow
        Dim subString As String
        subString = Left(splitPN, 8)
        Dim LocQuery As String
        Dim dalLoc As New DALControl
        LocQuery = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_PARTSIZE WHERE PART_NUMBER LIKE '%" & subString & "%'"
        dalLoc.RunQuery(LocQuery)
        GetPartSize = dalLoc.SQLDataset01.Tables(0).Rows(0)
    End Function

    Public Function GetPanelHole(sPanel As String) As DataRow

        'Gets the data row from sql that has the panel hole data
        Dim dalLoc As New DALControl
        Dim LocQuery = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_PARTSIZE WHERE PART_NUMBER LIKE '%" & sPanel & "%' AND SCREW_TYPE = 'PANEL'"
        dalLoc.RunQuery(LocQuery)
        Dim bPanel As Boolean = False
        While bPanel = False
            Try
                GetPanelHole = dalLoc.SQLDataset01.Tables(0).Rows(0)
                bPanel = True
            Catch ex As Exception
                MsgBox("Panel information not in the Database. Exit out of application and input information.")
                GetPanelHole = Nothing
            End Try
        End While
        Return GetPanelHole
    End Function

    Public Sub SendActivePart(sPartNumber As String)
        Dim subString As String
        subString = Left(sPartNumber, 6)
        Dim dalLoc2 As New DALControl
        Dim LocQuery = "UPDATE SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS SET ACTIVE = 'X' WHERE PART_NUMBER LIKE '%" & subString & "%'"
        dalLoc2.RunQuery(LocQuery)
    End Sub

    Public Sub ActivePartComplete()
        Dim dalLoc As New DALControl
        Dim LocQuery = "UPDATE SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS SET ACTIVE = NULL WHERE ACTIVE IS NOT NULL"
        dalLoc.RunQuery(LocQuery)
    End Sub

    Public Function GetDepthOffset()
        Dim dal As New DALControl
        Dim query As String = "SELECT DEPTH, OFFSET FROM SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS WHERE ACTIVE IS NOT NULL"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Sub SendComplete(PN As String, drawingNum As String)
        Dim dalLoc As New DALControl
        Dim LocQuery = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 1 WHERE PART_NUMBER = '" & PN & "' AND DRAWINGNUM = '" & drawingNum & "'"
        dalLoc.RunQuery(LocQuery)
    End Sub

    Public Function GetBuildForUI(drawingNum As String) As DataTable
        Dim dal As New DALControl
        Dim query As String = "SELECT PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, COMPLETED FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' ORDER BY X1, Y1"
        dal.RunQuery(query)

        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            GetBuildForUI = dal.SQLDataset01.Tables(0)
        Else
            GetBuildForUI = Nothing
        End If
    End Function
    Public Sub CabQueue()
        Dim DALControl4 As New DALControl
        DALControl4.RunQuery("SELECT PART_NUMBER,COMPLETED FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE COMPLETED = 0 OR COMPLETED IS NULL")
        MainForm.panelQueueDt.DataSource = DALControl4.SQLDataset01.Tables(0)
        MainForm.panelQueueDt.Update()
        MainForm.panelQueueDt.AutoResizeColumns()
    End Sub

    Public Sub CheckStorage()
        Dim dal As New DALControl
        Dim query = "SELECT FILL, PART_NUMBER, COMMON_NAME FROM SFM_RCH.Y_PANEL_BUILD_STORAGE_LOCATION ORDER BY SLOT"
        dal.RunQuery(query)

        'Fill
        MainForm.TextBox1.Text = dal.SQLDataset01.Tables(0).Rows(7).Item(0)
        MainForm.TextBox6.Text = dal.SQLDataset01.Tables(0).Rows(8).Item(0)
        MainForm.TextBox9.Text = dal.SQLDataset01.Tables(0).Rows(9).Item(0)
        MainForm.TextBox12.Text = dal.SQLDataset01.Tables(0).Rows(10).Item(0)
        MainForm.TextBox15.Text = dal.SQLDataset01.Tables(0).Rows(11).Item(0)
        MainForm.TextBox18.Text = dal.SQLDataset01.Tables(0).Rows(12).Item(0)
        MainForm.TextBox21.Text = dal.SQLDataset01.Tables(0).Rows(13).Item(0)
        MainForm.TextBox42.Text = dal.SQLDataset01.Tables(0).Rows(14).Item(0)
        MainForm.TextBox39.Text = dal.SQLDataset01.Tables(0).Rows(15).Item(0)
        MainForm.TextBox36.Text = dal.SQLDataset01.Tables(0).Rows(16).Item(0)
        MainForm.TextBox33.Text = dal.SQLDataset01.Tables(0).Rows(17).Item(0)
        MainForm.TextBox30.Text = dal.SQLDataset01.Tables(0).Rows(18).Item(0)
        MainForm.TextBox27.Text = dal.SQLDataset01.Tables(0).Rows(19).Item(0)
        MainForm.TextBox24.Text = dal.SQLDataset01.Tables(0).Rows(20).Item(0)

        'Part Number
        MainForm.TextBox2.Text = dal.SQLDataset01.Tables(0).Rows(7).Item(1)
        MainForm.TextBox5.Text = dal.SQLDataset01.Tables(0).Rows(8).Item(1)
        MainForm.TextBox8.Text = dal.SQLDataset01.Tables(0).Rows(9).Item(1)
        MainForm.TextBox11.Text = dal.SQLDataset01.Tables(0).Rows(10).Item(1)
        MainForm.TextBox14.Text = dal.SQLDataset01.Tables(0).Rows(11).Item(1)
        MainForm.TextBox17.Text = dal.SQLDataset01.Tables(0).Rows(12).Item(1)
        MainForm.TextBox20.Text = dal.SQLDataset01.Tables(0).Rows(13).Item(1)
        MainForm.TextBox41.Text = dal.SQLDataset01.Tables(0).Rows(14).Item(1)
        MainForm.TextBox38.Text = dal.SQLDataset01.Tables(0).Rows(15).Item(1)
        MainForm.TextBox35.Text = dal.SQLDataset01.Tables(0).Rows(16).Item(1)
        MainForm.TextBox32.Text = dal.SQLDataset01.Tables(0).Rows(17).Item(1)
        MainForm.TextBox29.Text = dal.SQLDataset01.Tables(0).Rows(18).Item(1)
        MainForm.TextBox26.Text = dal.SQLDataset01.Tables(0).Rows(19).Item(1)
        MainForm.TextBox23.Text = dal.SQLDataset01.Tables(0).Rows(20).Item(1)

        'Common Name
        MainForm.TextBox3.Text = dal.SQLDataset01.Tables(0).Rows(7).Item(2)
        MainForm.TextBox4.Text = dal.SQLDataset01.Tables(0).Rows(8).Item(2)
        MainForm.TextBox7.Text = dal.SQLDataset01.Tables(0).Rows(9).Item(2)
        MainForm.TextBox10.Text = dal.SQLDataset01.Tables(0).Rows(10).Item(2)
        MainForm.TextBox13.Text = dal.SQLDataset01.Tables(0).Rows(11).Item(2)
        MainForm.TextBox16.Text = dal.SQLDataset01.Tables(0).Rows(12).Item(2)
        MainForm.TextBox19.Text = dal.SQLDataset01.Tables(0).Rows(13).Item(2)
        MainForm.TextBox40.Text = dal.SQLDataset01.Tables(0).Rows(14).Item(2)
        MainForm.TextBox37.Text = dal.SQLDataset01.Tables(0).Rows(15).Item(2)
        MainForm.TextBox34.Text = dal.SQLDataset01.Tables(0).Rows(16).Item(2)
        MainForm.TextBox31.Text = dal.SQLDataset01.Tables(0).Rows(17).Item(2)
        MainForm.TextBox28.Text = dal.SQLDataset01.Tables(0).Rows(18).Item(2)
        MainForm.TextBox25.Text = dal.SQLDataset01.Tables(0).Rows(19).Item(2)
        MainForm.TextBox22.Text = dal.SQLDataset01.Tables(0).Rows(20).Item(2)
    End Sub

    Public Sub ChangeStorage(slot As Integer, change As Integer)
        Dim dal As New DALControl
        Dim query = "SELECT FILL FROM SFM_RCH.Y_PANEL_BUILD_STORAGE_LOCATION WHERE SLOT = '" & slot & "'"
        dal.RunQuery(query)
        Dim newValue = CInt(dal.SQLDataset01.Tables(0).Rows(0).Item(0)) + change
        Dim dal2 As New DALControl
        Dim query2 = "UPDATE SFM_RCH.Y_PANEL_BUILD_STORAGE_LOCATION SET FILL = '" & newValue & "' WHERE SLOT = '" & slot & "'"
        dal2.RunQuery(query2)
    End Sub

    Public Sub FuseAssemblyExtract(drawingNum As String)
        Dim DAL As New DALControl
        Dim i As Integer = 1
        Try
            'CREATES OUR COUNTER FOR OUR LOOP
            Dim countQuery As String = "Select COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER Like '%72180560819%' and COMPLETED = 0"
            DAL.RunQuery(countQuery)
            Dim counter As Integer = CInt(DAL.SQLDataset01.Tables(0).Rows(0).Item(0))

            'SO IT DOESNT THROW AN ERROR
            If counter > 0 Then
                'SELECTS SERIAL NUMBER FOR INSERT STATEMENT
                Dim selectQuery As String = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0"
                DAL.RunQuery(selectQuery)
                Dim serialNum = DAL.SQLDataset01.Tables(0).Rows(0).Item(0)

                'LOOPS AND REPLACES ALL ASSEMBLY PARTS WITH SEPERATE PARTS
                While counter >= i

                    'X1 VALUE FOR THE PART
                    Dim selectX1 As String = "SELECT X1 FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0"
                    DAL.RunQuery(selectX1)
                    Dim X1 As Double = DAL.SQLDataset01.Tables(0).Rows(0).Item(0)

                    'Y1 VALUE FOR THE PART
                    Dim selectY1 As String = "SELECT Y1 FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0"
                    DAL.RunQuery(selectY1)
                    Dim Y1 As Double = DAL.SQLDataset01.Tables(0).Rows(0).Item(0)

                    'ROUNDED VARS
                    Dim roundedX1 = Math.Round(X1, 1)
                    Dim roundedY1 = Math.Round(Y1, 1)

                    'CHECKS THE LOCATION OF THE PART IN X AND Y DIRECTION AND INSERTS THE TWO PARTS IN THAT PLACE
                    If roundedX1 = 12.6 Then
                        If roundedY1 = 24.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:11', 12.4925, 24.3375, 0.1196, 15.7575 , 26.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:11', 15.875, 24.2935, 0.1196, 19.375, 26.3435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 19.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:12', 12.4925, 19.8375, 0.1196, 15.7575 , 21.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:12', 15.875, 19.7935, 0.1196, 19.375, 21.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 15.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:13', 12.4925, 15.3375, 0.1196, 15.7575 , 17.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:13', 15.875, 15.2935, 0.1196, 19.375, 17.3475, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 10.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:14', 12.4925, 10.8375, 0.1196,15.7575 , 12.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:14', 15.875, 10.7935, 0.1196, 19.375, 12.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 6.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:15', 12.4925, 6.3375, 0.1196,15.7575 , 8.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:15', 15.875, 6.2935, 0.1196, 19.375, 8.3435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 1.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', '871148032:16', 12.4925, 1.8375, 0.1196,15.7575 , 3.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & drawingNum & "', 'W 666443:16', 15.875, 1.7935, 0.1196, 19.375, 3.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        End If
                    End If

                    'UPDATES THE ASSEMBLY TO COMPLETE
                    Dim updatePart As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 1 WHERE PART_NUMBER LIKE '%72180560819%' AND X1 = " & X1.ToString & " AND Y1 = " & Y1.ToString & ""
                    DAL.RunQuery(updatePart)
                    i += 1
                End While
                'Dim DalDelete As New DALControl
                'Dim deletePart As String = "DELETE FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & drawingNum & "' AND PART_NUMBER LIKE '%72180560819%'"
                'DalDelete.RunQuery(deletePart)
            End If
        Catch ex As Exception
            MsgBox("Error: FuseAssemblyExtract() failed to extract..")
        End Try
    End Sub
End Module
