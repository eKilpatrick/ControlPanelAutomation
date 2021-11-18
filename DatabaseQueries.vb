Imports DAL1
Module DatabaseQueries
    Public Sub UpdatePartsQueue()
        Dim dal As New DALControl
        Dim query = "SELECT PART_NUMBER, COMPLETED FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE COMPLETED = 0 OR COMPLETED IS NULL"
        dal.RunQuery(query)
        MainForm.RemainingParts.DataSource = dal.SQLDataset01.Tables(0)
        MainForm.RemainingParts.AutoResizeColumns()
        MainForm.RemainingParts.Update()
    End Sub

    Public Function RecentOrder()
        Dim dal As New DALControl
        Dim query = "SELECT SERIALNUM, DRAWINGNUM FROM SFM_RCH.Y_PANEL_BUILD_RECENTORDER"
        dal.RunQuery(query)
        Dim SN = dal.SQLDataset01.Tables(0).Rows(0).Item(0)
        Dim DN = dal.SQLDataset01.Tables(0).Rows(0).Item(1)
        Return SN & "," & DN
    End Function

    Public Sub PastOrder(SN As String, DN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_PASTORDERS WHERE SERIALNUM = '" & SN & "'"
        dal.RunQuery(query)

        If dal.SQLDataset01.Tables(0).Rows.Count = 0 Then
            Dim dal2 As New DALControl
            Dim query2 = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_PASTORDERS (SERIALNUM, DRAWINGNUM) VALUES ('" & SN & "', '" & DN & "')"
            dal2.RunQuery(query2)
        End If
    End Sub

    Public Function CheckExtraction(DN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & DN & "'"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

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

    Public Function GetFirstItem(DN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & DN & "' AND COMPLETED = 0 AND PART_NUMBER <> '72480649001:1' ORDER BY X1, Y1"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetStorageLocation(PN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_CURRENT_STORAGE WHERE PART_NUM = '" & PN & "' AND FILL <> 0 ORDER BY SLOT"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetPartSize(PN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_PARTSIZE WHERE PART_NUMBER = '" & PN & "'"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetOrigin(name As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_ORIGINS WHERE NAME = '" & name & "'"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetPanel(DN As String)
        Dim dal As New DALControl
        Dim query = "SELECT PART_NUMBER FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & DN & "' AND X1 = 0 AND Y1 = 0 AND Z1 = 0"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0).Item(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetPanelHole(sPanel As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_PARTSIZE WHERE PART_NUMBER LIKE '%" & sPanel & "%' AND SCREW_TYPE = 'PANEL'"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetRobotOffsets(PN As String)
        Dim dal As New DALControl
        Dim query = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS WHERE PART_NUMBER LIKE '%" & PN & "%'"
        dal.RunQuery(query)
        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    'Unused as of now
    Public Function SendActivePart(PN As String)
        Try
            Dim subString As String = Left(PN, 6)
            Dim dal As New DALControl
            Dim query = "UPDATE SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS SET ACTIVE = 'X' WHERE PART_NUMBER LIKE '%" & subString & "%'"
            dal.RunQuery(query)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetScrewDepthOffset()
        Dim dal As New DALControl
        Dim query = "SELECT DEPTH, OFFSET FROM SFM_RCH.Y_PANEL_BUILD_OFFSETS WHERE ACTIVE IS NOT NULL"
        dal.RunQuery(query)

        If dal.SQLDataset01.Tables(0).Rows.Count <> 0 Then
            Return dal.SQLDataset01.Tables(0).Rows(0)
        Else
            Return Nothing
        End If
    End Function

    Public Function SendComplete(PN As String, DN As String)
        Try
            Dim dal As New DALControl
            Dim query = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 1 WHERE PART_NUMBER = '" & PN & "' AND DRAWINGNUM = '" & DN & "'"
            dal.RunQuery(query)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function SetActiveComplete()
        Try
            Dim dal As New DALControl
            Dim query = "UPDATE SFM_RCH.Y_PANEL_BUILD_ROBOT_OFFSETS SET ACTIVE = NULL WHERE ACTIVE IS NOT NULL"
            dal.RunQuery(query)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Sub FuseAssemblyExtract(DN As String)
        Dim DAL As New DALControl
        Dim i As Integer = 1
        Try
            'CREATES OUR COUNTER FOR OUR LOOP
            Dim countQuery As String = "Select COUNT(*) FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER Like '%72180560819%' and COMPLETED = 0 and DRAWINGNUM = '" & DN & "'"
            DAL.RunQuery(countQuery)
            Dim counter As Integer = CInt(DAL.SQLDataset01.Tables(0).Rows(0).Item(0))

            'SO IT DOESNT THROW AN ERROR
            If counter > 0 Then
                'SELECTS SERIAL NUMBER FOR INSERT STATEMENT
                Dim selectQuery As String = "SELECT * FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0 and DRAWINGNUM = '" & DN & "'"
                DAL.RunQuery(selectQuery)

                'LOOPS AND REPLACES ALL ASSEMBLY PARTS WITH SEPERATE PARTS
                While counter >= i

                    'X1 VALUE FOR THE PART
                    Dim selectX1 As String = "SELECT X1 FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0 and DRAWINGNUM = '" & DN & "'"
                    DAL.RunQuery(selectX1)
                    Dim X1 As Double = DAL.SQLDataset01.Tables(0).Rows(0).Item(0)

                    'Y1 VALUE FOR THE PART
                    Dim selectY1 As String = "SELECT Y1 FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE PART_NUMBER LIKE '%72180560819%' AND COMPLETED = 0 and DRAWINGNUM = '" & DN & "'"
                    DAL.RunQuery(selectY1)
                    Dim Y1 As Double = DAL.SQLDataset01.Tables(0).Rows(0).Item(0)

                    'ROUNDED VARS
                    Dim roundedX1 = Math.Round(X1, 1)
                    Dim roundedY1 = Math.Round(Y1, 1)

                    'CHECKS THE LOCATION OF THE PART IN X AND Y DIRECTION AND INSERTS THE TWO PARTS IN THAT PLACE
                    If roundedX1 = 12.6 Then
                        If roundedY1 = 24.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:11', 12.4925, 24.3375, 0.1196, 15.7575 , 26.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:11', 15.875, 24.2935, 0.1196, 19.375, 26.3435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 19.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:12', 12.4925, 19.8375, 0.1196, 15.7575 , 21.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:12', 15.875, 19.7935, 0.1196, 19.375, 21.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 15.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:13', 12.4925, 15.3375, 0.1196, 15.7575 , 17.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:13', 15.875, 15.2935, 0.1196, 19.375, 17.3475, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 10.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:14', 12.4925, 10.8375, 0.1196,15.7575 , 12.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:14', 15.875, 10.7935, 0.1196, 19.375, 12.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 6.3 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:15', 12.4925, 6.3375, 0.1196,15.7575 , 8.2875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:15', 15.875, 6.2935, 0.1196, 19.375, 8.3435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        ElseIf roundedY1 = 1.8 Then
                            Dim dal2 As New DALControl
                            Dim insertQueryFUSE As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', '871148032:16', 12.4925, 1.8375, 0.1196,15.7575 , 3.7875, 2.42, 2.6421, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQueryFUSE)
                            Dim insertQuerySwitch As String = "INSERT INTO SFM_RCH.Y_PANEL_BUILD_POSITIONS (DRAWINGNUM, PART_NUMBER, X1, Y1, Z1, X2, Y2, Z2, VOLUME, COMPLETED, REQUEST_DATE) VALUES ('" & DN & "', 'W 666443:16', 15.875, 1.7935, 0.1196, 19.375, 3.8435, 2.42, 3.54, 0,'" & DateAndTime.Now & "')"
                            dal2.RunQuery(insertQuerySwitch)
                        End If
                    End If

                    'UPDATES THE ASSEMBLY TO COMPLETE
                    Dim updatePart As String = "UPDATE SFM_RCH.Y_PANEL_BUILD_POSITIONS SET COMPLETED = 1 WHERE PART_NUMBER LIKE '%72180560819%' AND X1 = " & X1.ToString & " AND Y1 = " & Y1.ToString & " AND DRAWINGNUM = '" & DN & "'"
                    DAL.RunQuery(updatePart)
                    i += 1
                End While
                Dim DalDelete As New DALControl
                Dim deletePart As String = "DELETE FROM SFM_RCH.Y_PANEL_BUILD_POSITIONS WHERE DRAWINGNUM = '" & DN & "' AND PART_NUMBER LIKE '%72180560819%'"
                DalDelete.RunQuery(deletePart)
            End If
        Catch ex As Exception
            MsgBox("Error: FuseAssemblyExtract() failed to extract..")
        End Try
    End Sub
End Module
