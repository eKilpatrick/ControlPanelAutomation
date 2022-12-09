Imports System.IO
Module InventorExtract
    'Private m_InventorApp As Inventor.Application = Nothing
    Private m_AddInInterface As Object
    Public Sub InitializeInventor()
        Try
            'm_InventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            Dim InventorType As Type = System.Type.GetTypeFromProgID("Inventor.Application")
            'm_InventorApp = System.Activator.CreateInstance(InventorType)
        End Try
    End Sub

    Public Sub ExtractInventorData(text As String)
        InitializeInventor()
        'm_InventorApp.SilentOperation = True
        Try
            'TODO// ADD THE DIFFERENT PANELS TO A DATABASE SO WHEN NEW ONES ARE ADDED IT UPDATES AUTOMATICALLY
            Dim PanelPartNumber As String = "72480649001"

            'Copies all files from subdirectories in the Inventor folder and pastes them all into one location and saves the path of the assembly file
            My.Computer.FileSystem.CreateDirectory("C:\Inventor\Orders\" & text & "_Extracted")
            Dim dir As New DirectoryInfo("C:\Inventor\Orders\" & text & "\Workspaces\Workspace")
            Dim fileArr() = dir.GetFiles("*", IO.SearchOption.AllDirectories)
            Dim file As FileInfo
            For Each file In fileArr
                If file.Name <> "lockfile.lck" Then
                    My.Computer.FileSystem.CopyFile(file.FullName, "C:\Inventor\Orders\" & text & "_Extracted\" & file.Name)
                End If
            Next file
            Dim AssyPath As String = "C:\Inventor\Orders\" & text & "_Extracted\" & text & ".iam"


            MainForm.txtLog.Text &= "   Inventor Extract Started" & Environment.NewLine
            Dim ComponentPartNumbers As List(Of String) = New List(Of String)

            Dim outputFolder As String = "C:\Inventor\Output"
            Dim OutputFilePath = System.IO.Path.Combine(outputFolder, DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") & "_PanelExtract.csv")

            'This is a list of all components in the assembly that should be included in the output.  This does not include the panel part number.
            'The panel part number is passed to the SI_Inventor_PPE object when instantiating
            ' *** Wildcards can be used in part numbers.  Use a percent sign (%) instead of an asterisk for wildcards as the part numbers are being matched via a SQL query
            MainForm.txtLog.Text &= "   Inventor Extract Started Pt2" & Environment.NewLine
            'TODO// REFERENCE THESE PARTS FROM OUR STORAGE BIN INSTEAD OF BEING HARDCODED. I WOULD PUT THE PARTS IN A LIST SO WHEN DATABASE IS UPDATED THE EXTRACT IS UPDATED AS WELL
            'Try
            'Dim dalParts As New DALControl
            'Dim queryParts As String = "SELECT PART_NUMBER FROM SFM_RCH.Y_PANEL_BUILD_PARTSIZE WHERE PART_NUMBER <> '72480649001:1'"
            'dalParts.RunQuery(queryParts)
            'Dim i = 0
            'Dim row As DataRow
            'For Each row In dalParts.SQLDataset01.Tables(0).Rows
            'ComponentPartNumbers.Add(dalParts.SQLDataset01.Tables(0).Rows(i).Item(0))
            'i += 1
            'Next
            'Catch ex As Exception

            'End Try

            ComponentPartNumbers.Add("knife switch%")
            ComponentPartNumbers.Add("fuseblock%")
            ComponentPartNumbers.Add("W 666443%")
            ComponentPartNumbers.Add("W 662201%")
            ComponentPartNumbers.Add("W 651573%")
            ComponentPartNumbers.Add("W 651550%")
            ComponentPartNumbers.Add("W 651524%")
            ComponentPartNumbers.Add("W 651259%")
            ComponentPartNumbers.Add("W 651244%")
            ComponentPartNumbers.Add("W 651229%")
            ComponentPartNumbers.Add("W 651214%")
            ComponentPartNumbers.Add("W 651222%")
            'ComponentPartNumbers.Add("W 617140%")
            'ComponentPartNumbers.Add("W 551106%")
            'ComponentPartNumbers.Add("W 500025%")
            'ComponentPartNumbers.Add("72183586001%") 'what's the difference here???? pattern of TB and PB
            ComponentPartNumbers.Add("72180560819%") '*****Fuse assembly extraction
            ComponentPartNumbers.Add("871148032%")
            ComponentPartNumbers.Add("72480649001%")

            'Instantiate a new SI_Inventor_PPE object by passing the Inventor application, a list of component part numbers to be extracted, and the panel back plate part number
            'Dim PPE As SI_Inventor_PPE.SI_Inventor_PPE = New SI_Inventor_PPE.SI_Inventor_PPE(m_InventorApp, ComponentPartNumbers, PanelPartNumber)
            'This method executes the extraction for the specified Inventor assembly file and write the data to the specified output path
            'PPE.ExtractPanelData(AssyPath, OutputFilePath)
        Catch ex As Exception
            'm_InventorApp.SilentOperation = False
            MessageBox.Show("Error generating output: " & ex.Message)
        End Try
        MainForm.txtLog.Text &= "Inventor Finished Extracting. Press RUN/RESUME Order" & Environment.NewLine

    End Sub
End Module