Imports DAL1
Imports System.IO
Imports System.Net

Module InventorExtraction
    Private m_InventorApp As Inventor.Application = Nothing
    Private m_AddInInterface As Object
    Public Sub InitializeInventor()
        Try
            m_InventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            Dim InventorType As Type = System.Type.GetTypeFromProgID("Inventor.Application")
            m_InventorApp = System.Activator.CreateInstance(InventorType)
        End Try
    End Sub

    Public Sub ExtractInventorData(text As String)
        Dim inventorInstances() = Process.GetProcessesByName("Inventor.Exe")
        If inventorInstances.Count = 0 Then
            MsgBox("You must have inventor open in order to extract")
            Exit Sub
        End If
        InitializeInventor()
        m_InventorApp.SilentOperation = True
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


            Dim ComponentPartNumbers As New List(Of String)

            Dim outputFolder As String = "C:\Inventor\Output"
            Dim OutputFilePath = System.IO.Path.Combine(outputFolder, DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") & "_PanelExtract.csv")

            ComponentPartNumbers.Add("W 666443%") 'KS
            ComponentPartNumbers.Add("871148032%") 'FB
            ComponentPartNumbers.Add("W 662201%") 'CB
            ComponentPartNumbers.Add("W 651550%") 'PB
            ComponentPartNumbers.Add("W 651573%") 'TB
            ComponentPartNumbers.Add("W 651229%") 'TB
            ComponentPartNumbers.Add("72180560819%") '*****Fuse assembly extraction

            'Instantiate a new SI_Inventor_PPE object by passing the Inventor application, a list of component part numbers to be extracted, and the panel back plate part number
            Dim PPE As SI_Inventor_PPE.SI_Inventor_PPE = New SI_Inventor_PPE.SI_Inventor_PPE(m_InventorApp, ComponentPartNumbers, PanelPartNumber)
            'This method executes the extraction for the specified Inventor assembly file and write the data to the specified output path
            PPE.ExtractPanelData(AssyPath, OutputFilePath)
        Catch ex As Exception
            m_InventorApp.SilentOperation = False
            MessageBox.Show("Error generating output: " & ex.Message)
        End Try

    End Sub
End Module