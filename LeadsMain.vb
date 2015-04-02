Imports System
Imports System.Collections
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Access
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.VisualBasic.DateAndTime
Public Class frmMain
    Public Structure AgentInfo
        Dim MyID As String          '
        Dim FirstName As String     '
        Dim LastName As String      '
        Dim MyEmail As String       '
        Dim MgrEmail As String      '
        Dim MyFolder As String      '
        Dim PPL As Int16            '
        Dim LastLeadDate As String  '
        Dim NumberOfLeads As String '
        Dim MySourceFile As String

        Dim MyNewLeadFile As String
        Dim NewLeadAmount As String

        Dim IsSelected As Boolean
        Dim IsProcessed As String 'Started, Skipped, Finished
    End Structure
    Dim SelectedAgents() As Int16
    Public Agents() As AgentInfo
    Dim m_SortingColumn As ColumnHeader
    Dim ExApp As New Excel.Application
    Dim ExWB As Workbook
    Dim ExWS As Worksheet
    Dim NewLeadWB As Workbook
    Dim NewLeadWS As Worksheet
    Public Sub LoadAgents()
        Dim AgentCount As Int16 = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows.Count - 1
        ReDim Agents(0 To AgentCount)

        'load base agent data
        For counter = 0 To AgentCount
            Agents(counter).MyID = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Agent_ID").ToString
            Agents(counter).FirstName = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Agent_First").ToString
            Agents(counter).LastName = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Agent_Last").ToString
            Agents(counter).MyEmail = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Agents_Email").ToString
            'ensure Manager Email isnt blank
            If Not IsDBNull(_GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Managers_Email")) Then
                Agents(counter).MgrEmail = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Managers_Email")
            Else
                Agents(counter).MgrEmail = ""
            End If
            'ensure Lead Folder isnt blank
            If Not IsDBNull(_GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Lead_Folder")) Then
                Agents(counter).MyFolder = _GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query.Rows(counter).Item("Lead_Folder")
            Else
                Agents(counter).MyFolder = "NONE"
            End If
        Next
        'load last lead date for all agents
        For counter = 0 To _GBA_DB_2_fe___PROGRAMDataSet.NewPPL_Pre_Query.Rows.Count - 1
            For incounter = 0 To AgentCount
                If Agents(incounter).MyID <> _GBA_DB_2_fe___PROGRAMDataSet.NewPPL_Pre_Query.Rows(counter).Item("Agent_ID") Then
                    Continue For
                Else
                    Agents(incounter).LastLeadDate = _GBA_DB_2_fe___PROGRAMDataSet.NewPPL_Pre_Query.Rows(counter).Item("MaxOfReceived_Date")
                    Exit For
                End If
            Next
        Next
        'load PPL/Lead data for agents with PPL
        For counter = 0 To _GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query.Rows.Count - 1
            For incounter = 0 To AgentCount
                If Agents(incounter).LastName <> _GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query.Rows(counter).Item("Agent_Last") Or Agents(incounter).FirstName <> _GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query.Rows(counter).Item("Agent_First") Then
                    Continue For
                Else
                    Agents(incounter).PPL = _GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query.Rows(counter).Item("EXPR2").ToString
                    Agents(incounter).NumberOfLeads = _GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query.Rows(counter).Item("Num_Leads")
                    Exit For
                End If
            Next
        Next
        'check/verify for Null data - PPL, NumLeads, LastLeadDate
        For counter = 0 To AgentCount
            If Agents(counter).PPL = Nothing Then
                Agents(counter).PPL = "0"
            End If
            If Agents(counter).NumberOfLeads = Nothing Then
                Agents(counter).NumberOfLeads = "0"
            End If
            If Agents(counter).LastLeadDate = Nothing Or Agents(counter).LastLeadDate = "" Then
                Agents(counter).LastLeadDate = "None Given"
            End If
        Next
    End Sub
    Public Sub LoadDB()
        frmLoading.Activate()
        frmLoading.Visible = True

        frmLoading.prgbrOverall.Maximum = 6
        frmLoading.prgbrOverall.Value = 1
        frmLoading.prgbrOverall.Step = 1
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Inventory Table"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_beDataSet.Inventory' table. You can move, or remove it, as needed.
        Me.InventoryTableAdapter.Fill(Me.GBA_DB_2_beDataSet.Inventory)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Leads Table"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_beDataSet.Leads' table. You can move, or remove it, as needed.
        Me.LeadsTableAdapter.Fill(Me.GBA_DB_2_beDataSet.Leads)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Active Agent Data"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_feDataSet.Active_Agent_Query' table. You can move, or remove it, as needed.
        Me.Active_Agent_QueryTableAdapter.Fill(Me._GBA_DB_2_fe___PROGRAMDataSet.Active_Agent_Query)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Initial PPL Data"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_feDataSet.Calculated_PPL_Query' table. You can move, or remove it, as needed.
        Me.Calculated_PPL_QueryTableAdapter.Fill(Me._GBA_DB_2_fe___PROGRAMDataSet.Calculated_PPL_Query)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Remaining PPL Data"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_feDataSet.NewPPL_Pre_Query' table. You can move, or remove it, as needed.
        Me.NewPPL_Pre_QueryTableAdapter.Fill(Me._GBA_DB_2_fe___PROGRAMDataSet.NewPPL_Pre_Query)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.lblOverall.Text = "Loading Remaining Inventory Data"
        frmLoading.lblOverall.Refresh()
        'TODO: This line of code loads data into the 'GBA_DB_2_feDataSet.Inventory_Full_Query' table. You can move, or remove it, as needed.
        Me.Inventory_Full_QueryTableAdapter.Fill(Me._GBA_DB_2_fe___PROGRAMDataSet.Inventory_Full_Query)
        frmLoading.prgbrOverall.PerformStep()
        frmLoading.prgbrOverall.Update()

        frmLoading.Close()
    End Sub
    Public Sub LoadList()
        'loads the agent list
        For counter = 0 To Agents.Length - 1
            Dim LVItem As New ListViewItem

            LVItem.Text = Agents(counter).LastName & ", " & Agents(counter).FirstName

            LVItem.SubItems.Add(Agents(counter).PPL)
            LVItem.SubItems.Add(Agents(counter).LastLeadDate)

            lvAgentList.Items.Add(LVItem)
        Next
    End Sub
    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        MsgBox("Hello Jason.")
        'Load DataTables
        LoadDB()
        'Load Agent Info
        LoadAgents()
        'Load listview box
        LoadList()
    End Sub
    Private Sub lvAgentList_ColumnClick(sender As System.Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lvAgentList.ColumnClick
        'allows the user to sort each column in the Agent List
        Dim Sort_Column As ColumnHeader = lvAgentList.Columns(e.Column)
        Dim Sort_Order As System.Windows.Forms.SortOrder

        If m_SortingColumn Is Nothing Then
            Sort_Order = SortOrder.Ascending
        Else
            If Sort_Column.Equals(m_SortingColumn) Then
                If m_SortingColumn.Text.StartsWith("> ") Then
                    Sort_Order = SortOrder.Descending
                Else
                    Sort_Order = SortOrder.Ascending
                End If
            Else
                Sort_Order = SortOrder.Ascending
            End If
            m_SortingColumn.Text = m_SortingColumn.Text.Substring(2)
        End If


        m_SortingColumn = Sort_Column

        If Sort_Order = SortOrder.Ascending Then
            m_SortingColumn.Text = "> " & m_SortingColumn.Text
        Else
            m_SortingColumn.Text = "< " & m_SortingColumn.Text
        End If

        lvAgentList.ListViewItemSorter = New clsListviewSorter(e.Column, Sort_Order)
        lvAgentList.Sort()
    End Sub
    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        'when Add is clicked, it checks which Agents have been checked
        'these are added to the Selected box, and the isSelected flag is set for those agents
        For counter = 0 To Agents.Length - 1
            If lvAgentList.Items(counter).Selected = True Then
                For incounter = 0 To Agents.Length - 1
                    If Agents(incounter).IsSelected = True Then
                        Continue For
                    End If
                    If lvAgentList.Items(counter).Text = Agents(incounter).LastName & ", " & Agents(incounter).FirstName Then
                        lstSelectedAgents.Items.Add(Agents(incounter).LastName & ", " & Agents(incounter).FirstName)
                        Agents(incounter).IsSelected = True
                        Exit For
                    End If
                Next
            End If
        Next
        btnGO.Enabled = True
    End Sub
    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        For counter = 0 To lstSelectedAgents.Items.Count - 1
            For incounter = 0 To Agents.Length - 1
                If lstSelectedAgents.SelectedItem = Agents(incounter).LastName & ", " & Agents(incounter).FirstName Then
                    lstSelectedAgents.Items.Remove(lstSelectedAgents.SelectedItem)
                    Agents(incounter).IsSelected = False
                    Exit For
                End If
            Next
        Next
        If lstSelectedAgents.Items.Count = 0 Then
            btnGO.Enabled = False
        End If
    End Sub
    Private Sub btnGO_Click(sender As System.Object, e As System.EventArgs) Handles btnGO.Click
        'lock selection fields
        lvAgentList.Enabled = False
        lstSelectedAgents.Enabled = False
        btnAdd.Enabled = False
        btnRemove.Enabled = False
        'hide Start button
        btnGO.Visible = False
        'enable excel controls
        grpbxOptions.Enabled = True
        cmbSender.Enabled = True
        cmbAmount.Enabled = True


        'initiate for Processing
        For counter = 0 To Agents.Length - 1
            If Agents(counter).IsSelected = True Then
                Try
                    ReDim Preserve SelectedAgents(0 To SelectedAgents.Length)
                Catch ex As System.Exception
                    ReDim Preserve SelectedAgents(0)
                End Try
                SelectedAgents(SelectedAgents.Length - 1) = counter
                Agents(counter).IsProcessed = "Started"
            End If
        Next

        'call first agent
        GetSourceFile(Agents(SelectedAgents(0)).MyFolder, SelectedAgents(0))
    End Sub
    Public Sub GetSourceFile(ByVal FolderPath As String, ByVal AgentIndex2 As Int16)
        lblWho.Text = Agents(AgentIndex2).FirstName & " " & Agents(AgentIndex2).LastName

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Excel Files (*.xls)|*.xls"
        OpenFileDialog1.FilterIndex = 2
        Dim SlashPos As Integer = FolderPath.LastIndexOf("\")
        Try
            FolderPath = FolderPath.Substring(0, SlashPos + 1)
        Catch ex As System.Exception

        End Try

        OpenFileDialog1.InitialDirectory = FolderPath
        OpenFileDialog1.Title = "Select the Source Lead File for " & Agents(AgentIndex2).FirstName & " " & Agents(AgentIndex2).LastName

        Dim dlgResult As DialogResult = OpenFileDialog1.ShowDialog()

        'check for/handle cancellation or continue
        If dlgResult = System.Windows.Forms.DialogResult.Cancel Then
            'confirm Cancel
            dlgResult = MsgBox("Do you wish to skip " & Agents(AgentIndex2).FirstName & " " & Agents(AgentIndex2).LastName & "?", vbYesNo, "Confirm")
            'cancel confirmed, mark current agent as Skipped, proceed to next agent
            If dlgResult = System.Windows.Forms.DialogResult.Yes Then
                Agents(AgentIndex2).IsProcessed = "Skipped"
                lstbxAgentStatus.Items.Add(Agents(AgentIndex2).LastName & ", " & Agents(AgentIndex2).FirstName & " - Skipped")
                'find next agent to be Started
                For counter = 0 To Agents.Length - 1
                    If Agents(counter).IsProcessed = Nothing Then
                        Continue For
                    End If
                    If Agents(counter).IsProcessed = "Started" Then
                        GetSourceFile(Agents(counter).MyFolder, counter)
                        Exit Sub
                    End If
                    If counter = SelectedAgents.Length - 1 Then
                        MsgBox("All Agents have been processed.  Terminating Program.")
                        System.Windows.Forms.Application.Exit()
                    End If
                Next
                'Cancel NOT confirmed, recall for current agent
            ElseIf dlgResult = System.Windows.Forms.DialogResult.No Then
                GetSourceFile(FolderPath, AgentIndex2)
            End If
            'File selected, proceed
        ElseIf dlgResult = System.Windows.Forms.DialogResult.OK Then
            Agents(AgentIndex2).MySourceFile = OpenFileDialog1.FileName
            lblSourceFilePath.Text = OpenFileDialog1.FileName
            'call SourceFileHandler
            SourceFileHandler(Agents(AgentIndex2).MySourceFile, FolderPath, AgentIndex2)
        End If
    End Sub
    Public Sub SourceFileHandler(ByRef SourceFile As String, ByRef FolderPath As String, ByRef AgentIndex As Int16)
        'checks validity of amount
        'updates cmbAmount

        Dim IsOpen As Boolean = IsFileOpen(SourceFile)
        'make sure file isn't already open elsewhere
        If IsOpen = True Then
            MsgBox("The file: " & SourceFile & " is in use by another user.  Please close the file before continuing.")
        End If

        Try
            ExWB = ExApp.Workbooks.Open(SourceFile)
            ExWS = CType(ExWB.Worksheets(1), Worksheet)
        Catch ex As System.Exception

        End Try
        'sort the Date column Oldest to Newest
        ExWS.Columns.Sort(ExWS.Range("A1", "A500"), XlSortOrder.xlAscending, , , , , , XlYesNoGuess.xlNo, , , XlSortOrientation.xlSortColumns, , XlSortDataOption.xlSortTextAsNumbers)

        Dim MaxRow As Int16 = 0
        Dim TempShit As String = ""
        'get the max row
        For counter = 1 To 350
            If IsNothing(ExWS.Cells(counter, "A").Value) Then
                MaxRow = counter - 1
                Exit For
            End If
            TempShit = ExWS.Cells(counter, "A").Value.ToString
            If TempShit = Nothing Or TempShit = "" Then
                MaxRow = counter - 1
                Exit For
            End If
        Next
        'make sure there are leads to give out
        If MaxRow = 0 Then
            MsgBox("The file " & SourceFile & " contains no leads.  Please choose another file.")
            Try
                ExWS = Nothing
                ExWB.Close()
            Catch ex As System.Exception

            End Try
            GetSourceFile(FolderPath, AgentIndex)
            Exit Sub
        End If

        'fill combobox
        cmbAmount.Items.Clear()
        For counter = 1 To MaxRow
            cmbAmount.Items.Add(counter)
        Next
        'default to 50 if there are 50 or more, otherwise default to max
        If MaxRow > 49 Then
            cmbAmount.Text = "50"
        Else
            cmbAmount.Text = MaxRow.ToString
        End If

        'put max in lblAmount
        lblTotalAvailable.Text = "of (" & MaxRow & ") available Leads."

    End Sub
    Private Sub btnProcess_Click(sender As System.Object, e As System.EventArgs) Handles btnProcess.Click
        'THIS IS THE MAIN ROUTINE from which the cycle is run

        'open blank spreadsheet
        'copy lines
        CopyLines(cmbAmount.Text)

        'save new file in agents folder
        'generate filename
        Dim AgentIndex3 As Integer = -1
        For counter = 0 To Agents.Length - 1
            If Agents(counter).IsProcessed = "Started" Then
                AgentIndex3 = counter
                Exit For
            End If
        Next
        'Save new File
        SaveNewFile(AgentIndex3)

        'clear rows from Source IF community is NOT checked
        If chkbxCommunity.Checked = False Then
            'clear used rows
            'ClearSourceRows(cmbAmount.Text)

            'over-write source file
            Try
                'copy source file to X:\OldLeads
                Dim slashpos As Integer = lblSourceFilePath.Text.LastIndexOf("\")
                Dim LeadFileName As String = lblSourceFilePath.Text.Substring(slashpos + 1)
                slashpos = LeadFileName.IndexOf(".")
                LeadFileName = LeadFileName.Insert(slashpos, " " & Agents(AgentIndex3).FirstName & " " & Agents(AgentIndex3).LastName)

                File.Copy(lblSourceFilePath.Text, "X:\OldLeads\" & DatePart(DateInterval.Month, Today) & "-" & DatePart(DateInterval.Day, Today) & "-" & DatePart(DateInterval.Year, Today) & "-" & LeadFileName)
            Catch ex As System.Exception

            End Try
            'clear the rows that were copied, save the source file
            Try
                ClearSourceRows(cmbAmount.Text)
                ExWS.SaveAs(lblSourceFilePath.Text)
            Catch ex As System.Exception
                MsgBox("The file: " & lblSourceFilePath.Text & " is in use by another user.  Please close the file before continuing." & vbCrLf & ex.Message)
                ClearSourceRows(cmbAmount.Text)
                ExWS.SaveAs(lblSourceFilePath.Text)
            End Try

            Try
                ExWS = Nothing
                ExWB.Close()
                ExApp.WindowState = XlWindowState.xlMinimized
            Catch ex As System.Exception

            End Try
            'add leads
            Dim LeadAmount As Integer = cmbAmount.Text
            If cmbLogAsAmount.Text <> "" Then
                LeadAmount = cmbLogAsAmount.Text
            End If

            UpdateInventory(Agents(AgentIndex3).MySourceFile)
            'update inventory, either as new entry or modify last entry
            If chkbxAdditional.Checked = False Then
                AddLeads(Agents(AgentIndex3).MyID, LeadAmount)
            Else
                AddToLastPack(AgentIndex3)
            End If

        Else
            'rename source file
            Dim NewFN As String = lblSourceFilePath.Text
            Dim ExtPos As Integer = NewFN.IndexOf(".xls")
            NewFN = NewFN.Substring(0, ExtPos) & "-cc.xls"
            ExtPos = NewFN.LastIndexOf("\")
            NewFN = NewFN.Substring(ExtPos + 1)
            Try
                ExWS = Nothing
                ExWB.Close()
                My.Computer.FileSystem.RenameFile(lblSourceFilePath.Text, NewFN)
            Catch ex As System.Exception
                MsgBox("The file: " & lblSourceFilePath.Text & " is in use by another user.  Please close the file before continuing." & vbCrLf & ex.Message)
                My.Computer.FileSystem.RenameFile(lblSourceFilePath.Text, NewFN)
            End Try

        End If

        'generate email
        CreateEmail(AgentIndex3)

        'on to the next one
        'set current as Finished, get next Started
        Agents(AgentIndex3).IsProcessed = "Finished"
        lstbxAgentStatus.Items.Add(Agents(AgentIndex3).LastName & ", " & Agents(AgentIndex3).FirstName & " - Finished")

        'find next agent to process
        AgentIndex3 = -1
        For counter = 0 To Agents.Length - 1
            If Agents(counter).IsProcessed = "Started" Then
                AgentIndex3 = counter
                Exit For
            End If
        Next

        If AgentIndex3 = -1 Then
            'finished all
            MsgBox("All Agents have been processed.  Terminating Program.")
            System.Windows.Forms.Application.Exit()
        Else
            chkbxCommunity.Checked = False
            chkbxNewest.Checked = False
            chkbxLogAsAmount.Checked = False
            chkbxAdditional.Checked = False
            cmbLogAsAmount.Text = ""
            GetSourceFile(Agents(AgentIndex3).MyFolder, AgentIndex3)
        End If
    End Sub
    Public Sub CreateEmail(ByVal AgentIndex As String)
        Dim Mailer As Object = CreateObject("Outlook.Application")
        Dim Outgoing As Object = Mailer.CreateItem(0)

        Dim BodyText As String = Agents(AgentIndex).FirstName & ", attached are your new Leads." & vbCrLf & vbCrLf

        If cmbSender.Text = "Rebecca" Then
            BodyText &= "Thanks," & vbCrLf
            BodyText &= "Rebecca Simmers" & vbCrLf
            BodyText &= "General Brokerage Agency" & vbCrLf
            BodyText &= "Phone: 610-590-4520 (*email preferred)" & vbCrLf
            BodyText &= "Fax ALL new business to: 888-350-1184"
        End If
        If cmbSender.Text = "Robert" Then
            BodyText &= "Thanks," & vbCrLf
            BodyText &= "Robert Ritschard" & vbCrLf
            BodyText &= "General Brokerage Agency" & vbCrLf
            BodyText &= "Phone: 610-590-4521 (*email preferred)" & vbCrLf
            BodyText &= "Fax ALL new business to: 888-350-1184"
        End If

        If cmbSender.Text = "Jason" Then
            BodyText &= "Thanks," & vbCrLf
            BodyText &= "Jason Siebert" & vbCrLf
            BodyText &= "General Brokerage Agency" & vbCrLf
            BodyText &= "Fax ALL new business to: 888-350-1184"
        End If

        With Outgoing
            .Subject = "Leads"
            .To = Agents(AgentIndex).MyEmail & "; " & Agents(AgentIndex).MgrEmail
            .Body = BodyText
            .Attachments.Add(Agents(AgentIndex).MyNewLeadFile)
        End With

        Outgoing.Display()

    End Sub
    Public Sub AddToLastPack(ByVal AgentIndex5 As Integer)
        'find Lead ID with Agent ID for date same as Today
        For counter = 0 To GBA_DB_2_beDataSet.Leads.Rows.Count - 1
            'find agents leads
            If GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Agent_ID") = Agents(AgentIndex5).MyID Then
                'find leads given within the last 48 hours
                If GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Received_Date") >= Today.AddDays(-1) Then

                    Dim NewAmount As Integer = 0

                    Dim OldAmount As Integer = GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Num_Leads")
                    'calculate new amount
                    If cmbLogAsAmount.Text = "" Then
                        NewAmount = OldAmount + CInt(cmbAmount.Text)
                    Else
                        NewAmount = OldAmount + CInt(cmbLogAsAmount.Text)
                    End If

                    'summary message box
                    Dim msgResult As MsgBoxResult = MsgBox("The last Lead Pack for " & Agents(AgentIndex5).FirstName & " " & _
                                                           Agents(AgentIndex5).LastName & " was on " & GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Received_Date") & _
                                                           " for " & OldAmount & " leads.  Add " & (NewAmount - OldAmount) & " for a total of " & NewAmount & "?" _
                                                           , MsgBoxStyle.YesNo, "Verify")

                    'if not verified, log as new entry
                    If msgResult = MsgBoxResult.No Then
                        MsgBox("Adding a new entry.")
                        UpdateInventory(lblSourceFilePath.Text)
                        Exit Sub
                    End If

                    Dim InsertString As String = ""
                    Dim ConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=W:\@database\Back End\GBA_DB_2_be.accdb"

                    Dim InvConn As New System.Data.OleDb.OleDbConnection(ConnString)
                    Dim UpdateCmd As New OleDbCommand(InsertString, InvConn)

                    InsertString = "UPDATE [Leads] SET [Num_Leads] = @Count WHERE Lead_ID= @ID;"

                    UpdateCmd.CommandText = (InsertString)
                    UpdateCmd.Connection = (InvConn)

                    UpdateCmd.Parameters.Clear()

                    UpdateCmd.Parameters.AddWithValue("@Count", NewAmount)
                    UpdateCmd.Parameters.AddWithValue("@ID", CInt(GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Lead_ID")))

                    Try
                        InvConn.Open()
                        UpdateCmd.ExecuteNonQuery()
                    Catch ex As System.Exception
                        MsgBox("Unable to access the Inventory DataSet.  Please manually update the Inventory." & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
                    Finally
                        InvConn.Close()
                    End Try

                    Exit Sub
                End If
            End If
        Next
        'if agent doesnt have leads (or within the last 48 hours) log as new entry
        MsgBox("Unable to find a recent entry for " & Agents(AgentIndex5).FirstName & " " & Agents(AgentIndex5).LastName & ".  Creating a new Lead entry.")
        UpdateInventory(lblSourceFilePath.Text)
        Exit Sub
    End Sub
    Public Sub UpdateInventory(ByVal SourceFile As String)
        'cycle inventory for sourcefile match, get ID
        Dim InvID As String = ""
        For counter = 0 To GBA_DB_2_beDataSet.Inventory.Rows.Count - 1
            If GBA_DB_2_beDataSet.Inventory.Rows(counter).Item("Lead_File") = SourceFile Then
                InvID = GBA_DB_2_beDataSet.Inventory.Rows(counter).Item("ID")
                Exit For
            End If
        Next

        Dim NewCount As String = (cmbAmount.Items(cmbAmount.Items.Count - 1) - cmbAmount.Text).ToString

        Dim InsertString As String = ""
        Dim ConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=W:\@database\Back End\GBA_DB_2_be.accdb"

        Dim InvConn As New System.Data.OleDb.OleDbConnection(ConnString)
        Dim UpdateCmd As New OleDbCommand(InsertString, InvConn)

        Dim DateGiven As String = Today.ToString
        'update Count first
        InsertString = "UPDATE [Inventory] SET [Count] = @Count WHERE ID= @ID;"

        UpdateCmd.CommandText = (InsertString)
        UpdateCmd.Connection = (InvConn)

        UpdateCmd.Parameters.Clear()

        UpdateCmd.Parameters.AddWithValue("@Count", NewCount)
        UpdateCmd.Parameters.AddWithValue("@ID", InvID)

        Try
            InvConn.Open()
            UpdateCmd.ExecuteNonQuery()
        Catch ex As System.Exception
            MsgBox("Unable to access the Inventory DataSet.  Please manually update the Inventory." & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        Finally
            InvConn.Close()
        End Try
        'update date Leads were given (today)
        InsertString = "UPDATE [Inventory] SET [Last_Date_Given] = @LDG WHERE ID= @ID;"

        UpdateCmd.CommandText = (InsertString)
        UpdateCmd.Connection = (InvConn)

        UpdateCmd.Parameters.Clear()

        UpdateCmd.Parameters.AddWithValue("@LDG", DateGiven)
        UpdateCmd.Parameters.AddWithValue("@ID", InvID)

        Try
            InvConn.Open()
            UpdateCmd.ExecuteNonQuery()
        Catch ex As System.Exception
            MsgBox("Unable to access the Inventory DataSet.  Please manually update the Inventory." & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        Finally
            InvConn.Close()
        End Try

    End Sub
    Public Sub AddLeads(ByVal AgentID As String, ByVal LeadAmount As String)
        'logs the leads 
        Dim FinMax As Integer = 0
        Dim TempInt As Integer = 0
        For counter = 0 To GBA_DB_2_beDataSet.Leads.Rows.Count - 1
            TempInt = Convert.ToInt16(GBA_DB_2_beDataSet.Leads.Rows(counter).Item("Lead_ID"))
            If TempInt > FinMax Then
                FinMax = TempInt
            End If
        Next

        FinMax += 1

        Dim LRStr As String = FinMax.ToString()
        Dim ToDate As String = "#" & Month(Now).ToString() & "/" & Day(Now).ToString() & "/" & Year(Now).ToString() & "#"
        Dim InsertString As String = "INSERT INTO Leads (Lead_ID, Num_Leads, Received_Date, End_Date, Agent_ID) VALUES (" & LRStr & ", " & LeadAmount & ", " & ToDate & ", " & ToDate & ", " & AgentID & ")"

        Dim ConnString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=W:\@database\Back End\GBA_DB_2_be.accdb"
        Dim LeadsConn As New System.Data.OleDb.OleDbConnection(ConnString)
        Dim InsertCmd As New OleDbCommand(InsertString, LeadsConn)

        Try
            LeadsConn.Open()
            InsertCmd.ExecuteNonQuery()
            LeadsConn.Close()
        Catch ex As System.Exception
            MsgBox("Unable to Update leads for this agent.  Please verify today's leads in the Database and make the appropriate changes.")
        End Try

        Me.LeadsTableAdapter.Fill(Me.GBA_DB_2_beDataSet.Leads)
    End Sub
    Public Function CleanFileName(ByVal FileName As String)
        '. / ? < > \ : * |
        FileName = FileName.Replace(".", "")
        FileName = FileName.Replace("/", "")
        FileName = FileName.Replace("?", "")
        FileName = FileName.Replace("<", "")
        FileName = FileName.Replace(">", "")
        FileName = FileName.Replace("\", "")
        FileName = FileName.Replace(":", "")
        FileName = FileName.Replace("*", "")
        FileName = FileName.Replace("|", "")

        Return FileName
    End Function
    Public Sub SaveNewFile(ByRef AgentIndex As Integer)

        'ExWS.SaveAs(frmLeads.lblSourceFile.Text)
        SaveFileDialog1.Filter = "Excel Workbook (*.xls)|*.xls"
        SaveFileDialog1.InitialDirectory = Agents(AgentIndex).MyFolder

        'get Area
        Dim SourceFolder As DirectoryInfo = Directory.GetParent(lblSourceFilePath.Text)
        Dim FNIndex As Integer = SourceFolder.FullName.LastIndexOf("\")
        Dim FN As String = SourceFolder.FullName.Substring(FNIndex + 1)

        'remove any dots to avoid errant filenames
        FN = CleanFileName(FN)

        'process according to Community Leads
        If chkbxCommunity.Checked = False Then
            'get Date
            Dim TempDate As Date = Today
            Dim ThisDate As String = TempDate.ToString("MMddyy")

            Dim LogAmount As String = cmbAmount.SelectedIndex + 1
            If cmbLogAsAmount.Text <> "" Then
                LogAmount &= "=" & cmbLogAsAmount.Text
            End If

            SaveFileDialog1.FileName = LogAmount & "-" & FN & "-" & ThisDate
        ElseIf chkbxCommunity.Checked = True Then
            Dim NewFN As String = lblSourceFilePath.Text
            Dim SlashPos As Integer = NewFN.LastIndexOf("\")
            NewFN = NewFN.Substring(SlashPos + 1)
            'cut off old extension
            Dim TempFN As String = NewFN.Substring(0, NewFN.Length - 4)
            'append COMMUNITY rename
            NewFN = TempFN & "-COMMUNITY.xls"

            SaveFileDialog1.FileName = NewFN
        End If

        'SAVE IT
        Me.BringToFront()
        If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            NewLeadWB.SaveAs(SaveFileDialog1.FileName, -4143)
            Agents(AgentIndex).MyNewLeadFile = SaveFileDialog1.FileName
        Else
            MsgBox("You must save the new file in order to continue.")

            SaveFileDialog1.ShowDialog()
            NewLeadWB.SaveAs(SaveFileDialog1.FileName, -4143)
            Agents(AgentIndex).MyNewLeadFile = SaveFileDialog1.FileName
        End If

        'after New is saved, close it
        Try
            NewLeadWS = Nothing
            NewLeadWB.Close()
        Catch ex As System.Exception

        End Try
    End Sub
    Public Sub CopyLines(ByVal Lines As Integer)
        'open new workbook
        Try
            NewLeadWB = ExApp.Workbooks.Add(System.Reflection.Missing.Value)
            NewLeadWS = CType(NewLeadWB.Worksheets(1), Worksheet)
        Catch ex As System.Exception

        End Try

        ExWS.Columns.Sort(ExWS.Range("A1", "A500"), XlSortOrder.xlAscending, , , , , , XlYesNoGuess.xlNo, , , XlSortOrientation.xlSortColumns, , XlSortDataOption.xlSortTextAsNumbers)
        ExApp.Visible = True
        ExWS.Visible = CType(True, XlSheetVisibility)
        NewLeadWS.Visible = CType(True, XlSheetVisibility)

        'display both worksheets
        ExApp.WindowState = XlWindowState.xlMaximized
        ExApp.Windows.Arrange(XlArrangeStyle.xlArrangeStyleVertical)

DoIt:
        'check for Newest Leads
        Dim StartAt As Integer = 1
        If chkbxNewest.Checked = True Then
            StartAt = cmbAmount.Items(cmbAmount.Items.Count - 1) - Lines - 1
        End If
        'grab range
        Dim ExRange As Excel.Range = ExWS.Range(ExWS.Cells(StartAt, 1), ExWS.Cells(StartAt + Lines - 1, 4))
        ExRange.Copy()
        'paste range
        Dim NewRg As Excel.Range = CType(NewLeadWS.Cells(1, 1), Range)
        NewLeadWS.Paste()
        'fit columns
        For counter = 1 To 4
            NewLeadWS.Columns(counter).AutoFit()
        Next
        'make sure text was copied properly
        If ExWS.Cells(cmbAmount.Text, "A").Value.ToString = "" Then
            Dim msgCheck As MsgBoxResult = MsgBox("Possible error during copy.  Have the leads been correctly copied to the new file?", MsgBoxStyle.YesNo, "Alert")
            If msgCheck = MsgBoxResult.No Then
                GoTo DoIt
            Else
                Exit Sub
            End If
        End If

    End Sub
    Public Function IsFileOpen(ByVal FileName As String) As Boolean
        Dim IsOpen As Boolean = False
        Try
            'CREATE A FILE STREAM FROM THE FILE, OPENING IT FOR READ ONLY EXCLUSIVE ACCESS
            Dim FS As IO.FileStream = IO.File.Open(FileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
            'CLOSE AND CLEAN UP RIGHT AWAY, IF THE OPEN SUCCEEDED, WE HAVE OUR ANSWER ALREADY
            FS.Close()
            FS.Dispose()
            FS = Nothing
            'MessageBox.Show("Yes, you are the only one using this file")

        Catch ex As IO.IOException
            'IF AN IO EXCEPTION IS THROWN, WE COULD NOT GET EXCLUSIVE ACCESS TO THE FILE
            'MessageBox.Show("No someone else has this file open" & Environment.NewLine & ex.Message)
            IsOpen = True
        Catch ex As System.Exception
            MessageBox.Show("Unknown error occured" & Environment.NewLine & ex.Message)
        End Try

        Return IsOpen
    End Function
    Private Sub cmbSender_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSender.SelectedIndexChanged
        btnProcess.Enabled = True
    End Sub
    Public Sub ClearSourceRows(ByVal Lines As Integer)
        Dim StartRow As Integer = 1
        'get row range
        If chkbxNewest.Checked = True Then
            StartRow = cmbAmount.Items(cmbAmount.Items.Count - 1) - Lines + 1
            Lines = cmbAmount.Items(cmbAmount.Items.Count - 1)
        End If

        Dim Rg As Excel.Range = ExWS.Range(ExWS.Rows(StartRow), ExWS.Rows(Lines))
        Rg.Delete()
    End Sub
    Private Sub frmMain_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'clean your room!
        Try
            ExWS = Nothing
            ExWB.Close()
            NewLeadWS = Nothing
            NewLeadWB.Close()
            ExApp.Quit()
        Catch ex As System.Exception

        End Try
    End Sub
    Private Sub lvAgentList_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles lvAgentList.MouseDoubleClick
        'when Add is clicked, it checks which Agents have been checked
        'these are added to the Selected box, and the isSelected flag is set for those agents
        For counter = 0 To Agents.Length - 1
            If lvAgentList.Items(counter).Selected = True Then
                For incounter = 0 To Agents.Length - 1
                    If Agents(incounter).IsSelected = True Then
                        Continue For
                    End If
                    If lvAgentList.Items(counter).Text = Agents(incounter).LastName & ", " & Agents(incounter).FirstName Then
                        lstSelectedAgents.Items.Add(Agents(incounter).LastName & ", " & Agents(incounter).FirstName)
                        Agents(incounter).IsSelected = True
                        Exit For
                    End If
                Next
            End If
        Next
        btnGO.Enabled = True
    End Sub
End Class
Public Class clsListviewSorter
    Implements System.Collections.IComparer ' Implements a comparer Implements IComparer 
    Private m_ColumnNumber As Integer
    Private m_SortOrder As SortOrder
    Public Sub New(ByVal column_number As Integer, ByVal sort_order As SortOrder)
        m_ColumnNumber = column_number
        m_SortOrder = sort_order
    End Sub
    ' Compare the items in the appropriate column 
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim item_x As ListViewItem = DirectCast(x, ListViewItem)
        Dim item_y As ListViewItem = DirectCast(y, ListViewItem)
        ' Get the sub-item values. 
        Dim string_x As String
        If item_x.SubItems.Count <= m_ColumnNumber Then
            string_x = ""
        Else
            string_x = item_x.SubItems(m_ColumnNumber).Text
        End If
        Dim string_y As String
        If item_y.SubItems.Count <= m_ColumnNumber Then
            string_y = ""
        Else
            string_y = item_y.SubItems(m_ColumnNumber).Text
        End If
        ' Compare them. 
        If m_SortOrder = SortOrder.Ascending Then
            If IsNumeric(string_x) And IsNumeric(string_y) Then
                Return (Val(string_x).CompareTo(Val(string_y)))
            ElseIf IsDate(string_x) And IsDate(string_y) Then
                Return (DateTime.Parse(string_x).CompareTo(DateTime.Parse(string_y)))
            Else
                Return (String.Compare(string_x, string_y))
            End If
        Else
            If IsNumeric(string_x) And IsNumeric(string_y) Then
                Return (Val(string_y).CompareTo(Val(string_x)))
            ElseIf IsDate(string_x) And IsDate(string_y) Then
                Return (DateTime.Parse(string_y).CompareTo(DateTime.Parse(string_x)))
            Else
                Return (String.Compare(string_y, string_x))
            End If
        End If
    End Function
End Class
