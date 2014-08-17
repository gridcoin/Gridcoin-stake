﻿Imports Finisar.SQLite
Imports System.Windows.Forms
Imports System.Text

Public Class frmSQL
    Private mData As Sql

    Private Sub frmSQL_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call tSync_Tick(Nothing, Nothing)

    End Sub

    Private Sub frmSQL_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ReplicateDatabase("gridcoin_ro")

        mData = New Sql("gridcoin_ro")


        'Query available tables
        Dim dr As GridcoinReader

        lvTables.View = Windows.Forms.View.Details
        Dim h1 As New System.Windows.Forms.ColumnHeader

        dr = mData.GetGridcoinReader("SELECT * FROM sqlite_master WHERE type='table';")

        'dr = mData.Query(".tables;")
        lvTables.Columns.Clear()
        lvTables.Columns.Add("Table")
        lvTables.Columns.Add("Rows")
       
        lvTables.Columns(0).Width = (lvTables.Width * 0.59) - 3
        lvTables.Columns(1).Width = (lvTables.Width * 0.41) - 3

        lvTables.FullRowSelect = True
        lvTables.HeaderStyle = Windows.Forms.ColumnHeaderStyle.Nonclickable
        Dim iRow As Long
        AddHandler lvTables.DrawColumnHeader, AddressOf lvTables_DrawColumnHeader
        AddHandler lvTables.DrawSubItem, AddressOf lvTables_DrawSubItem
        Dim lRC As Long
        Dim grr As GridcoinReader.GridcoinRow

        For y = 1 To dr.Rows
            grr = dr.GetRow(y)
            Dim sTable As String = grr.Values(1)
            Dim lvItem As New System.Windows.Forms.ListViewItem(sTable)
            Dim sql As String
            sql = "Select count(*) as Count1 from " + Trim(sTable) + ";"
            lRC = Val(mData.QueryFirstRow(sql, "Count1"))
            lvItem.SubItems.Add(Trim(lRC))
            lvItem.BackColor = Drawing.Color.Black
            lvItem.ForeColor = Drawing.Color.Lime
            lvTables.Items.Add(lvItem)
            iRow = iRow + 1
        Next y
  
        lvTables.BackColor = Drawing.Color.Black
        lvTables.ForeColor = Drawing.Color.Lime

    End Sub

    Private Sub lvTables_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        e.Graphics.FillRectangle(Drawing.Brushes.Black, e.Bounds)
        e.Graphics.DrawString(e.Header.Text, lvTables.Font, Drawing.Brushes.Lime, e.Bounds)
        e.Graphics.DrawLine(Drawing.Pens.White, e.Bounds.X, e.Bounds.Y + 15, e.Bounds.X + 40, e.Bounds.Y + 15)
    End Sub
    Private Sub lvTables_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        e.Graphics.FillRectangle(Drawing.Brushes.Black, e.Bounds)
        e.DrawText()
    End Sub
    Private Sub btnExec_Click(sender As System.Object, e As System.EventArgs) Handles btnExec.Click
        Dim dr As GridcoinReader
        mData.bThrowUIErrors = True

        Try
            dr = mData.GetGridcoinReader(rtbQuery.Text)
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Gridcoin Query Analayzer")
            Exit Sub
        End Try
        If dr Is Nothing Then Exit Sub
        dgv.Rows.Clear()
        dgv.Columns.Clear()
        dgv.BackgroundColor = Drawing.Color.Black
        dgv.ForeColor = Drawing.Color.Lime
        Dim sValue As String
        Dim iRow As Long
        If dr.Rows = 0 Then Exit Sub

        Try
            Dim grr As New GridcoinReader.GridcoinRow
            grr = dr.GetRow(1)

            For x = 0 To grr.FieldNames.Count - 1

                Dim dc As New System.Windows.Forms.DataGridViewColumn
                dc.Name = grr.FieldNames(x)
                Dim dgvct As New System.Windows.Forms.DataGridViewTextBoxCell
                dgvct.Style.BackColor = Drawing.Color.Black
                dgvct.Style.ForeColor = Drawing.Color.Lime
                dc.CellTemplate = dgvct
                dgv.Columns.Add(dc)
            Next x
            Dim dgcc As New DataGridViewCellStyle

            dgcc.ForeColor = System.Drawing.Color.SandyBrown
            dgv.ColumnHeadersDefaultCellStyle = dgcc
            For x = 0 To grr.FieldNames.Count - 1
                dgv.Columns(x).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            Next

            
            For y = 1 To dr.Rows
                grr = dr.GetRow(y)
                dgv.Rows.Add()
                For x = 0 To grr.FieldNames.Count - 1
                    sValue = grr.Values(x).ToString
                    dgv.Rows(iRow).Cells(x).Value = sValue
                Next x
                iRow = iRow + 1


            Next
            For x = 0 To grr.FieldNames.Count - 1
                dgv.Columns(x).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            Next

            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, "Gridcoin Query Analayzer")

        End Try

    End Sub
    Public Function SerializeTable(sTable As String, lStartRow As Long, lEndRow As Long) As StringBuilder
        Dim sql As String
        sql = "Select * From " + sTable + " WHERE ID >= " + Trim(lStartRow) + " AND ID <= " + Trim(lEndRow)
        Dim dr As GridcoinReader

        dr = mData.GetGridcoinReader(sql)
        Dim iRow As Long
        Dim sbOut As New StringBuilder
        Dim sRow As String
        Dim sValue As String
        Dim grr As GridcoinReader.GridcoinRow

        For y = 1 To dr.Rows
            grr = dr.GetRow(y)
            iRow = iRow + 1
            sRow = ""
            For x = 0 To grr.FieldNames.Count - 1
                sValue = "" & grr.Values(x).ToString
                sRow = sRow & sValue & "|"
            Next x
            sbOut.AppendLine(sRow)
        Next
        Return sbOut
    End Function
    Public Function GetManifestForTable(sTable As String) As String
        Dim sql As String
        sql = "Select min(id) as lmin From " + sTable
        Dim lStart As Long
        Dim lEnd As Long
        lStart = mData.QueryFirstRow(sql, "lmin")
        sql = "Select max(id) as lmax from " + sTable
        lEnd = mData.QueryFirstRow(sql, "lmax")
        Dim sOut As String
        sOut = Trim(sTable) + "," + Trim(lStart) + "," + Trim(lEnd)
        Return sOut

    End Function
    Public Function CreateManifest() As StringBuilder
        Dim dr As GridcoinReader

        dr = mData.GetGridcoinReader("SELECT * FROM sqlite_master WHERE type='table';")
        'todo order by
        Dim iRow As Long
        Dim sRow As String
        Dim sbManifest As New StringBuilder
        Dim grr As GridcoinReader.GridcoinRow

        For y = 1 To dr.Rows
            grr = dr.GetRow(y)
            Dim sTable As String = grr.Values(1)
            sRow = GetManifestForTable(sTable)
            sbManifest.AppendLine(sRow)
            iRow = iRow + 1
        Next

        Return sbManifest

    End Function
    
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        Dim s As New StringBuilder
        s = CreateManifest()
        s = SerializeTable("peers", 1, 1)
        s = SerializeTable("system", 1, 1)
    End Sub
    Private Sub rtbQuery_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles rtbQuery.KeyDown
        If e.KeyCode = Keys.F5 Then
            Call btnExec_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub btnRefreshLeaderboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim thUpdateLeaderboard As New System.Threading.Thread(AddressOf modBoincLeaderboard.RefreshLeaderboard)
            Log("Starting background update leaderboard thread.")
            thUpdateLeaderboard.IsBackground = False
            thUpdateLeaderboard.Start()
        Catch ex As Exception
            Log("UpdateLeaderboard:" + ex.Message)
        End Try

    End Sub

    Private Sub tSync_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tSync.Tick
        If modBoincLeaderboard.SQLInSync Then
            lblSync.Visible = False
            pbSync.Visible = False

        Else
            lblSync.Visible = True
            pbSync.Visible = True
            SetPb(pbSync, nBestBlock, mlSqlBestBlock)

        End If
    End Sub
    Private Function SetPb(ByVal pb As ProgressBar, ByVal valMax As Long, ByVal valActual As Long)
        Dim max As Long
        If valMax > max Then max = valMax
        If valActual > max Then max = valActual
        If max < 40000 Then max = 40000
        pb.Maximum = max
        If valActual < 1 Then valActual = 1
        If valActual > max Then valActual = max
        pb.Value = valActual
        pb.Refresh()
        pb.Update()
        Me.Update()


    End Function
End Class