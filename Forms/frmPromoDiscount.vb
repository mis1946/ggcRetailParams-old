Imports System.Threading
Imports System.Drawing
Imports System.Windows.Forms
Imports ggcRetailParams

Public Class frmPromoDiscount
    Private p_oDT As DataTable
    Private p_oDTSel As DataTable
    Private p_oInventory As clsInventory
    Private pnLoadx As Integer
    Private poControl As Control
    Private p_bCancelled As Boolean
    Private pnActiveRow As Integer

    ReadOnly Property Cancelled() As Boolean
        Get
            Return p_bCancelled
        End Get
    End Property

    WriteOnly Property Inventory As clsInventory
        Set(oInventory As clsInventory)
            p_oInventory = oInventory
        End Set
    End Property

    Property DataDiscount() As DataTable
        Set(ByVal oDT As DataTable)
            p_oDT = oDT
        End Set
        Get
            Return p_oDTSel
        End Get
    End Property

    Private Sub frmAddOns_Keydown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                p_bCancelled = True
                Me.Close()
                Me.Dispose()
            Case Keys.Return, Keys.Down
                SetNextFocus()
            Case Keys.Up
                SetPreviousFocus()
        End Select
    End Sub

    Private Sub frmAddOns_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If pnLoadx = 0 Then
            setVisible()

            clearFields()
            loadDetail()

            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)

            pnLoadx = 1
        End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 'ok
                Call SelectData()

                p_bCancelled = False
            Case 1 'close
                p_bCancelled = True
        End Select

        Me.Close()
        Me.Dispose()
endProc:
        Exit Sub
    End Sub

    Private Sub SelectData()
        p_oDTSel = Nothing
        p_oDTSel = p_oDT.Clone

        p_oDTSel.Rows.Add()

        Dim lnRem As Integer
        Dim lnQuotient As Integer

        lnQuotient = (Math.DivRem(p_oDT(pnActiveRow)("nQuantity"), p_oDT(pnActiveRow)("nBaseQtyx"), lnRem))
        p_oDT(pnActiveRow)("nPromBght") = p_oDT(pnActiveRow)("nMaxQtyxx") * lnQuotient

        If p_oDT(pnActiveRow)("sStockIDx") = "" And p_oDT(pnActiveRow)("xDsctdItm") <> "" Then Call searchItem(txtField01.Text)

        p_oDTSel(0)("xPromoItm") = p_oDT(pnActiveRow)("xPromoItm")
        p_oDTSel(0)("xDsctdItm") = p_oDT(pnActiveRow)("xDsctdItm")
        p_oDTSel(0)("nMinQtyxx") = p_oDT(pnActiveRow)("nMinQtyxx")
        p_oDTSel(0)("nBaseQtyx") = p_oDT(pnActiveRow)("nBaseQtyx")
        p_oDTSel(0)("nDDiscRte") = p_oDT(pnActiveRow)("nDDiscRte")
        p_oDTSel(0)("nDDiscAmt") = p_oDT(pnActiveRow)("nDDiscAmt")
        p_oDTSel(0)("nMaxQtyxx") = p_oDT(pnActiveRow)("nMaxQtyxx")
        p_oDTSel(0)("nDiscRate") = p_oDT(pnActiveRow)("nDiscRate")
        p_oDTSel(0)("nDiscAmtx") = p_oDT(pnActiveRow)("nDiscAmtx")
        p_oDTSel(0)("xCategrCd") = p_oDT(pnActiveRow)("xCategrCd")
        p_oDTSel(0)("xStockIDx") = p_oDT(pnActiveRow)("xStockIDx")
        p_oDTSel(0)("xUnitPrce") = p_oDT(pnActiveRow)("xUnitPrce")
        p_oDTSel(0)("cSelected") = p_oDT(pnActiveRow)("cSelected")
        p_oDTSel(0)("sStockIDx") = p_oDT(pnActiveRow)("sStockIDx")
        p_oDTSel(0)("sBarCodex") = p_oDT(pnActiveRow)("sBarCodex")
        p_oDTSel(0)("sBriefDsc") = p_oDT(pnActiveRow)("sBriefDsc")
        p_oDTSel(0)("sDescript") = p_oDT(pnActiveRow)("sDescript")
        p_oDTSel(0)("sCategrCd") = p_oDT(pnActiveRow)("sCategrCd")
        p_oDTSel(0)("nUnitPrce") = p_oDT(pnActiveRow)("nUnitPrce")
        p_oDTSel(0)("nQuantity") = p_oDT(pnActiveRow)("nQuantity")
        p_oDTSel(0)("nPromBght") = p_oDT(pnActiveRow)("nPromBght")
        p_oDTSel(0)("nWeightxx") = p_oDT(pnActiveRow)("nWeightxx")
        p_oDTSel(0)("cComboMlx") = p_oDT(pnActiveRow)("cComboMlx")
    End Sub

    Private Sub setVisible()
        Me.Visible = False
        Me.TransparencyKey = Nothing
        Me.Location = New Point(507, 90)
        Me.Visible = True
    End Sub

    Private Sub clearFields()
        With p_oDT
            lblMaster00.Text = .Rows(0).Item("xPromoItm")
            lblMaster04.Text = Format(.Rows(0).Item("nDDiscRte"), "##0.00")
            lblMaster05.Text = Format(.Rows(0).Item("nDDiscAmt"), "##0.00")
        End With
    End Sub

    Private Sub loadDetail(Optional ByVal lnRow As Integer = 0)
        Dim lnCtr As Integer

        Call InitializeDataGrid()
        Call initGrid()

        With DataGridView1
            .RowCount = p_oDT.Rows.Count

            For lnCtr = 0 To p_oDT.Rows.Count - 1
                .Item(0, lnCtr).Value = lnCtr + 1
                .Item(1, lnCtr).Value = p_oDT.Rows(lnCtr)("nBaseQtyx")
                .Item(2, lnCtr).Value = p_oDT.Rows(lnCtr)("nDDiscRte")
                .Item(3, lnCtr).Value = p_oDT.Rows(lnCtr)("nDDiscAmt")
                .Item(4, lnCtr).Value = p_oDT.Rows(lnCtr)("xDsctdItm")
                .Item(5, lnCtr).Value = p_oDT.Rows(lnCtr)("xUnitPrce")
                .Item(6, lnCtr).Value = p_oDT.Rows(lnCtr)("nMaxQtyxx")
            Next

            pnActiveRow = lnRow
            .ClearSelection()
            .Rows(pnActiveRow).Selected = True

            setFieldInfo()
        End With
    End Sub

    Private Sub setFieldInfo()
        txtField01.Text = p_oDT.Rows(pnActiveRow).Item("xDsctdItm")
        txtField06.Text = p_oDT.Rows(pnActiveRow).Item("nQuantity")
        txtField07.Text = Format(p_oDT.Rows(pnActiveRow).Item("nDiscRate"), "##0.00")
        txtField08.Text = Format(p_oDT.Rows(pnActiveRow).Item("nDiscAmtx"), "##0.00")

        Dim row As DataRow

        txtField01.AutoCompleteCustomSource.Clear()

        If p_oDT(pnActiveRow)("sStockIDx") = "" And p_oDT(pnActiveRow)("xDsctdItm") <> "" Then
            For Each row In p_oInventory.SearchItems(p_oDT(pnActiveRow)("sCategrCd")).Rows
                txtField01.AutoCompleteCustomSource.Add(row.Item("sBriefDsc").ToString())
            Next

            txtField01.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtField01.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        End If

        txtField01.ReadOnly = p_oDT(pnActiveRow)("sStockIDx") <> ""
    End Sub

    Private Sub searchItem(ByVal lsBriefDsc As String)
        With p_oInventory
            If .SearchRecord(lsBriefDsc, p_oDT(pnActiveRow)("sCategrCd")) Then
                p_oDT(pnActiveRow)("sStockIDx") = .Master("sStockIDx")
                p_oDT(pnActiveRow)("sBarcodex") = .Master("sBarcodex")
                p_oDT(pnActiveRow)("sBriefDsc") = .Master("sBriefDsc")
                p_oDT(pnActiveRow)("sDescript") = .Master("sDescript")
            End If
        End With
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub PreventFlicker()
        With Me
            .SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            .SetStyle(ControlStyles.UserPaint, True)
            .SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            .UpdateStyles()
        End With
    End Sub

    Private Sub InitializeDataGrid()
        With DataGridView1
            ' Initialize basic DataGridView properties.
            .Dock = DockStyle.Fill
            .BackgroundColor = Color.LightGray
            .BorderStyle = BorderStyle.Fixed3D

            ' Set property values appropriate for read-only display and 
            ' limited interactivity. 
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .ReadOnly = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .AllowUserToResizeColumns = False
            .ColumnHeadersHeightSizeMode = _
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .AllowUserToResizeRows = False
            .RowHeadersWidthSizeMode = _
                DataGridViewRowHeadersWidthSizeMode.DisableResizing

            ' Set the selection background color for all the cells.
            .DefaultCellStyle.SelectionBackColor = Color.Empty
            .DefaultCellStyle.SelectionForeColor = Color.Black

            ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default
            ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty 'Color.White

            ' Set the background color for all rows and for alternating rows. 
            ' The value for alternating rows overrides the value for all rows. 
            .RowsDefaultCellStyle.BackColor = Color.DimGray
            .AlternatingRowsDefaultCellStyle.BackColor = Color.DarkGray

            ' Set the row and column header styles.
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.BackColor = Color.Black
            .RowHeadersDefaultCellStyle.BackColor = Color.Black

            .Font = New Font("Tahoma", 10)
            .RowTemplate.Height = 23
            .ColumnHeadersHeight = 28
        End With

        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Navy
            .ForeColor = Color.White
            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
        End With
    End Sub

    Private Sub initGrid()
        With DataGridView1
            .RowCount = 0

            'Set No of Columns
            .ColumnCount = 7

            'Set Column Headers
            .Columns(0).HeaderText = "No"
            .Columns(1).HeaderText = "REQUIRED"
            .Columns(2).HeaderText = "D. RATE"
            .Columns(3).HeaderText = "D. AMOUNT"
            .Columns(4).HeaderText = "DISC. ITEM"
            .Columns(5).HeaderText = "U-PRICE"
            .Columns(6).HeaderText = "MAX QTY"

            'Set Column Sizes
            .Columns(0).Width = 30
            .Columns(1).Width = 85
            .Columns(2).Width = 80
            .Columns(3).Width = 100
            .Columns(4).Width = 180
            .Columns(5).Width = 70
            .Columns(6).Width = 50

            .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            'Set No of Rows
            .RowCount = 1
        End With
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As System.EventArgs) Handles DataGridView1.Click
        With DataGridView1
            pnActiveRow = .CurrentCell.RowIndex

            Call setFieldInfo()
        End With
    End Sub

    Private Sub frmAddOns_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Me.Focus()
    End Sub

    Private Sub txtField_Click(sender As Object, e As System.EventArgs) Handles txtField01.Click, txtField06.Click
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 1, 6
                    loTxt.SelectAll()
            End Select
        End If
    End Sub

    Private Sub txtField06_LostFocus(sender As Object, e As System.EventArgs) Handles txtField06.LostFocus
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 6
                    If Not IsNumeric(loTxt.Text) Then
                        loTxt.Text = p_oDT(pnActiveRow)("nQuantity")
                        Exit Sub
                    End If

                    p_oDT(pnActiveRow)("nQuantity") = loTxt.Text
            End Select
        End If
    End Sub
End Class