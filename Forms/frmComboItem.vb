Imports System.Windows.Forms
Imports System.Drawing

Public Class frmComboItem
    Private WithEvents p_oRecord As clsInventory

    Private pnLoadx As Integer
    Private poControl As Control

    Private pbUpdate As Boolean
    Private pnAcRow As Integer
    Private pnIndex As Integer
    Private p_nEditMode As Integer

    ReadOnly Property Save
        Get
            Return pbUpdate
        End Get
    End Property

    WriteOnly Property EditMode
        Set(value)
            p_nEditMode = value
        End Set
    End Property

    Private Sub setFieldInfo()
        With DataGridView1
            txtField03.Text = .Item(3, pnAcRow).Value
            txtField05.Text = .Item(1, pnAcRow).Value
            txtField06.Text = .Item(2, pnAcRow).Value

            Dim loTxt As RadioButton
            Dim lbEdit As Boolean

            loTxt = CType(FindRadioButton(Me, "optButton" & Format(CInt(p_oRecord.ComboItem(pnAcRow, "cStatusxx")), "00")), RadioButton)
            loTxt.Checked = True

            lbEdit = CInt(p_oRecord.ComboItem(pnAcRow, "cStatusxx")) <> 1
            optButton00.Enabled = lbEdit
            optButton01.Enabled = Not lbEdit
            optButton02.Enabled = lbEdit

            txtField05.Enabled = Not lbEdit
            txtField06.Enabled = Not lbEdit
            txtField03.Enabled = Not lbEdit

            If Not lbEdit Then
                txtField05.Focus()
            Else
                txtField03.Focus()
            End If
        End With
    End Sub

    Private Sub clearFields()
        txtField03.Text = ""
        txtField05.Text = ""
        txtField06.Text = ""
    End Sub

    Private Sub loadDetail()
        Dim lnCtr As Integer

        initGrid()
        With DataGridView1
            .RowCount = p_oRecord.ComboItemCount


            pnAcRow = .RowCount - 1
            For lnCtr = 0 To pnAcRow
                .Item(0, lnCtr).Value = lnCtr + 1
                .Item(1, lnCtr).Value = p_oRecord.ComboItem(lnCtr, "sBarCodex")
                .Item(2, lnCtr).Value = p_oRecord.ComboItem(lnCtr, "sDescript")
                .Item(3, lnCtr).Value = p_oRecord.ComboItem(lnCtr, "nQuantity")
            Next

            .ClearSelection()
            .Rows(pnAcRow).Selected = True

            setFieldInfo()
        End With
    End Sub

    'design grid
    Private Sub initGrid()
        With DataGridView1
            .RowCount = 0

            'Set No of Columns
            .ColumnCount = 4

            'Set Column Headers
            .Columns(0).HeaderText = "No"
            .Columns(1).HeaderText = "Barcode"
            .Columns(2).HeaderText = "Description"
            .Columns(3).HeaderText = "Qty"

            'Set Column Sizes
            .Columns(0).Width = 35
            .Columns(1).Width = 160
            .Columns(2).Width = 273
            .Columns(3).Width = 55

            .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

            .Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            'Set No of Rows
            .RowCount = 1
        End With

        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Navy
            .ForeColor = Color.White
            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
        End With
    End Sub

    Private Sub frmComboItem_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        If pnLoadx = 1 Then
            pnLoadx = 2
        End If

        txtField05.Focus()
    End Sub

    Private Sub frmComboItem_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Return, Keys.Up, Keys.Down
                Select Case e.KeyCode
                    Case Keys.Return, Keys.Down
                        SetNextFocus()
                    Case Keys.Up
                        SetPreviousFocus()
                End Select
        End Select
    End Sub

    Private Sub frmComboItem_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        If pnLoadx = 0 Then
            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
            Call grpCancelHandler(Me, GetType(TextBox), "txtField", "Validating", AddressOf txtField_Validating)
            Call grpKeyHandler(Me, GetType(TextBox), "txtField", "KeyDown", AddressOf txtField_KeyDown)

            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            Call grpEventHandler(Me, GetType(RadioButton), "optButton", "Click", AddressOf optButton_Click)

            initGrid()
            clearFields()

            loadDetail()
            initButton()

            pnLoadx = 1
        End If
    End Sub

    Private Sub txtField_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 3
                    If Not IsNumeric(loTxt.Text) Then loTxt.Text = 0

                    p_oRecord.ComboItem(pnAcRow, loIndex) = loTxt.Text
                Case 5, 6
                    p_oRecord.ComboItem(pnAcRow, loIndex) = loTxt.Text
            End Select
        End If
    End Sub

    Private Sub txtField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 1
            End Select
        End If

        pnIndex = loIndex
        poControl = loTxt

        loTxt.BackColor = Color.Azure
        loTxt.SelectAll()
    End Sub

    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 5, 6
                        p_oRecord.SearchComboItem(pnAcRow, loIndex, loTxt.Text)
                        txtField03.Focus()
                End Select
            End If
        End If
    End Sub

    Private Sub txtField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 1
            End Select
        End If

        loTxt.BackColor = SystemColors.Window
        poControl = Nothing
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        With p_oRecord
            Select Case lnIndex
                Case 0 'add detail
                    p_oRecord.addComboItem()
                    loadDetail()
                Case 1 'save
                    pbUpdate = True
                    Me.Hide()
                Case 2 'search
                    If pnIndex = 5 Or pnIndex = 6 Then
                        p_oRecord.SearchComboItem(pnAcRow, pnIndex, "")
                        txtField03.Focus()
                    End If
                Case 3 'close
                    pbUpdate = False
                    Me.Hide()
            End Select
        End With
endProc:
        Exit Sub
    End Sub

    Public Sub New(oRecord As clsInventory)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        p_oRecord = oRecord
    End Sub

    Private Sub DataGridView1_Click(sender As Object, e As System.EventArgs) Handles DataGridView1.Click
        pnAcRow = DataGridView1.CurrentRow.Index

        setFieldInfo()
    End Sub

    Private Sub p_oRecord_ComboItemRetrieved(Row As Integer, Index As Integer, Value As Object) Handles p_oRecord.ComboItemRetrieved
        With DataGridView1
            Select Case Index
                Case 3
                    .Rows(pnAcRow).Cells(3).Value = Value
                    txtField03.Text = Value
                Case 5
                    '.Item(1, Row).Value = Value
                    .Rows(pnAcRow).Cells(1).Value = Value
                    txtField05.Text = Value
                Case 6
                    '.Item(2, Row).Value = Value
                    .Rows(pnAcRow).Cells(2).Value = Value
                    txtField06.Text = Value
                Case 9
            End Select
        End With
    End Sub

    Private Sub initButton()
        Dim lbShow As Boolean

        lbShow = p_nEditMode = 1 Or p_nEditMode = 2
        Panel2.Enabled = lbShow
        cmdButton00.Visible = lbShow
        cmdButton01.Visible = lbShow
        cmdButton02.Visible = lbShow
    End Sub

    Private Sub optButton_Click(sender As Object, e As System.EventArgs)
        Dim loChk As RadioButton
        loChk = CType(sender, System.Windows.Forms.RadioButton)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        p_oRecord.ComboItem(pnAcRow, 9) = lnIndex
    End Sub
End Class