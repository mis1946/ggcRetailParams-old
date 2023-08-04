'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Inventory Maintenance Object
'
' Copyright 2016 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Kalyptus [ 10/10/2016 04:02 pm ]
'      Started creating this object.
'  Note: Possible values of combo meal
'       0 -> Original
'       1 -> Added
'       2 -> Remove
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ggcAppDriver

Public Class clsInventory
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTComb As DataTable
    Private p_oOthersx As New Others

    Private p_nEditMode As xeEditMode
    Private p_sParent As String

    Private Const p_sMasTable1 As String = "Inventory"
    Private Const p_sMasTable2 As String = "Inventory_Master"
    Private Const p_sMasTable3 As String = "Combo_Meals"

    Private Const p_sMsgHeadr As String = "Stock/Inventory Maintenance"

    Public Event MasterRetrieved(ByVal Index As Integer, _
                                  ByVal Value As Object)

    Public Event ComboItemRetrieved(ByVal Row As Integer, _
                                    ByVal Index As Integer, _
                                    ByVal Value As Object)

    Public ReadOnly Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
    End Property

    Public Property Master(ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sCategrNm) = "" Then
                            getCategory(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCategrNm
                    Case 81
                        If Trim(IFNull(p_oDTMstr(0).Item(5))) <> "" And Trim(p_oOthersx.sSizeName) = "" Then
                            getSize(5, 81, p_oDTMstr(0).Item(5), True, False)
                        End If
                        Return p_oOthersx.sSizeName
                    Case 82
                        If Trim(IFNull(p_oDTMstr(0).Item(6))) <> "" And Trim(p_oOthersx.sMeasurNm) = "" Then
                            getMeasure(6, 82, p_oDTMstr(0).Item(6), True, False)
                        End If
                        Return p_oOthersx.sMeasurNm
                    Case 83
                        If Trim(IFNull(p_oDTMstr(0).Item(15))) <> "" And Trim(p_oOthersx.sSectnNme) = "" Then
                            getSection(15, 83, p_oDTMstr(0).Item(15), True, False)
                        End If
                        Return p_oOthersx.sSectnNme
                    Case 84
                        If Trim(IFNull(p_oDTMstr(0).Item(16))) <> "" And Trim(p_oOthersx.sBinNamex) = "" Then
                            getBin(16, 84, p_oDTMstr(0).Item(16), True, False)
                        End If
                        Return p_oOthersx.sBinNamex
                    Case 85
                        If Trim(IFNull(p_oDTMstr(0).Item(17))) <> "" And Trim(p_oOthersx.sInvTypex) = "" Then
                            getInvType(7, 85, p_oDTMstr(0).Item(7), True, False)
                        End If
                        Return p_oOthersx.sInvTypex
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 80 ' sCategrNm
                        getCategory(4, 80, value, False, False)
                    Case 81 ' sSizeName
                        getSize(5, 81, value, False, False)
                    Case 82 ' sMeasurNm
                        getMeasure(6, 82, value, False, False)
                    Case 83 ' sSectnNme
                        getSection(15, 83, value, False, False)
                    Case 84 ' sBinNamex
                        getBin(16, 84, value, False, False)
                    Case 85 ' sInvTypex
                        getInvType(7, 85, value, False, False)

                        'Please give validation for the assignment of values for this fields...
                    Case 8, 9  'nUnitPrce, nSelPrice
                        p_oDTMstr(0).Item(Index) = value
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 10, 11, 12, 13  'nDiscLev1, nDiscLev2, nDiscLev3, nDealrDsc
                        If Not IsNumeric(value) Then value = 0

                        p_oDTMstr(0).Item(Index) = value
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 17 'dBegInvxx
                        If (p_oDTMstr(0).Item("sStockID2") = "" Or p_sParent <> "") And IsDate(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 18 'nBegQtyxx
                        If (p_oDTMstr(0).Item("sStockID2") = "" Or p_sParent <> "") And IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value

                            If p_oDTMstr(0).Item("sStockID2") = "" Then
                                p_oDTMstr(0).Item("nQtyOnHnd") = value
                                RaiseEvent MasterRetrieved(19, p_oDTMstr(0).Item(Index))
                            End If

                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 19 ' nQtyOnHnd
                        If (p_oDTMstr(0).Item("sStockID2") = "" Or p_sParent <> "") And IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 20, 21 'nMinLevel, nMaxLevel
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))

                    Case 22, 23, 24 ' nResvOrdr, nBackOrdr, nAveMonSl
                        If (p_oDTMstr(0).Item("sStockID2") = "" Or p_sParent <> "") And IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If

                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property Master(ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case "scategrnm" ' 80
                        If Trim(IFNull(p_oDTMstr(0).Item(4))) <> "" And Trim(p_oOthersx.sCategrNm) = "" Then
                            getCategory(4, 80, p_oDTMstr(0).Item(4), True, False)
                        End If
                        Return p_oOthersx.sCategrNm
                    Case "ssizename" ' 81
                        If Trim(IFNull(p_oDTMstr(0).Item(5))) <> "" And Trim(p_oOthersx.sSizeName) = "" Then
                            getSize(5, 81, p_oDTMstr(0).Item(5), True, False)
                        End If
                        Return p_oOthersx.sSizeName
                    Case "smeasurnm" ' 82
                        If Trim(IFNull(p_oDTMstr(0).Item(6))) <> "" And Trim(p_oOthersx.sMeasurNm) = "" Then
                            getMeasure(6, 82, p_oDTMstr(0).Item(6), True, False)
                        End If
                        Return p_oOthersx.sMeasurNm
                    Case "ssectnnme" ' 83
                        If Trim(IFNull(p_oDTMstr(0).Item(15))) <> "" And Trim(p_oOthersx.sSectnNme) = "" Then
                            getSection(15, 83, p_oDTMstr(0).Item(15), True, False)
                        End If
                        Return p_oOthersx.sSectnNme
                    Case "sbinnamex" '84 
                        If Trim(IFNull(p_oDTMstr(0).Item(16))) <> "" And Trim(p_oOthersx.sBinNamex) = "" Then
                            getBin(16, 84, p_oDTMstr(0).Item(16), True, False)
                        End If
                        Return p_oOthersx.sBinNamex
                    Case "sinvtypex" '85 
                        If Trim(IFNull(p_oDTMstr(0).Item(17))) <> "" And Trim(p_oOthersx.sInvTypex) = "" Then
                            getInvType(7, 85, p_oDTMstr(0).Item(17), True, False)
                        End If
                        Return p_oOthersx.sInvTypex
                    Case Else
                        Return p_oDTMstr(0).Item(Index)
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    Case "scategrnm" ' 80  
                        getCategory(4, 80, value, False, False)
                    Case "ssizename" ' 81  
                        getSize(5, 81, value, False, False)
                    Case "smeasurnm" ' 82  
                        getMeasure(6, 82, value, False, False)
                    Case "ssectnnme" ' 83  
                        getSection(15, 83, value, False, False)
                    Case "sbinnamex" '84 
                        getBin(16, 84, value, False, False)
                    Case "sinvtypex" '85 
                        getInvType(7, 85, value, False, False)
                    Case "nunitprce", "nselprice", "ndisclev1", "ndisclev2", "ndisclev3", "ndealrdsc"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "nbegqtyxx", "nqtyonhnd", "nminlevel", "nmaxlevel", "nresvordr", "nbackordr", "navemonsl"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case "dbeginvxx"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                    Case Else
                        p_oDTMstr(0).Item(Index) = value
                End Select
            End If
        End Set
    End Property

    '#COMBO
    Public Property ComboItem(ByVal Row As Integer, ByVal Index As Integer) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    Case 0, 1, 3, 4, 7 To 9
                        Return p_oDTComb(Row).Item(Index)
                    Case 2
                        Return p_oDTComb(Row).Item(Index)
                    Case 5, 6
                        If Trim(p_oDTComb(Row).Item("sStockIDx")) <> "" And Trim(Trim(p_oDTComb(Row).Item(Index))) = "" Then
                            getInvType(7, 85, p_oDTComb(0).Item(7), True, False)
                        End If

                        Return p_oDTComb(Row).Item(Index)
                    Case Else
                        Return vbEmpty
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case Index
                    'This fields are not allowed to be assigned outside this class
                    Case 0, 1, 4, 7, 8  'Case "scomboidx", "nentrynox", "dmodified", "xcomboidx", "xquantity"
                    Case 2              'Case "sstockidx"
                        If Trim(p_oDTComb(Row).Item(Index)) = Trim(value) Then
                            p_oDTComb(Row).Item(Index) = value
                            p_oDTComb(Row).Item(5) = ""
                            p_oDTComb(Row).Item(6) = ""
                        End If
                    Case 5, 6        'Case "sbarcodex", "sdescript"

                    Case 3              'Case "nquantity"
                        If IsNumeric(value) Then
                            If value > 0 Then
                                p_oDTComb(Row).Item(Index) = value
                            End If
                        End If
                        RaiseEvent ComboItemRetrieved(Row, Index, p_oDTComb(Row).Item(Index))
                    Case 9              'Case "cstatusxx"
                        If value = "0" Or value = "2" Then
                            p_oDTComb(Row).Item(Index) = value
                        End If
                        RaiseEvent ComboItemRetrieved(Row, Index, p_oDTComb(Row).Item(Index))
                End Select
            End If
        End Set
    End Property

    'Property Master(String)
    Public Property ComboItem(ByVal Row As Integer, ByVal Index As String) As Object
        Get
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Index = LCase(Index)
                Select Case Index
                    Case "scomboidx", "nentrynox", "sstockidx", "nquantity", "dmodified", _
                         "sbarcodex", "sdescript", "xcomboidx", "cstatusxx", "xquantity"
                        Return p_oDTComb(Row).Item(Index)
                    Case Else
                        Return vbEmpty
                End Select
            Else
                Return vbEmpty
            End If
        End Get

        Set(ByVal value As Object)
            If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                Select Case LCase(Index)
                    'This fields are not allowed to be assigned outside this class
                    Case "scomboidx", "nentrynox", "dmodified", "xcomboidx", "xquantity"
                    Case "sstockidx"
                    Case "sbarcodex", "sdescript"
                        ComboItem(Row, p_oDTComb.Columns(Index).Ordinal) = value
                    Case "nquantity"
                        ComboItem(Row, p_oDTComb.Columns(Index).Ordinal) = value
                    Case "cstatusxx"
                        ComboItem(Row, p_oDTComb.Columns(Index).Ordinal) = value
                End Select
            End If
        End Set
    End Property

    Public ReadOnly Property ComboItemCount() As Integer
        Get
            Return p_oDTComb.Rows.Count
        End Get
    End Property

    'Property EditMode()
    Public ReadOnly Property EditMode() As xeEditMode
        Get
            Return p_nEditMode
        End Get
    End Property

    Public Property Parent() As String
        Get
            Return p_sParent
        End Get
        Set(ByVal value As String)
            p_sParent = value
        End Set
    End Property

    Public Function UpdateRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Public Function CancelUpdate() As Boolean
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    'Public Function NewRecord()
    Public Function NewRecord() As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "0=1")
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())

        Call initMaster()
        Call initOthers()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    'Public Function OpenRecord(String)
    Public Function OpenRecord(ByVal fsRecdIDxx As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQ_Master, "a.sStockIDx = " & strParm(fsRecdIDxx))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        Call initOthers()

        If p_oDTMstr(0).Item("sStockID2") = "" Then
            p_nEditMode = xeEditMode.MODE_ADDNEW
        Else
            p_nEditMode = xeEditMode.MODE_READY

            If p_oDTMstr(0).Item("cComboMlx") = "1" Then loadComboTable()
        End If

        Return True
    End Function

    Public Function SearchItems(ByVal fsCategID As String) As DataTable
        Dim loDT As DataTable

        loDT = p_oApp.ExecuteQuery(AddCondition(getSQ_Browse, "sCategrID = " & strParm(fsCategID)))

        Return loDT
    End Function

    Public Function SearchRecord(ByVal fsValue As String _
                                , ByVal fsCategID As String) As Boolean

        Dim lsSQL As String

        'Make sure that the parameter for search has value
        fsValue = Trim(fsValue)

        'Initialize SQL filter
        lsSQL = getSQ_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String

        lsFilter = "sBriefDsc LIKE " & strParm(fsValue & "%")

        If fsCategID <> "" Then lsSQL = AddCondition(lsSQL, "sCategrID = " & strParm(fsCategID))

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sBarCodex»sDescript" _
                                        , "Barcode»Description", _
                                        , "sBarCodex»sDescript" _
                                        , 1)

        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenRecord(loDta.Item("sStockIDx"))
        End If
    End Function

    'Public Function SearchRecord(String, Boolean, Boolean=False)
    Public Function SearchRecord(ByVal fsValue As String _
                                , Optional ByVal fbByCode As Boolean = False ) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sBarCodex") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sBriefDsc") Then Return True
            End If
        End If

        'Make sure that the parameter for search has value
        fsValue = Trim(fsValue)

        'Initialize SQL filter
        lsSQL = getSQ_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String

        If fbByCode Then
            lsFilter = "sBarCodex = " & strParm(fsValue)
        Else
            lsFilter = "sBriefDsc LIKE " & strParm(fsValue & "%")
        End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , False _
                                        , lsFilter _
                                        , "sBarCodex»sBriefDsc»sDescript" _
                                        , "Barcode»sBriefDsc»Description",
                                        , "sBarCodex»sBriefDsc»sDescript" _
                                        , IIf(fbByCode, 0, 1))



        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenRecord(loDta.Item("sStockIDx"))
        End If
    End Function

    'Public Function SearchMaster
    Public Sub SearchMaster(ByVal fnIndex As Integer, _
                                 ByVal fsValue As String)
        Select Case fnIndex
            'iMac 2017.01.31
            'we may allow searching records if value is empty,
            'for ease entry of records
            Case 80
                'If fsValue <> "" Then
                getCategory(4, fnIndex, fsValue, False, True)
                'End If
            Case 81
                'If fsValue <> "" Then
                getSize(5, fnIndex, fsValue, False, True)
                'End If
            Case 82
                'If fsValue <> "" Then
                getMeasure(6, fnIndex, fsValue, False, True)
                'End If
            Case 83
                'If fsValue <> "" Then
                getSection(15, fnIndex, fsValue, False, True)
                'End If
            Case 84
                'If fsValue <> "" Then
                getBin(16, fnIndex, fsValue, False, True)
                'End If
            Case 85
                'If fsValue <> "" Then
                getInvType(7, fnIndex, fsValue, False, True)
                'End If
        End Select
    End Sub

    'Public Function SearchMaster
    Public Sub SearchComboItem(ByVal fnRowxx As Integer, _
                               ByVal fnIndex As Integer, _
                               ByVal fsValue As String)
        Select Case fnIndex
            Case 5, 6
                Call getComboItem(fnRowxx, 2, fnIndex, fsValue, False, True)
        End Select
    End Sub


    'Public Function SaveTransaction
    Public Function SaveRecord() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        'Try
        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String = ""

        If p_sParent = "" Then p_oApp.BeginTransaction()

        Dim lsExcluded As String
        lsExcluded = "sBranchCd»sSectnIDx»sBinIDxxx»dBegInvxx»nBegQtyxx»nQtyOnHnd»nMinLevel»nMaxLevel»nResvOrdr»nBackOrdr»nAveMonSl»sStockID1»sStockID2"

        'Save Inventory table 
        If p_oDTMstr(0).Item("sStockID1") = "" Then
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable1, , p_oApp.UserID, p_oApp.SysDate, lsExcluded)
        Else
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable1, "sStockIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")), p_oApp.UserID, Format(p_oApp.SysDate, xsDATE_SHORT), lsExcluded)
        End If

        If lsSQL <> "" Then
            p_oApp.Execute(lsSQL, p_sMasTable1)
        End If

        lsExcluded = "sBarCodex»sDescript»sBriefDsc»sCategrID»sSizeIDxx»sMeasurID»sInvTypID»nUnitPrce»nSelPrice»nDiscLev1»nDiscLev2»nDiscLev3»nDealrDsc»cComboMlx»cWthPromo»sStockID1»sStockID2»sImgePath"

        'Save Inventory_Master table 
        If p_oDTMstr(0).Item("sStockID2") = "" Then
            p_oDTMstr(0).Item("sBranchCd") = p_oApp.BranchCode
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable2, , p_oApp.UserID, p_oApp.SysDate, lsExcluded)
        Else
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable2, "sStockIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")) & " AND sBranchCD = " & strParm(p_oDTMstr(0).Item("sBranchCD")), p_oApp.UserID, Format(p_oApp.SysDate, xsDATE_SHORT), lsExcluded)
        End If

        If lsSQL <> "" Then
            p_oApp.Execute(lsSQL, p_sMasTable2)
        End If

        'Save combo meals detail
        If p_oDTMstr(0).Item("cComboMlx") = 1 Then
            Dim lnCtr As Integer
            Dim lnValidCtr As Integer
            lnValidCtr = 0

            For lnCtr = 0 To p_oDTComb.Rows.Count - 1
                lsSQL = ""
                If p_oDTComb(lnCtr).Item("sStockIDx") <> "" Then
                    Select Case p_oDTComb(lnCtr).Item("cStatusxx")
                        Case "0"  'Original 
                            ' check changes in nEntryNox value 
                            If p_oDTComb(lnCtr).Item("nEntryNox") <> lnValidCtr Then
                                lsSQL = ", nEntryNox = " & lnValidCtr
                            End If

                            ' check changes in nQuantity value 
                            If p_oDTComb(lnCtr).Item("nQuantity") <> p_oDTComb(lnCtr).Item("xQuantity") Then
                                lsSQL = ", nQuantity = " & p_oDTComb(lnCtr).Item("nQuantity")
                            End If

                            'Create the UPDATE SQL statement if there is/are changes detected from above
                            If lsSQL <> "" Then
                                lsSQL = "UPDATE " & p_sMasTable3 & _
                                       " SET " & Mid(lsSQL, 2) & _
                                       " WHERE sComboIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")) & _
                                         " AND sStockIDx = " & strParm(p_oDTComb(lnCtr).Item("sStockIDx"))
                            End If
                        Case "1" 'Add
                            'Create the INSERT statement
                            lsSQL = "INSERT INTO " & p_sMasTable3 & _
                                   " SET sComboIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")) & _
                                      ", nEntryNox = " & lnValidCtr & _
                                      ", sStockIDx = " & strParm(p_oDTComb(lnCtr).Item("sStockIDx")) & _
                                      ", nQuantity = " & p_oDTComb(lnCtr).Item("nQuantity") & _
                                      ", dModified = " & dateParm(p_oApp.getSysDate)
                        Case "2" 'Remove
                            'Create the DELETE statement
                            lsSQL = "DELETE FROM " & p_sMasTable3 & _
                                   " WHERE sComboIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")) & _
                                     " AND sStockIDx = " & strParm(p_oDTComb(lnCtr).Item("sStockIDx"))
                    End Select

                End If

                If lsSQL <> "" Then
                    p_oApp.Execute(lsSQL, p_sMasTable3)
                End If
            Next
        End If

        If p_sParent = "" Then p_oApp.CommitTransaction()

        p_nEditMode = xeEditMode.MODE_READY

        Return True

        'Catch

    End Function

    'Public Function CancelTransaction
    Public Function CancelRecord() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        Dim lsSQL As String

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("cRecdStat") = "3"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable1, "sStockIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx")))
        p_oApp.Execute(lsSQL, p_sMasTable1, p_oApp.BranchCode)

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getBin(ByVal fnColIdx As Integer _
                     , ByVal fnColDsc As Integer _
                     , ByVal fsValue As String _
                     , ByVal fbIsCode As Boolean _
                     , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sBinNamex <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sBinNamex And fsValue <> "" Then Exit Sub
        End If

        Dim loBin As clsBin
        loBin = New clsBin(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loBin.SearchBin(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBinNamex = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sBinIDxxx")
                p_oOthersx.sBinNamex = loRow.Item("sBinNamex")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBinNamex)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loBin.GetBin(fsValue, True)
            Else
                loDta = loBin.GetBin(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBinNamex = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sBinIDxxx")
                p_oOthersx.sBinNamex = loDta(0).Item("sBinNamex")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBinNamex)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getSection(ByVal fnColIdx As Integer _
                         , ByVal fnColDsc As Integer _
                         , ByVal fsValue As String _
                         , ByVal fbIsCode As Boolean _
                         , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sSectnNme <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sSectnNme And fsValue <> "" Then Exit Sub
        End If

        Dim loSection As clsSection
        loSection = New clsSection(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loSection.SearchSection(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sSectnNme = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sSectnIDx")
                p_oOthersx.sSectnNme = loRow.Item("sSectnNme")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sSectnNme)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loSection.GetSection(fsValue, True)
            Else
                loDta = loSection.GetSection(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sSectnNme = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sSectnIDx")
                p_oOthersx.sSectnNme = loDta(0).Item("sSectnNme")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sSectnNme)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getSize(ByVal fnColIdx As Integer _
                      , ByVal fnColDsc As Integer _
                      , ByVal fsValue As String _
                      , ByVal fbIsCode As Boolean _
                      , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sSizeName <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sSizeName And fsValue <> "" Then Exit Sub
        End If

        Dim loSize As clsSize
        loSize = New clsSize(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loSize.SearchSize(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sSizeName = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sSizeIDxx")
                p_oOthersx.sSizeName = loRow.Item("sSizeName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sSizeName)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loSize.GetSize(fsValue, True)
            Else
                loDta = loSize.GetSize(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sSizeName = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sSizeIDxx")
                p_oOthersx.sSizeName = loDta(0).Item("sSizeName")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sSizeName)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getMeasure(ByVal fnColIdx As Integer _
                         , ByVal fnColDsc As Integer _
                         , ByVal fsValue As String _
                         , ByVal fbIsCode As Boolean _
                         , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sMeasurNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sMeasurNm And fsValue <> "" Then Exit Sub
        End If

        Dim loMeasure As clsMeasure
        loMeasure = New clsMeasure(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loMeasure.SearchMeasure(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sMeasurNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sMeasurID")
                p_oOthersx.sMeasurNm = loRow.Item("sMeasurNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sMeasurNm)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loMeasure.GetMeasure(fsValue, True)
            Else
                loDta = loMeasure.GetMeasure(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sMeasurNm = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sMeasurID")
                p_oOthersx.sMeasurNm = loDta(0).Item("sMeasurNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sMeasurNm)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getCategory(ByVal fnColIdx As Integer _
                          , ByVal fnColDsc As Integer _
                          , ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sCategrNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sCategrNm And fsValue <> "" Then Exit Sub
        End If

        Dim loCategory As clsProductCategory
        loCategory = New clsProductCategory(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loCategory.SearchCategory(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sCategrNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sCategrCd")
                p_oOthersx.sCategrNm = loRow.Item("sDescript")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCategrNm)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loCategory.GetCategory(fsValue, True)
            Else
                loDta = loCategory.GetCategory(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sCategrNm = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sCategrCd")
                p_oOthersx.sCategrNm = loDta(0).Item("sDescript")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sCategrNm)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getInvType(ByVal fnColIdx As Integer _
                          , ByVal fnColDsc As Integer _
                          , ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sInvTypex <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sInvTypex And fsValue <> "" Then Exit Sub
        End If

        Dim loInventory As clsInvType
        loInventory = New clsInvType(p_oApp)

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = loInventory.SearchInvType(fsValue, False)

            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sInvTypex = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sInvTypCd")
                p_oOthersx.sInvTypex = loRow.Item("sDescript")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sInvTypex)
            Exit Sub
        End If

        Dim loDta As DataTable

        If fsValue <> "" Then
            If fbIsCode Then
                loDta = loInventory.GetInvType(fsValue, True)
            Else
                loDta = loInventory.GetInvType(fsValue, False)
            End If

            If loDta.Rows.Count = 0 Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sInvTypex = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sInvTypCd")
                p_oOthersx.sInvTypex = loDta(0).Item("sDescript")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sInvTypex)
        End If
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getComboItem(ByVal fnItemRw As Integer _
                           , ByVal fnColIdx As Integer _
                           , ByVal fnColDsc As Integer _
                           , ByVal fsValue As String _
                           , ByVal fbIsCode As Boolean _
                           , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTComb(fnItemRw).Item(fnColIdx) And fsValue <> "" And p_oDTComb(fnItemRw).Item(fnColDsc) <> "" Then Exit Sub
        Else
            If fsValue = p_oDTComb(fnItemRw).Item(fnColDsc) And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBarCodex" & _
                       ", a.sDescript" & _
                       ", a.sStockIDx" & _
               " FROM `Inventory` a" & _
               " WHERE a.cComboMlx = '0'" & _
        IIf(fbIsCode = False, " AND a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBarCodex»sDescript»sStockIDx" _
                                             , "Barcode»Description»Stock ID", _
                                             , "a.sBarCodex»a.sDescript»a.sStockIDx" _
                                             , IIf(fbIsCode, 2, IIf(fnColDsc = 5, 0, 1)))

            If IsNothing(loRow) Then
                p_oDTComb(fnItemRw).Item(fnColIdx) = ""
                p_oDTComb(fnItemRw).Item(5) = ""
                p_oDTComb(fnItemRw).Item(6) = ""
            Else
                p_oDTComb(fnItemRw).Item(fnColIdx) = loRow.Item("sStockIDx")
                p_oDTComb(fnItemRw).Item(5) = loRow.Item("sBarCodex")
                p_oDTComb(fnItemRw).Item(6) = loRow.Item("sDescript")
            End If

            RaiseEvent ComboItemRetrieved(fnItemRw, 5, p_oDTComb(fnItemRw).Item(5))
            RaiseEvent ComboItemRetrieved(fnItemRw, 6, p_oDTComb(fnItemRw).Item(6))
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                If fbIsCode Then
                    lsSQL = AddCondition(lsSQL, "a.sStockIDx = " & strParm(fsValue))
                Else
                    If fnColDsc = 5 Then
                        lsSQL = AddCondition(lsSQL, "a.sBarCodex = " & strParm(fsValue))
                    Else
                        lsSQL = AddCondition(lsSQL, "a.sDescript = " & strParm(fsValue))
                    End If
                End If
            End If

            Dim loDta As DataTable
            loDta = p_oApp.ExecuteQuery(lsSQL)

            If loDta.Rows.Count = 0 Then
                p_oDTComb(fnItemRw).Item(fnColIdx) = ""
                p_oDTComb(fnItemRw).Item(5) = ""
                p_oDTComb(fnItemRw).Item(6) = ""
            ElseIf loDta.Rows.Count = 1 Then
                p_oDTComb(fnItemRw).Item(fnColIdx) = loDta(0).Item("sStockIDx")
                p_oDTComb(fnItemRw).Item(5) = loDta(0).Item("sBarCodex")
                p_oDTComb(fnItemRw).Item(6) = loDta(0).Item("sDescript")
            End If

            RaiseEvent ComboItemRetrieved(fnItemRw, 5, p_oDTComb(fnItemRw).Item(5))
            RaiseEvent ComboItemRetrieved(fnItemRw, 6, p_oDTComb(fnItemRw).Item(6))
        End If
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "sstockidx"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable1, "sStockIDx", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "sbarcodex"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable1, "sBarcodex", True, p_oApp.Connection)
                Case "dmodified", "smodified"
                Case "sbranchcd"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.BranchCode
                Case "dbeginvxx"
                    p_oDTMstr(0).Item(lnCtr) = p_oApp.SysDate
                Case "crecdstat"
                    p_oDTMstr(0).Item(lnCtr) = "1"
                Case "ccombomlx", "cwthpromo"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "nunitprce", "nselprice", "ndisclev1", "ndisclev2", "ndisclev3", "ndealrdsc"
                    p_oDTMstr(0).Item(lnCtr) = 0.0
                Case "nbegqtyxx", "nqtyonhnd", "nminlevel", "nmaxlevel", "nresvordr", "nbackordr", "navemonsl"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub initOthers()
        p_oOthersx.sCategrNm = ""
        p_oOthersx.sSizeName = ""
        p_oOthersx.sMeasurNm = ""
        p_oOthersx.sSectnNme = ""
        p_oOthersx.sBinNamex = ""
        p_oOthersx.sInvTypex = ""
    End Sub

    Public Sub showComboMeals()
        If p_nEditMode = xeEditMode.MODE_UNKNOWN Then Exit Sub

        If p_oDTMstr(0).Item("cComboMlx") = "1" Then

            If p_oDTMstr(0).Item("sStockID1") = "" Then
                Call createComboTable()
            Else
                Call loadComboTable()
            End If

            If p_oDTComb.Rows.Count = 0 Then
                p_oDTComb.Rows.Add()
                Call initComboTable()
            End If

            Dim loDTTemp As DataTable
            loDTTemp = Nothing
            loDTTemp = p_oDTComb.Clone
            For lnCtr = 0 To p_oDTComb.Rows.Count - 1
                loDTTemp.ImportRow(p_oDTComb(lnCtr))
            Next

            'Show form
            Dim loForm As frmComboItem

            loForm = New frmComboItem(Me)
            loForm.EditMode = p_nEditMode

            'pass the combo datable for rollback purpose
            loForm.ShowDialog()

            'loForm.Save = false means rollback/disregard update on combo
            'so set the combo datatable to the original form that
            'we have passed on the form.
            If Not loForm.Save Then
                If p_nEditMode <> xeEditMode.MODE_UNKNOWN Then
                    p_oDTComb = loDTTemp 'rollback to original datatable
                End If
            End If
        Else
            MsgBox("This is not a combo meal!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
        End If
    End Sub

    Public Sub addComboItem()
        Dim lnCtr As Integer

        For lnCtr = 0 To p_oDTComb.Rows.Count - 1
            If p_oDTComb(lnCtr).Item("sComboIDx") = "" Then
                MsgBox("There is a combo item that needs to be filled up first!" & vbCrLf & _
                       "Please fill this up before adding a new item...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Exit Sub
            End If
        Next

        p_oDTComb.Rows.Add()
        Call initComboTable()
    End Sub

    Private Sub createComboTable()
        p_oDTComb = New DataTable
        With p_oDTComb
            .Columns.Add("sComboIDx", GetType(String)).MaxLength = 12
            .Columns.Add("nEntryNox", GetType(Integer))
            .Columns.Add("sStockIDx", GetType(String)).MaxLength = 12
            .Columns.Add("nQuantity", GetType(Integer))
            .Columns.Add("dModified", GetType(Date))
            .Columns.Add("sBarCodex", GetType(String)).MaxLength = 12
            .Columns.Add("sDescript", GetType(String)).MaxLength = 64
            .Columns.Add("xComboIDx", GetType(String)).MaxLength = 12
            .Columns.Add("xQuantity", GetType(Integer))
            .Columns.Add("cStatusxx", GetType(String)).MaxLength = 1
        End With
    End Sub

    Private Sub initComboTable()
        With p_oDTComb
            .Rows(.Rows.Count - 1)("sComboIDx") = p_oDTMstr(0).Item("sStockIDx")
            .Rows(.Rows.Count - 1)("nEntryNox") = .Rows.Count - 1
            .Rows(.Rows.Count - 1)("sStockIDx") = ""
            .Rows(.Rows.Count - 1)("nQuantity") = 0.0
            .Rows(.Rows.Count - 1)("dModified") = p_oApp.SysDate
            .Rows(.Rows.Count - 1)("sBarCodex") = ""
            .Rows(.Rows.Count - 1)("sDescript") = ""
            .Rows(.Rows.Count - 1)("xComboIDx") = ""
            .Rows(.Rows.Count - 1)("cStatusxx") = "1"
            .Rows(.Rows.Count - 1)("xQuantity") = 0.0
        End With
    End Sub

    Private Sub loadComboTable()
        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sComboIDx" & _
                       ", a.nEntryNox" & _
                       ", a.sStockIDx" & _
                       ", a.nQuantity" & _
                       ", a.dModified" & _
                       ", b.sBarCodex" & _
                       ", b.sDescript" & _
                       ", a.sComboIDx xComboIDx" & _
                       ", a.nQuantity xQuantity" & _
                       ", '0' cStatusxx" & _
             " FROM " & p_sMasTable3 & " a" & _
                " LEFT JOIN " & p_sMasTable1 & " b ON a.sStockIDx = b.sStockIDx" & _
             " WHERE a.sComboIDx = " & strParm(p_oDTMstr(0).Item("sStockIDx"))
        p_oDTComb = p_oApp.ExecuteQuery(lsSQL)
    End Sub

    Private Function isEntryOk() As Boolean
        If Trim(p_oDTMstr(0).Item("sBarCodex")) = "" Then
            MsgBox("Barcode No seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        If Trim(p_oDTMstr(0).Item("sDescript")) = "" Then
            MsgBox("sDescript seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'If Trim(p_oDTMstr(0).Item("cWthPromo")) = 0 Then
        '    If Trim(p_oDTMstr(0).Item("nUnitPrce")) = 0 Then
        '        MsgBox("Unit Price seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
        '        Return False
        '    End If
        'End If

        'If Trim(p_oDTMstr(0).Item("cWthPromo")) = 0 Then
        '    If Trim(p_oDTMstr(0).Item("nSelPrice")) = 0 Then
        '        MsgBox("Selling Price seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
        '        Return False
        '    End If
        'End If

        'Check for validity of entries if item is a COMBO MEAL
        If p_oDTMstr(0).Item("cComboMlx") = 1 Then
            Dim lnCtr As Integer
            Dim lnValidCtr As Integer

            lnValidCtr = 0

            If Not IsNothing(p_oDTComb) Then
                For lnCtr = 0 To p_oDTComb.Rows.Count - 1
                    'Check if entry for stockid and status is okey
                    'We don't need to check if status is remove...
                    If p_oDTComb(lnCtr).Item("sStockIDx") <> "" And p_oDTComb(lnCtr).Item("cStatusxx") <> "2" Then
                        'Check if quantity is valid
                        If p_oDTComb(lnCtr).Item("nQuantity") <= 0 Then
                            MsgBox("Combo item " & p_oDTComb(lnCtr).Item("sDescript") & " has an invalid quantity!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                            Return False
                        End If

                        'Count the total number of  valid combo...
                        lnValidCtr = lnValidCtr + 1
                    End If
                Next
            End If

            'Check if combo meal has more than 1 valid item...
            If lnValidCtr < 2 Then
                MsgBox("Combo meal with less than 2 item is not allowed!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, p_sMsgHeadr)
                Return False
            End If
        End If

        Return True
    End Function

    Private Function getSQ_Master() As String
        Return "SELECT a.sStockIDx" &
                    ", a.sBarCodex" &
                    ", a.sDescript" &
                    ", a.sBriefDsc" &
                    ", a.sCategrID" &
                    ", a.sSizeIDxx" &
                    ", a.sMeasurID" &
                    ", a.sInvTypID" &
                    ", a.nUnitPrce" &
                    ", a.nSelPrice" &
                    ", a.nDiscLev1" &
                    ", a.nDiscLev2" &
                    ", a.nDiscLev3" &
                    ", a.nDealrDsc" &
                    ", b.sBranchCd" &
                    ", b.sSectnIDx" &
                    ", b.sBinIDxxx" &
                    ", b.dBegInvxx" &
                    ", b.nBegQtyxx" &
                    ", b.nQtyOnHnd" &
                    ", b.nMinLevel" &
                    ", b.nMaxLevel" &
                    ", b.nResvOrdr" &
                    ", b.nBackOrdr" &
                    ", b.nAveMonSl" &
                    ", a.cComboMlx" &
                    ", a.cWthPromo" &
                    ", a.cRecdStat" &
                    ", a.sModified" &
                    ", a.dModified" &
                    ", IFNULL(a.sStockIDx, '') sStockID1" &
                    ", IFNULL(b.sStockIDx, '') sStockID2" &
                    ", IFNULL(sImgePath, '') sImgePath" &
                " FROM " & p_sMasTable1 & " a" &
                    " LEFT JOIN " & p_sMasTable2 & " b on a.sStockIDx = b.sStockIDx" &
                        " AND b.sBranchCd = " & strParm(p_oApp.BranchCode)
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT a.sStockIDx" & _
                    ", a.sBarCodex" & _
                    ", a.sBriefDsc" & _
                    ", a.sDescript" & _
                    ", IFNULL(b.sStockIDx, '') xStockIDx" & _
              " FROM " & p_sMasTable1 & " a" & _
                  " LEFT JOIN " & p_sMasTable2 & " b ON a.sStockIDx = b.sStockIDx"
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Private Class Others
        Public sCategrNm As String
        Public sSizeName As String
        Public sMeasurNm As String
        Public sSectnNme As String
        Public sBinNamex As String
        Public sInvTypex As String
    End Class
End Class
