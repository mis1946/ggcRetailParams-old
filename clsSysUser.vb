'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     System User Maintenance Object
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
'  Kalyptus [ 11/05/2016 09:15 am ]
'      Started creating this object.
'  Note: 
'Log Name
'Password 

'UserName
'Employee No
'Department
'Position 

'User Level : Engineer-Lahat;Sysadmin-lahat ng below sysadmin
'User Type

'Branch     :  Engineer/Puwede lagyan;Sysadmin
'Product ID :  Same as Branch
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ggcAppDriver

Public Class clsSysUser
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oOthersx As New Others

    Private p_nEditMode As xeEditMode
    Private p_sParent As String

    Private Const p_sMasTable As String = "xxxSysUser"
    Private Const xsSignature As String = "08220326"
    Private Const p_sMsgHeadr As String = "System User Maintenance"

    Public Event MasterRetrieved(ByVal Index As Integer, _
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
                    Case 3
                        If p_oDTMstr(0).Item(4) <> "" Then
                            Return p_oOthersx.sClientNm
                        Else
                            Return p_oDTMstr(0).Item(Index)
                        End If
                    Case 81
                        Return p_oOthersx.sDeptName
                    Case 82
                        Return p_oOthersx.sPositnNm
                    Case 83
                        Return p_oOthersx.sBranchNm
                    Case 84
                        Return p_oOthersx.sProdctNm
                    Case 6
                        Return DecToNum(p_oDTMstr(0).Item(Index))
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
                    Case 80, 81, 82, 84 'others
                    Case 83
                        Call getBranch(5, 82, value, False, False)
                    Case 0, 15, 16, 17  'sUserIDxx, cUserStat, sModified, dModified
                    Case 4
                        Call getClient(4, 80, value, True, False)
                    Case 8
                        Call getProduct(8, 84, value, True, False)
                    Case 10             'nSysError
                        If IsNumeric(value) Then
                            p_oDTMstr(0).Item(Index) = value
                        End If
                        RaiseEvent MasterRetrieved(Index, p_oDTMstr(0).Item(Index))
                    Case 6 'nUserLevl
                        p_oDTMstr(0).Item(Index) = NumToDec(value)
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
                Select LCase(Index)
                    Case "sclientnm"
                        Return p_oOthersx.sClientNm
                    Case "sdeptname"
                        Return p_oOthersx.sDeptName
                    Case "spositnnm"
                        Return p_oOthersx.sPositnNm
                    Case "sbranchnm"
                        Return p_oOthersx.sBranchNm
                    Case "sprodctnm"
                        Return p_oOthersx.sProdctNm
                    Case "nuserlevl"
                        Return DecToNum(p_oDTMstr(0).Item(Index))
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
                    Case "semployno"
                        Call getClient(4, 80, value, True, False)
                    Case "sclientnm", "sdeptname", "spositnme"
                    Case "sbranchnm"
                        Call getBranch(5, 83, value, False, False)
                    Case "sprodctnm"
                        Call getProduct(8, 84, value, False, False)
                    Case "nuserlevl"
                        Master(p_oDTMstr.Columns(Index).Ordinal) = NumToDec(value)
                    Case Else
                        Master(p_oDTMstr.Columns(Index).Ordinal) = value
                End Select
            End If
        End Set
    End Property

    ReadOnly Property Other(ByVal Index As String)
        Get
            Select Case LCase(Index)
                Case "sclientnm"
                    Other = p_oOthersx.sClientNm
                Case "sdeptname"
                    Other = p_oOthersx.sDeptName
                Case "spositnnm"
                    Other = p_oOthersx.sPositnNm
                Case "sbranchnm"
                    Other = p_oOthersx.sBranchNm
                Case Else
                    Other = ""
            End Select
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

        lsSQL = AddCondition(getSQ_Master, "sUserIDxx = " & strParm(fsRecdIDxx))
        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        If p_oDTMstr.Rows.Count <= 0 Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        End If

        p_oDTMstr(0).Item("sLogNamex") = Decrypt(p_oDTMstr(0).Item("sLogNamex"), xsSignature)
        p_oDTMstr(0).Item("sPassword") = Decrypt(p_oDTMstr(0).Item("sPassword"), xsSignature)
        p_oDTMstr(0).Item("sUserName") = Decrypt(p_oDTMstr(0).Item("sUserName"), xsSignature)

        Call initOthers()

        p_nEditMode = xeEditMode.MODE_READY
        Return True
    End Function

    'Public Function SearchMaster
    Public Sub SearchMaster(ByVal fnIndex As Integer, _
                                 ByVal fsValue As String)
        Select Case fnIndex
            Case 4
                Call getClient(4, fnIndex, fsValue, False, True)
            Case 83
                Call getBranch(5, fnIndex, fsValue, False, True)
            Case 84
                Call getProduct(8, fnIndex, fsValue, False, True)
        End Select
    End Sub

    Public Function SearchRecord(ByVal fvUserName As String, _
                                 ByVal fvPassword As String) As DataTable
        Dim lsSQL As String
        Dim lsCondition As String
        Dim loDT As DataTable

        lsSQL = "SELECT sUserIDxx" & _
                    ", nUserLevl" & _
                    ", sProdctID" & _
                " FROM xxxSysUser" & _
                " WHERE sProdctID = " & strParm(p_oApp.ProductID) & _
                    " AND cUserStat = '1'" & _
                " ORDER BY nUserLevl DESC LIMIT 1"

        If fvUserName = "" Or fvPassword = "" Then Return Nothing

        lsCondition = "sLogNamex = " & strParm(fvUserName) & _
                    " AND sPassword = " & strParm(fvPassword)

        lsSQL = AddCondition(lsSQL, lsCondition)

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then Return Nothing

        Return loDT
    End Function

    'Public Function SearchTransaction(String, Boolean, Boolean=False)
    Public Function SearchRecord( _
                        ByVal fsValue As String _
                      , Optional ByVal fbByCode As Boolean = False) As Boolean

        Dim lsSQL As String

        'Check if already loaded base on edit mode
        If p_nEditMode = xeEditMode.MODE_READY Or p_nEditMode = xeEditMode.MODE_UPDATE Then
            If fbByCode Then
                If fsValue = p_oDTMstr(0).Item("sUserIDxx") Then Return True
            Else
                If fsValue = p_oDTMstr(0).Item("sUserName") Then Return True
            End If
        End If

        'Initialize SQL filter
        lsSQL = getSQ_Browse()

        'create Kwiksearch filter
        Dim lsFilter As String
        If fbByCode Then
            lsFilter = "sUserIDxx = " & strParm(fsValue)
        Else
            lsFilter = "sUserName LIKE " & strParm(fsValue & "%")
        End If

        'Dim loDT As DataTable
        'loDT = p_oApp.ExecuteQuery(AddCondition(lsSQL, lsFilter))

        'If loDT.Rows.Count = 0 Then
        '    p_nEditMode = xeEditMode.MODE_UNKNOWN
        '    Return False
        'Else

        'End If

        Dim loDta As DataRow = KwikSearch(p_oApp _
                                        , lsSQL _
                                        , True _
                                        , fsValue _
                                        , "sUserIDxx»sUserName" _
                                        , "User ID»User Name" _
                                        , "" _
                                        , "sUserIDxx»sUserName" _
                                        , IIf(fbByCode, 0, 1))

        If IsNothing(loDta) Then
            p_nEditMode = xeEditMode.MODE_UNKNOWN
            Return False
        Else
            Return OpenRecord(loDta.Item("sUserIDxx"))
        End If
    End Function

    'Public Function SaveTransaction
    Public Function SaveRecord() As Boolean
        If Not (p_nEditMode = xeEditMode.MODE_ADDNEW Or _
                p_nEditMode = xeEditMode.MODE_READY Or _
                p_nEditMode = xeEditMode.MODE_UPDATE) Then

            MsgBox("Invalid Edit Mode detected!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, p_sMsgHeadr)
            Return False
        End If

        If Not isEntryOk() Then
            Return False
        End If

        Dim lsSQL As String = ""

        If p_sParent = "" Then p_oApp.BeginTransaction()

        p_oDTMstr(0).Item("sUserName") = Decrypt(p_oDTMstr(0).Item("sUserName"), xsSignature)
        p_oDTMstr(0).Item("sLogNamex") = Decrypt(p_oDTMstr(0).Item("sLogNamex"), xsSignature)
        p_oDTMstr(0).Item("sPassword") = Decrypt(p_oDTMstr(0).Item("sPassword"), xsSignature)

        'Save master table 
        If p_nEditMode = xeEditMode.MODE_ADDNEW Then
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, , p_oApp.UserID, p_oApp.SysDate)
        Else
            lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sUserIDxx = " & strParm(p_oDTMstr(0).Item("sUserIDxx")), p_oApp.UserID, Format(p_oApp.SysDate, xsDATE_SHORT))
        End If

        If lsSQL <> "" Then
            p_oApp.Execute(lsSQL, p_sMasTable)
        End If

        If p_sParent = "" Then p_oApp.CommitTransaction()

        p_nEditMode = xeEditMode.MODE_READY

        Return True
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

        p_oDTMstr(0).Item("cUserStat") = "0"
        lsSQL = ADO2SQL(p_oDTMstr, p_sMasTable, "sUserIDxx = " & strParm(p_oDTMstr(0).Item("sUserIDxx")))
        p_oApp.Execute(lsSQL, p_sMasTable, p_oApp.BranchCode)

        If p_sParent = "" Then p_oApp.CommitTransaction()

        Return True
    End Function

    Public Function UpdateRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Public Function CancelUpdate() As Boolean
        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    'This method implements a search master where id and desc are not joined.
    Private Sub getClient(ByVal fnColIdx As Integer _
                          , ByVal fnColDsc As Integer _
                          , ByVal fsValue As String _
                          , ByVal fbIsCode As Boolean _
                          , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sClientNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sClientNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sClientID" & _
                       ", CONCAT(a.sFrstName, ' ', LEFT(a.sMiddName, 1), '. ', a.sLastName, IF(IFNULL(a.sSuffixNm, '') = '', '', CONCAT(' ', a.sSuffixNm))) sCompnyNm" & _
                       ", c.sDeptName" & _
                       ", d.sPositnNm" & _
               " FROM Employee_Master001 b" & _
                " LEFT JOIN Client_Master a ON b.sEmployID = a.sClientID" & _
                " LEFT JOIN Department c ON b.sDeptIDxx = c.sDeptIDxx" & _
                " LEFT JOIN `Position` d ON b.sPositnID = d.sPositnID" & _
        IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison

        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sClientID»sCompnyNm»sDeptName" _
                                             , "ID»Employee Name»Address", _
                                             , "a.sClientID»CONCAT(a.sFrstName, ' ', LEFT(a.sMiddName, 1), '. ', a.sLastName, IF(IFNULL(a.sSuffixNm, '') = '', '', CONCAT(' ', a.sSuffixNm)))»c.sDeptName" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sClientNm = ""
                p_oOthersx.sDeptName = ""
                p_oOthersx.sPositnNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sClientID")
                p_oOthersx.sClientNm = loRow.Item("sCompnyNm")
                p_oOthersx.sDeptName = loRow.Item("sDeptName")
                p_oOthersx.sPositnNm = loRow.Item("sPositnNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oDTMstr(0).Item(fnColIdx))
            RaiseEvent MasterRetrieved(3, p_oOthersx.sClientNm)
            RaiseEvent MasterRetrieved(81, p_oOthersx.sDeptName)
            RaiseEvent MasterRetrieved(82, p_oOthersx.sPositnNm)

            Exit Sub

        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sClientID = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sCompnyNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sClientNm = ""
            p_oOthersx.sDeptName = ""
            p_oOthersx.sPositnNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sClientID")
            p_oOthersx.sClientNm = loDta(0).Item("sCompnyNm")
            p_oOthersx.sDeptName = loDta(0).Item("sDeptName")
            p_oOthersx.sPositnNm = loDta(0).Item("sPositnNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oDTMstr(0).Item(fnColIdx))
        RaiseEvent MasterRetrieved(3, p_oOthersx.sClientNm)
        RaiseEvent MasterRetrieved(81, p_oOthersx.sDeptName)
        RaiseEvent MasterRetrieved(82, p_oOthersx.sPositnNm)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getBranch(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sBranchNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sBranchNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT" & _
                       "  a.sBranchCD" & _
                       ", a.sBranchNm" & _
               " FROM `Branch` a" & _
               IIf(fbIsCode = False, " WHERE a.cRecdStat = '1'", "")

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sBranchCD»sBranchNm" _
                                             , "Code»Branch", _
                                             , "a.sBranchCD»a.sBranchNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sBranchNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sBranchCD")
                p_oOthersx.sBranchNm = loRow.Item("sBranchNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBranchNm)
            Exit Sub

        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "a.sBranchCD = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "a.sBranchNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sBranchNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sBranchCD")
            p_oOthersx.sBranchNm = loDta(0).Item("sBranchNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sBranchNm)
    End Sub

    'This method implements a search master where id and desc are not joined.
    Private Sub getProduct(ByVal fnColIdx As Integer _
                        , ByVal fnColDsc As Integer _
                        , ByVal fsValue As String _
                        , ByVal fbIsCode As Boolean _
                        , ByVal fbIsSrch As Boolean)

        'Compare the value to be search against the value in our column
        If fbIsCode Then
            If fsValue = p_oDTMstr(0).Item(fnColIdx) And fsValue <> "" And p_oOthersx.sProdctNm <> "" Then Exit Sub
        Else
            If fsValue = p_oOthersx.sProdctNm And fsValue <> "" Then Exit Sub
        End If

        Dim lsSQL As String
        lsSQL = "SELECT sProdctID" & _
                    ", sProdctNm" & _
                " FROM xxxSysObject" & _
                " ORDER BY sApplName"

        'Are we using like comparison or equality comparison
        If fbIsSrch Then
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                             , lsSQL _
                                             , True _
                                             , fsValue _
                                             , "sProdctID»sProdctNm" _
                                             , "Code»Product Name", _
                                             , "sProdctID»sProdctNm" _
                                             , IIf(fbIsCode, 0, 1))
            If IsNothing(loRow) Then
                p_oDTMstr(0).Item(fnColIdx) = ""
                p_oOthersx.sProdctNm = ""
            Else
                p_oDTMstr(0).Item(fnColIdx) = loRow.Item("sProdctID")
                p_oOthersx.sProdctNm = loRow.Item("sProdctNm")
            End If

            RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sProdctNm)
            Exit Sub
        End If

        If fsValue <> "" Then
            If fbIsCode Then
                lsSQL = AddCondition(lsSQL, "sBranchCD = " & strParm(fsValue))
            Else
                lsSQL = AddCondition(lsSQL, "sBranchNm = " & strParm(fsValue))
            End If
        End If

        Dim loDta As DataTable
        loDta = p_oApp.ExecuteQuery(lsSQL)

        If loDta.Rows.Count = 0 Then
            p_oDTMstr(0).Item(fnColIdx) = ""
            p_oOthersx.sProdctNm = ""
        ElseIf loDta.Rows.Count = 1 Then
            p_oDTMstr(0).Item(fnColIdx) = loDta(0).Item("sProdctID")
            p_oOthersx.sProdctNm = loDta(0).Item("sProdctNm")
        End If

        RaiseEvent MasterRetrieved(fnColDsc, p_oOthersx.sProdctNm)
    End Sub

    Private Sub initMaster()
        Dim lnCtr As Integer
        For lnCtr = 0 To p_oDTMstr.Columns.Count - 1
            Select Case LCase(p_oDTMstr.Columns(lnCtr).ColumnName)
                Case "suseridxx"
                    p_oDTMstr(0).Item(lnCtr) = GetNextCode(p_sMasTable, "sUserIDxx", True, p_oApp.Connection, True, p_oApp.BranchCode)
                Case "nuserlevl", "nsyserror"
                    p_oDTMstr(0).Item(lnCtr) = 0
                Case "dmodified", "smodified"
                Case "cusertype", "clogstatx", "clockstat"
                    p_oDTMstr(0).Item(lnCtr) = "0"
                Case "cuserstat", "callwlock", "callwview"
                    p_oDTMstr(0).Item(lnCtr) = "1"
                Case Else
                    p_oDTMstr(0).Item(lnCtr) = ""
            End Select
        Next
    End Sub

    Private Sub initOthers()
        p_oOthersx.sClientNm = ""
        p_oOthersx.sDeptName = ""
        p_oOthersx.sPositnNm = ""
        p_oOthersx.sBranchNm = ""
    End Sub

    Private Function isEntryOk() As Boolean
        'Check for the information about the card
        If Trim(p_oDTMstr(0).Item("sLogNamex")) = "" Then
            MsgBox("Log-in Name seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Trim(p_oDTMstr(0).Item("sPassword")) = "" Then
            MsgBox("Password seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        'Check how much does he intends to borrow
        If Trim(p_oDTMstr(0).Item("sUserName")) = "" Then
            MsgBox("User's Full Name seems to have a problem! Please check your entry....", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, p_sMsgHeadr)
            Return False
        End If

        Return True
    End Function

    Private Function getSQ_Master() As String
        Return "SELECT a.sUserIDxx" & _
                    ", a.sLogNamex" & _
                    ", a.sPassword" & _
                    ", a.sUserName" & _
                    ", a.sEmployNo" & _
                    ", a.sBranchCD" & _
                    ", a.nUserLevl" & _
                    ", a.cUserType" & _
                    ", a.sProdctID" & _
                    ", a.sCompName" & _
                    ", a.nSysError" & _
                    ", a.cAllwLock" & _
                    ", a.cAllwView" & _
                    ", a.cLogStatx" & _
                    ", a.cLockStat" & _
                    ", a.cUserStat" & _
                    ", a.sModified" & _
                    ", a.dModified" & _
                " FROM " & p_sMasTable & " a"
    End Function

    Private Function getSQ_Browse() As String
        Return "SELECT sUserIDxx" & _
                    ", sUserName" & _
              " FROM " & p_sMasTable & _
              " ORDER BY dModified DESC"
    End Function

    Private Function NumToDec(ByVal fnValue As Integer) As Integer
        Select Case fnValue
            Case 0 : Return 1
            Case 1 : Return 2
            Case 2 : Return 4
            Case 3 : Return 8
            Case 4 : Return 16
            Case 5 : Return 32
            Case 6 : Return 64
            Case 7 : Return 128
            Case Else : Return 0
        End Select
    End Function

    Private Function DecToNum(ByVal fnValue As Integer) As Integer
        Select Case fnValue
            Case 1 : Return 0
            Case 2 : Return 1
            Case 4 : Return 2
            Case 8 : Return 3
            Case 16 : Return 4
            Case 32 : Return 5
            Case 64 : Return 6
            Case 128 : Return 7
            Case Else : Return 0
        End Select
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_nEditMode = xeEditMode.MODE_UNKNOWN
    End Sub

    Private Class Others
        Public sClientNm As String
        Public sDeptName As String
        Public sPositnNm As String
        Public sBranchNm As String
        Public sProdctNm As String
    End Class
End Class
