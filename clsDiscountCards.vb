'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Retail Discount Cards
'
' Copyright 2012 and Beyond
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
'  iMac [ 10/12/2016 01:20 pm ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Imports MySql.Data.MySqlClient
Imports ggcAppDriver

Public Class clsDiscountCards
    Enum xeSpclDisc
        xeUnknown = 0
        xeSpecial = 1
        xeNonSpecial = 2
    End Enum

    Private Const pxeModuleName As String = "clsDiscountCards"
    Private Const pxeMasterTble As String = "Discount_Card"
    Private Const pxeDetailTble As String = "Discount_Card_Detail"

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetl As DataTable
    Private p_oCategr As clsProductCategory
    Private p_nEditMode As xeEditMode
    Private p_nRecdStat As xeRecordStat
    Private p_sCategIDx As String

    Public Event MasterRetreive(ByVal lnIndex As Integer)
    Public Event DetailRetreive(ByVal lnRow As Integer, ByVal lnIndex As Integer)

    Public Property Master(ByVal Index As Integer) As Object
        Get
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 6
                    Master = p_oDTMstr(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get

        Set(value As Object)
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 6
                    p_oDTMstr(0)(Index) = value

                    RaiseEvent MasterRetreive(Index)
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Set
    End Property

    Public Property Master(ByVal Index As String) As Object
        Get
            Select Case Index
                Case "sCardIDxx", "sCardDesc", "sCompnyCd", "dPrtSince", "dStartxxx", "dExpiratn", "cNoneVatx", "cRecdStat"
                    Master = p_oDTMstr(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get

        Set(value As Object)
            Select Case Index
                Case "sCardIDxx", "sCardDesc", "sCompnyCd", "dPrtSince", "dStartxxx", "dExpiratn", "cNoneVatx", "cRecdStat"
                    p_oDTMstr(0)(Index) = value
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Set
    End Property

    Public Property Detail(ByVal Row As Integer, ByVal Index As Integer) As Object
        Get
            Select Case Index
                Case 1, 2, 3, 4
                    Detail = p_oDTDetl(Row)(Index)
                Case Else
                    Detail = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set!", MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 1
                    searchDetail(Row, value, False)
                Case 2, 3, 4
                    p_oDTDetl(Row)(Index) = value
                    RaiseEvent DetailRetreive(Row, Index)
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set!", MsgBoxStyle.Critical, "Warning")
            End Select
        End Set
    End Property

    Public Property Detail(ByVal Row As Integer, ByVal Index As String) As Object
        Get
            Select Case Index
                Case "sCategrID", "nMinAmtxx", "nDiscRate", "sDescript", "nDiscAmtx", "sDescript"
                    Detail = p_oDTDetl(Row)(Index)
                Case Else
                    Detail = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set!", MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sCategrID", "nMinAmtxx", "nDiscRate", "sDescript", "nDiscAmtx", "sDescript"
                    p_oDTDetl(Row)(Index) = value
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set!", MsgBoxStyle.Critical, "Warning")
            End Select
        End Set
    End Property

    ReadOnly Property EditMode()
        Get
            Return p_nEditMode
        End Get
    End Property

    ReadOnly Property ItemCount()
        Get
            Return p_oDTDetl.Rows.Count
        End Get
    End Property

    WriteOnly Property FilterCategory As String
        Set(ByVal Value As String)
            p_sCategIDx = Value
        End Set
    End Property

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider

        p_oCategr = New clsProductCategory(p_oApp)
        p_nRecdStat = -1
        InitRecord()
    End Sub

    Function InitRecord() As Boolean
        p_oDTMstr = Nothing
        p_oDTDetl = Nothing

        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Function NewRecord() As Boolean
        p_oDTMstr = p_oApp.ExecuteQuery(AddCondition(getSQLMaster, "0=1"))
        p_oDTMstr.Rows.Add(p_oDTMstr.NewRow())
        p_oDTMstr(0)("sCardIDxx") = GetNextCode(pxeMasterTble, "sCardIDxx", False, p_oApp.Connection)
        p_oDTMstr(0)("sCardDesc") = ""
        p_oDTMstr(0)("sCompnyCd") = ""
        p_oDTMstr(0)("dPrtSince") = p_oApp.SysDate
        p_oDTMstr(0)("dStartxxx") = p_oApp.SysDate
        p_oDTMstr(0)("dExpiratn") = p_oApp.SysDate
        p_oDTMstr(0)("cNoneVatx") = CInt(xeRecordStat.RECORD_EMPTY)
        p_oDTMstr(0)("cRecdStat") = CInt(xeRecordStat.RECORD_NEW)

        p_oDTDetl = p_oApp.ExecuteQuery(AddCondition(getSQLDetail, "0=1"))
        AddDetail()

        p_nEditMode = xeEditMode.MODE_ADDNEW

        Return True
    End Function

    Function UpdateRecord() As Boolean
        If IsNothing(p_oDTMstr) Then Return False
        If p_oDTMstr(0)("sCardIDxx") = "" Then Return False

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Function SaveRecord() As Boolean
        Dim lsSQL As String
        Dim lsSQL1 As String
        Dim lnCtr As Integer

        Dim lsProcName As String = pxeModuleName & "." & "SaveRecord"

        If p_nEditMode <> xeEditMode.MODE_ADDNEW And p_nEditMode <> xeEditMode.MODE_UPDATE Then Return False

        Try
            p_oApp.BeginTransaction()
            'save master
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0)("sCardIDxx") = GetNextCode(pxeMasterTble, "sCardIDxx", True, p_oApp.Connection, False)
                lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTble, , p_oApp.UserID, p_oApp.SysDate)
            Else
                lsSQL = "UPDATE " & pxeMasterTble & _
                            " SET  sCardDesc= " & strParm(p_oDTMstr(0)("sCardDesc")) & _
                                ", sCompnyCd= " & strParm(p_oDTMstr(0)("sCompnyCd")) & _
                                ", dPrtSince= " & dateParm(p_oDTMstr(0)("dPrtSince")) & _
                                ", dStartxxx= " & dateParm(p_oDTMstr(0)("dStartxxx")) & _
                                ", dExpiratn= " & dateParm(p_oDTMstr(0)("dExpiratn")) & _
                                ", cNoneVatx= " & strParm(p_oDTMstr(0)("cNoneVatx")) & _
                                ", cRecdStat= " & strParm(p_oDTMstr(0)("cRecdStat")) & _
                                ", sModified= " & strParm(p_oApp.UserID) & _
                                ", dModified= " & datetimeParm(p_oApp.SysDate) & _
                            " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))

                lsSQL1 = "DELETE FROM " & pxeDetailTble & " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))
                p_oApp.Execute(lsSQL1, pxeDetailTble)
            End If
            p_oApp.Execute(lsSQL, pxeMasterTble)

            'save detail
            For lnCtr = 0 To p_oDTDetl.Rows.Count - 1
                If p_oDTDetl(lnCtr)("sCategrID") = "" Then
                    p_oDTDetl(lnCtr).Delete()
                Else
                    Dim loDT As DataTable
                    loDT = New DataTable
                    loDT = p_oDTDetl.Clone
                    loDT.ImportRow(p_oDTDetl(lnCtr))
                    loDT(0)("sCardIDxx") = p_oDTMstr(0)("sCardIDxx")

                    lsSQL = ADO2SQL(loDT, pxeDetailTble, , , p_oApp.SysDate, "sDescript")
                    p_oApp.Execute(lsSQL, pxeDetailTble) 'individual saving
                End If
            Next
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")

            Return False
        End Try

        p_oApp.CommitTransaction()
        MsgBox("Record Saved Successfully.", MsgBoxStyle.Information, "Success")

        Return InitRecord()
    End Function

    Function SaveTransaction() As Boolean
        Dim lsSQL As String
        Dim lsSQL1 As String
        Dim lnCtr As Integer

        Dim lsProcName As String = pxeModuleName & "." & "SaveTransaction"

        If p_nEditMode <> xeEditMode.MODE_ADDNEW And p_nEditMode <> xeEditMode.MODE_UPDATE Then Return False

        Try
            p_oApp.BeginTransaction()
            'save master
            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                p_oDTMstr(0)("sCardIDxx") = GetNextCode(pxeMasterTble, "sCardIDxx", True, p_oApp.Connection, False)
                lsSQL = ADO2SQL(p_oDTMstr, pxeMasterTble, , p_oApp.UserID, p_oApp.SysDate)
            Else
                lsSQL = "UPDATE " & pxeMasterTble & _
                            " SET  sCardDesc= " & strParm(p_oDTMstr(0)("sCardDesc")) & _
                                ", sCompnyCd= " & strParm(p_oDTMstr(0)("sCompnyCd")) & _
                                ", dPrtSince= " & dateParm(p_oDTMstr(0)("dPrtSince")) & _
                                ", dStartxxx= " & dateParm(p_oDTMstr(0)("dStartxxx")) & _
                                ", dExpiratn= " & dateParm(p_oDTMstr(0)("dExpiratn")) & _
                                ", cNoneVatx= " & strParm(p_oDTMstr(0)("cNoneVatx")) & _
                                ", cRecdStat= " & strParm(p_oDTMstr(0)("cRecdStat")) & _
                                ", sModified= " & strParm(p_oApp.UserID) & _
                                ", dModified= " & datetimeParm(p_oApp.SysDate) & _
                            " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))

                lsSQL1 = "DELETE FROM " & pxeDetailTble & " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))
                p_oApp.Execute(lsSQL1, pxeDetailTble)
            End If
            p_oApp.Execute(lsSQL, pxeMasterTble)


            Dim loDT As DataTable
            loDT = New DataTable
            loDT = p_oDTDetl.Clone
            loDT.ImportRow(p_oDTDetl(lnCtr))
            loDT(0)("sCardIDxx") = p_oDTMstr(0)("sCardIDxx")
            loDT(0)("sCategrID") = ""

            lsSQL = ADO2SQL(loDT, pxeDetailTble, , , p_oApp.SysDate, "sDescript")
            p_oApp.Execute(lsSQL, pxeDetailTble) 'individual saving

        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")

            Return False
        End Try

        p_oApp.CommitTransaction()
        MsgBox("Record Saved Successfully.", MsgBoxStyle.Information, "Success")

        Return InitRecord()
    End Function

    Function BrowseRecord(Optional fnDiscount As xeSpclDisc = xeSpclDisc.xeUnknown) As Boolean
        Dim loDT As DataTable
        Dim lsSQL As String

        Select Case fnDiscount
            Case xeSpclDisc.xeUnknown
                lsSQL = getSQLMaster()
            Case Else
                lsSQL = "SELECT" & _
                             "  a.sCardIDxx" & _
                             ", a.sCardDesc" & _
                             ", a.sCompnyCd" & _
                             ", a.dPrtSince" & _
                             ", a.dStartxxx" & _
                             ", a.dExpiratn" & _
                             ", a.cNoneVatx" & _
                             ", a.cRecdStat" & _
                             ", a.sModified" & _
                             ", a.dModified" & _
                        " FROM " & pxeMasterTble & " a" & _
                            ", Discount_Card_Detail b" & _
                        " WHERE a.sCardIDxx = b.sCardIDxx" & _
                            " AND b.sCategrID" & IIf(fnDiscount = xeSpclDisc.xeSpecial, " = ''", " <> ''") & _
                        " GROUP BY a.sCardIDxx" & _
                        " ORDER BY a.sCardIDxx ASC"
        End Select

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 1 Then
            Return OpenRecord(loDT(0)("sCardIDxx"))
        Else
            Dim loRow As DataRow = KwikSearch(p_oApp _
                                            , lsSQL _
                                            , True _
                                            , "" _
                                            , "sCardIDxx»sCardDesc»sCompnyCd" _
                                            , "ID»Descript»Company", _
                                            , "sCardIDxx»sCardDesc»sCompnyCd" _
                                            , 1)
            If IsNothing(loRow) Then
                Return False
            Else
                Return OpenRecord(loRow.Item("sCardIDxx"))
            End If
        End If
    End Function

    Function OpenRecord(ByVal fsCardIDxx As String) As Boolean
        Dim lsSQL As String

        lsSQL = AddCondition(getSQLMaster, "sCardIDxx = " & strParm(fsCardIDxx))

        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)
        If p_oDTMstr.Rows.Count = 0 Then GoTo endwithClear

        lsSQL = AddCondition(getSQLDetail, "sCardIDxx = " & strParm(fsCardIDxx))

        p_oDTDetl = p_oApp.ExecuteQuery(lsSQL)
        If ItemCount = 0 Then AddDetail()

        Return True
endwithClear:
        p_oDTMstr = Nothing
        p_oDTDetl = Nothing
        Return False
    End Function

    Function DeleteRecord() As Boolean
        Dim lsSQL As String

        Dim lsProcName As String = pxeModuleName & ".DeleteRecord"

        If IsNothing(p_oDTMstr) Then Return False
        If MsgBox("Are you sure to delete this record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
            Return False
        End If

        Try
            p_oApp.BeginTransaction()

            lsSQL = "DELETE FROM " & pxeMasterTble & " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))
            p_oApp.Execute(lsSQL, pxeMasterTble)

            lsSQL = "DELETE FROM " & pxeDetailTble & " WHERE sCardIDxx = " & strParm(p_oDTMstr(0)("sCardIDxx"))
            p_oApp.Execute(lsSQL, pxeDetailTble)
        Catch ex As Exception
            p_oApp.RollBackTransaction()

            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")

            Return False
        End Try

        p_oApp.CommitTransaction()
        MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, "Success")

        Return InitRecord()
    End Function

    Function SearchItem(ByVal fnRow As Integer, _
                        ByVal fsValue As String) As Boolean

        Return searchDetail(fnRow, fsValue, False)
    End Function

    Function SearchCard(Optional ByVal lsValue As String = "",
                         Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sCardIDxx = " & strParm(lsValue))
        Else
            lsSQL = AddCondition(getSQLBrowse, "sCardDesc LIKE " & strParm(lsValue & "%"))
        End If
        'lsSQL = IIf(p_sCategIDx = "", lsSQL, AddCondition(lsSQL, "b.sCategrID IN (" & p_sCategIDx & ")"))

        Debug.Print(lsSQL)
        loDT = p_oApp.ExecuteQuery(lsSQL)


        If loDT.Rows.Count = 0 Then
            Return Nothing
            'ElseIf loDT.Rows.Count = 1 Then
            '    Return loDT.Rows(0)
        Else
            Return loDT.Rows(0)
            'Else
            '    Return loDT.Rows(0)
            '    'Return KwikSearch(p_oApp _
            '    '                , getSQLBrowse _
            '    '                , True _
            '    '                , lsValue _
            '    '                , "sCardIDxx»sCardDesc" _
            '    '                , "Card ID»Cards" _
            '    '                , "" _
            '    '                , "sCardIDxx»sCardDesc" _
            '    '                , IIf(lbCode, 1, 2))
        End If
    End Function

    Function SearchCategoryCard(Optional ByVal lsCardIDx As String = "",
                                Optional ByVal lsCategID As String = "") As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = "SELECT" &
                        "  nMinAmtxx" &
                        ", nDiscRate" &
                        ", nDiscAmtx" &
                        ", cNoneVatx" &
                    " FROM Discount_Card_Detail" &
                    " WHERE sCardIDxx = " & strParm(lsCardIDx) &
                        " AND sCategrID = " & strParm(lsCategID)

        loDT = p_oApp.ExecuteQuery(lsSQL)


        If loDT.Rows.Count = 0 Then
            Return Nothing
            'ElseIf loDT.Rows.Count = 1 Then
            '    Return loDT.Rows(0)
        Else
            Return loDT.Rows(0)
            'Else
            '    Return loDT.Rows(0)
            '    'Return KwikSearch(p_oApp _
            '    '                , getSQLBrowse _
            '    '                , True _
            '    '                , lsValue _
            '    '                , "sCardIDxx»sCardDesc" _
            '    '                , "Card ID»Cards" _
            '    '                , "" _
            '    '                , "sCardIDxx»sCardDesc" _
            '    '                , IIf(lbCode, 1, 2))
        End If
    End Function

    Function GetCard() As DataTable
        Return p_oApp.ExecuteQuery(getSQLBrowse)
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" & _
                    "  sCardIDxx" & _
                    ", sCardDesc" & _
                    ", dExpiratn" & _
                    ", cNoneVatx" & _
                    ", cRecdStat" & _
                " FROM " & pxeMasterTble & _
                " WHERE cRecdStat > " & strParm(xeLogical.NO)

        'Return "SELECT" & _
        '            " a.sCardIDxx" & _
        '            ", a.sCardDesc" & _
        '            ", a.dExpiratn" & _
        '            ", a.cNoneVatx" & _
        '            ", a.cRecdStat" & _
        '            ", b.nDiscRate" & _
        '            ", b.nDiscAmtx" & _
        '            " FROM " & pxeMasterTble & " a," & _
        '            " Discount_Card_Detail b" & _
        '        " WHERE a.sCardIDxx = b.sCardIDxx" & _
        '                " AND a.cRecdStat > " & strParm(xeLogical.NO) & _
        '        " GROUP BY a.sCardIDxx"
    End Function

    Function AddDetail() As Boolean
        If ItemCount > 0 Then If p_oDTDetl(ItemCount - 1)("sCategrID") = "" Then Return False
        If ItemCount >= p_oCategr.GetCategory.Rows.Count Then
            MsgBox("All product categories has been encoded.", MsgBoxStyle.Information, "Notice")
            Return False
        End If

        p_oDTDetl.Rows.Add(p_oDTDetl.NewRow())
        p_oDTDetl(ItemCount - 1)("sCardIDxx") = ""
        p_oDTDetl(ItemCount - 1)("sCategrID") = ""
        p_oDTDetl(ItemCount - 1)("nMinAmtxx") = 0
        p_oDTDetl(ItemCount - 1)("nDiscRate") = 0
        p_oDTDetl(ItemCount - 1)("nDiscAmtx") = 0
        p_oDTDetl(ItemCount - 1)("sDescript") = ""

        Return True
    End Function

    Function AddDetails() As Boolean
        Dim loDT As DataTable
        Dim lnCtr As Integer

        loDT = p_oCategr.GetCategory

        If loDT.Rows.Count = 0 Then Return False

        For lnCtr = 0 To loDT.Rows.Count - 1
            p_oDTDetl.Rows.Add(p_oDTDetl.NewRow())
            p_oDTDetl(ItemCount - 1)("sCardIDxx") = ""
            p_oDTDetl(ItemCount - 1)("sCategrID") = loDT(lnCtr)("sCategrCd")
            p_oDTDetl(ItemCount - 1)("nMinAmtxx") = 0
            p_oDTDetl(ItemCount - 1)("nDiscRate") = 0
            p_oDTDetl(ItemCount - 1)("nDiscAmtx") = 0
            p_oDTDetl(ItemCount - 1)("sDescript") = loDT(lnCtr)("sDescript")
        Next

        Return True
    End Function

    Function DeleteDetail(ByVal fnRow As Integer) As Boolean
        p_oDTDetl.Rows.Remove(p_oDTDetl(fnRow))

        If ItemCount = 0 Then Return AddDetail()

        Return True
    End Function

    Private Function searchDetail(ByVal fnRow As Integer, _
                                    ByVal fsValue As String, _
                                    Optional ByVal fbByCode As Boolean = False) As Boolean
        Dim loDR As DataRow

        If fsValue = "" Then Return True

        If fsValue = p_oDTDetl(fnRow)(1) Then Return True

        loDR = p_oCategr.SearchCategory(fsValue, fbByCode)

        If Not IsNothing(loDR) Then
            p_oDTDetl(fnRow)("sCategrID") = loDR(0)
            p_oDTDetl(fnRow)("sDescript") = loDR(1)
        Else
            GoTo endWithClear
        End If

        RaiseEvent DetailRetreive(fnRow, 1)
        Return True

endWithClear:
        p_oDTDetl(fnRow)("sCategrID") = ""
        p_oDTDetl(fnRow)("sDescript") = ""
        RaiseEvent DetailRetreive(fnRow, 1)
        Return False
    End Function

    Private Function getSQLMaster() As String
        Return "SELECT" & _
                     "  sCardIDxx" & _
                     ", sCardDesc" & _
                     ", sCompnyCd" & _
                     ", dPrtSince" & _
                     ", dStartxxx" & _
                     ", dExpiratn" & _
                     ", cNoneVatx" & _
                     ", cRecdStat" & _
                     ", sModified" & _
                     ", dModified" & _
                " FROM " & pxeMasterTble & _
                " ORDER BY sCardIDxx ASC"
    End Function

    Private Function getSQLDetail() As String
        Return "SELECT" & _
                     "  a.sCategrID" & _
                     ", b.sDescript" & _
                     ", a.nMinAmtxx" & _
                     ", a.nDiscRate" & _
                     ", a.nDiscAmtx" & _
                     ", a.dModified" & _
                     ", a.sCardIDxx" & _
                " FROM " & pxeDetailTble & " a" & _
                    " LEFT JOIN Product_Category b" & _
                        " ON a.sCategrID = b.sCategrCd"
    End Function
End Class