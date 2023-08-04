'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Cash Register Machine
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
'  iMac [ 10/19/2016 09:54 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ggcAppDriver

Public Class clsMachine
    Private Const pxeModuleName As String = "clsMachine"
    Private Const pxeTableNamex As String = "Cash_Reg_Machine"

    Private p_oApp As GRider
    Private p_oDT As DataTable
    Private p_nEditMode As xeEditMode
    Private p_nRecdStat As xeRecordStat
    Private p_bModified As Boolean
    Private p_bShowMsgx As Boolean

    Public Event MasterRetreive(ByVal lnIndex As Integer)

    Public Sub New(ByVal foRider As GRider, Optional ByVal fbShowMsg As Boolean = False)
        p_oApp = foRider
        p_nRecdStat = -1
        p_bShowMsgx = fbShowMsg

        InitRecord()
    End Sub

    Public Sub New(ByVal foRider As GRider, ByVal fnRecdStat As xeRecordStat, Optional ByVal fbShowMsg As Boolean = False)
        p_oApp = foRider
        p_nRecdStat = fnRecdStat
        p_bShowMsgx = fbShowMsg

        InitRecord()
    End Sub

    Property Master(ByVal Index As Integer) As Object
        Get
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 7, 9
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case 8
                    If p_oDT(0)(Index) <> CDate(value) Then
                        p_oDT(0)(Index) = CDate(value)
                        p_bModified = True
                    End If
                Case 6
                    If p_oDT(0)(Index) <> value Then
                        If Not IsNumeric(value) Then value = 0

                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select

            RaiseEvent MasterRetreive(Index)
        End Set
    End Property

    Property Master(ByVal Index As String) As Object
        Get
            Select Case Index
                Case "sIDNumber", "sAccredtn", "sApproval", "sPermitNo", "nPOSNumbr", "sORNoxxxx", "cTranMode", "nSalesTot", "dExpiratn", "nSChargex"
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set!", MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sIDNumber", "sAccredtn", "sApproval", "sPermitNo", "nPOSNumbr", "sORNoxxxx", "cTranMode", "nSChargex"
                    If p_oDT(0)(Index) <> value Then
                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case "dExpiratn"
                    If p_oDT(0)(Index) <> CDate(value) Then
                        p_oDT(0)(Index) = CDate(value)
                        p_bModified = True
                    End If
                Case "nSalesTot"
                    If p_oDT(0)(Index) <> value Then
                        If Not IsNumeric(value) Then value = 0

                        p_oDT(0)(Index) = value
                        p_bModified = True
                    End If
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Set
    End Property

    ReadOnly Property MasFldSze(ByVal Index As Integer)
        Get
            Select Case Index
                Case 0, 1, 2, 3, 4, 5
                    MasFldSze = p_oDT.Columns(Index).MaxLength
                Case Else
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
                    Return 0
            End Select
        End Get
    End Property

    ReadOnly Property EditMode()
        Get
            EditMode = p_nEditMode
        End Get
    End Property

    Function InitRecord() As Boolean
        p_oDT = New DataTable
        p_oDT.Columns.Add("sIDNumber", GetType(String)).MaxLength = 17
        p_oDT.Columns.Add("sAccredtn", GetType(String)).MaxLength = 24
        p_oDT.Columns.Add("sApproval", GetType(String)).MaxLength = 24
        p_oDT.Columns.Add("sPermitNo", GetType(String)).MaxLength = 24
        p_oDT.Columns.Add("nPOSNumbr", GetType(String)).MaxLength = 2
        p_oDT.Columns.Add("sORNoxxxx", GetType(String)).MaxLength = 10
        p_oDT.Columns.Add("nSalesTot", GetType(Double))
        p_oDT.Columns.Add("cTranMode", GetType(String)).MaxLength = 1
        p_oDT.Columns.Add("dExpiratn", GetType(Date))
        p_oDT.Columns.Add("nSChargex", GetType(Double))

        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Function NewRecord() As Boolean
        InitRecord()
        If IsNothing(p_oDT) Then Return False

        p_oDT.Rows.Add()
        p_oDT(0)("sIDNumber") = ""
        p_oDT(0)("sAccredtn") = ""
        p_oDT(0)("sApproval") = ""
        p_oDT(0)("sPermitNo") = ""
        p_oDT(0)("nPOSNumbr") = ""
        p_oDT(0)("sORNoxxxx") = ""
        p_oDT(0)("nSalesTot") = 0
        p_oDT(0)("cTranMode") = "D"
        p_oDT(0)("dExpiratn") = p_oApp.SysDate
        p_oDT(0)("nSChargex") = 0

        p_nEditMode = xeEditMode.MODE_ADDNEW
        Return True
    End Function

    Function UpdateRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False
        If IsNothing(p_oDT) Then Return False
        If p_oDT.Rows.Count = 0 Then Return False

        p_nEditMode = xeEditMode.MODE_UPDATE

        Return True
    End Function

    Function CancelUpdate() As Boolean
        p_nEditMode = xeEditMode.MODE_READY

        Return True
    End Function

    Function SaveRecord() As Boolean
        Dim lsSQL As String

        Dim lsProcName As String = pxeModuleName & "." & "SaveRecord"

        If p_nEditMode <> xeEditMode.MODE_ADDNEW And p_nEditMode <> xeEditMode.MODE_UPDATE Then Return False

        Try
            If Not isEntryOK() Then Return False

            If p_nEditMode = xeEditMode.MODE_ADDNEW Then
                lsSQL = ADO2SQL(p_oDT, pxeTableNamex, , p_oApp.UserID, p_oApp.SysDate)
            Else
                If p_bModified = False Then GoTo endProc

                lsSQL = "UPDATE " & pxeTableNamex & " SET" & _
                                "  sAccredtn = " & strParm(p_oDT(0)("sAccredtn")) & _
                                ", sApproval = " & strParm(p_oDT(0)("sApproval")) & _
                                ", sPermitNo = " & strParm(p_oDT(0)("sPermitNo")) & _
                                ", nPOSNumbr = " & strParm(p_oDT(0)("nPOSNumbr")) & _
                                ", sORNoxxxx = " & strParm(p_oDT(0)("sORNoxxxx")) & _
                                ", nSalesTot = " & p_oDT(0)("nSalesTot") & _
                                ", cTranMode = " & strParm(p_oDT(0)("cTranMode")) & _
                                ", dExpiratn = " & dateParm(p_oDT(0)("dExpiratn")) & _
                                ", nSChargex = " & p_oDT(0)("nSChargex") & _
                                ", sModified = " & strParm(p_oApp.UserID) & _
                                ", dModified = " & datetimeParm(p_oApp.SysDate) & _
                        " WHERE sIDNumber = " & strParm(p_oDT(0)("sIDNumber"))
            End If
            p_oApp.BeginTransaction()
            If p_oApp.Execute(lsSQL, pxeTableNamex) = 0 Then GoTo endWithRoll
        Catch ex As Exception
            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")
            GoTo endWithRoll
        End Try

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        p_oApp.CommitTransaction()
endProc:
        If p_bShowMsgx Then MsgBox("Record Saved Successfully.", MsgBoxStyle.Information, "Success")
        Return True
endWithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Function DeleteRecord() As Boolean
        If p_nEditMode <> xeEditMode.MODE_READY Then Return False
        Dim lsSQL As String

        Dim lsProcName As String = pxeModuleName & ".DeleteRecord"

        If IsNothing(p_oDT) Then Return False
        If p_oDT.Rows.Count = 0 Then Return False

        If p_bShowMsgx Then
            If MsgBox("Are you sure to delete this record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                Return False
            End If
        End If

        Try
            p_oApp.BeginTransaction()

            lsSQL = "DELETE FROM " & pxeTableNamex & " WHERE sIDNumber = " & strParm(p_oDT(0)("sIDNumber"))
            p_oApp.Execute(lsSQL, pxeTableNamex)
        Catch ex As Exception
            MsgBox(lsProcName & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Warning")
            GoTo endWithRoll
        End Try

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        p_oApp.CommitTransaction()

        If p_bShowMsgx Then MsgBox("Record Deleted Successfully.", MsgBoxStyle.Information, "Success")
endProc:
        Return InitRecord()
endWithRoll:
        p_oApp.RollBackTransaction()
        Return False
    End Function

    Function BrowseRecord() As Boolean
        Dim loRow As DataRow

        loRow = SearchMachine()
        If Not IsNothing(loRow) Then
            InitRecord()
            p_oDT.Rows.Add()
            p_oDT(0)("sIDNumber") = loRow(0)
            p_oDT(0)("sAccredtn") = loRow(1)
            p_oDT(0)("sApproval") = loRow(2)
            p_oDT(0)("sPermitNo") = loRow(3)
            p_oDT(0)("nPOSNumbr") = loRow(4)
            p_oDT(0)("sORNoxxxx") = loRow(5)
            p_oDT(0)("nSalesTot") = loRow(6)
            p_oDT(0)("cTranMode") = loRow(7)
            p_oDT(0)("dExpiratn") = loRow(8)
            p_oDT(0)("nSChargex") = loRow(9)

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        End If

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        Return False
    End Function

    Function SearchMachine(Optional ByVal lsValue As String = "", _
                         Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sIDNumber LIKE " & strParm(lsValue & "%"))
        Else
            lsSQL = AddCondition(getSQLBrowse, "nPOSNumbr LIKE " & strParm(lsValue & "%"))
        End If

        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            Return Nothing
        ElseIf loDT.Rows.Count = 1 Then
            Return loDT.Rows(0)
        Else
            Return KwikSearch(p_oApp _
                            , getSQLBrowse _
                            , True _
                            , lsValue _
                            , "sIDNumber»sAccredtn»sApproval»sPermitNo»nPOSNumbr" _
                            , "ID No»Accreditation»Approval»Permit No»POS No" _
                            , "" _
                            , "sIDNumber»sAccredtn»sApproval»sPermitNo»nPOSNumbr" _
                            , IIf(lbCode, 1, 5))
        End If
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" & _
                    "  sIDNumber" & _
                    ", sAccredtn" & _
                    ", sApproval" & _
                    ", sPermitNo" & _
                    ", nPOSNumbr" & _
                    ", sORNoxxxx" & _
                    ", nSalesTot" & _
                    ", cTranMode" & _
                    ", dExpiratn" & _
                    ", nSChargex" & _
                " FROM " & pxeTableNamex
    End Function

    Private Function isEntryOK() As Boolean
        If p_oDT(0)("sIDNumber") = "" Or _
            p_oDT(0)("sAccredtn") = "" Or _
            p_oDT(0)("sApproval") = "" Or _
            p_oDT(0)("sPermitNo") = "" Or _
            p_oDT(0)("nPOSNumbr") = "" Then Return False

        Return True
    End Function
End Class
