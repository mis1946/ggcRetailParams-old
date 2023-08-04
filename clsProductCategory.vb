'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Product Category
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

Public Class clsProductCategory
    Private Const pxeModuleName As String = "clsProductCategory"
    Private Const pxeTableNamex As String = "Product_Category"

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
                Case 0, 1, 2, 3, 4, 5, 6
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case 0, 1, 2, 3, 4, 5, 6
                    If IFNull(p_oDT(0)(Index), "") <> value Then
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
                Case "sCategrCd", "sDescript", "sImgePath", "cForwardx", "cPriority", "cRecdStat", "sPrntPath"
                    Master = p_oDT(0)(Index)
                Case Else
                    Master = ""
                    MsgBox(pxeModuleName & vbCrLf & "Field Index not set! - " & Index, MsgBoxStyle.Critical, "Warning")
            End Select
        End Get
        Set(value As Object)
            Select Case Index
                Case "sCategrCd", "sDescript", "sImgePath", "cForwardx", "cPriority", "cRecdStat", "sPrntPath"
                    If IFNull(p_oDT(0)(Index), "") <> value Then
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
                Case 0, 1, 2, 3, 4, 5, 6
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
        p_oDT.Columns.Add("sCategrCd", GetType(String)).MaxLength = 4
        p_oDT.Columns.Add("sDescript", GetType(String)).MaxLength = 32
        p_oDT.Columns.Add("sImgePath", GetType(String)).MaxLength = 128
        p_oDT.Columns.Add("cForwardx", GetType(String)).MaxLength = 1
        p_oDT.Columns.Add("cPriority", GetType(String)).MaxLength = 1
        p_oDT.Columns.Add("cRecdStat", GetType(String)).MaxLength = 1
        p_oDT.Columns.Add("sPrntPath", GetType(String)).MaxLength = 128

        p_nEditMode = xeEditMode.MODE_UNKNOWN

        Return True
    End Function

    Function NewRecord() As Boolean
        InitRecord()
        If IsNothing(p_oDT) Then Return False

        p_oDT.Rows.Add()
        p_oDT(0)("sCategrCd") = GetNextCode(pxeTableNamex, "sCategrCd", False, p_oApp.Connection, False)
        p_oDT(0)("sDescript") = ""
        p_oDT(0)("sImgePath") = ""
        p_oDT(0)("cForwardx") = "0"
        p_oDT(0)("cPriority") = "0"
        p_oDT(0)("cRecdStat") = "1"
        p_oDT(0)("sPrntPath") = ""

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

                lsSQL = "UPDATE " & pxeTableNamex & " SET" &
                                "  sDescript = " & strParm(p_oDT(0)("sDescript")) &
                                ", sImgePath = " & strParm(p_oDT(0)("sImgePath")) &
                                ", cForwardx = " & strParm(p_oDT(0)("cForwardx")) &
                                ", cPriority = " & strParm(p_oDT(0)("cPriority")) &
                                ", cRecdStat = " & strParm(p_oDT(0)("cRecdStat")) &
                                ", sPrntPath = " & strParm(p_oDT(0)("sPrntPath")) &
                                ", sModified = " & strParm(p_oApp.UserID) &
                                ", dModified = " & dateParm(p_oApp.SysDate) &
                        " WHERE sCategrCd = " & strParm(p_oDT(0)("sCategrCd"))
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

            lsSQL = "DELETE FROM " & pxeTableNamex & " WHERE sCategrCd = " & strParm(p_oDT(0)("sCategrCd"))
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

        loRow = SearchCategory()
        If Not IsNothing(loRow) Then
            InitRecord()
            p_oDT.Rows.Add()
            p_oDT(0)("sCategrCd") = loRow(0)
            p_oDT(0)("sDescript") = loRow(1)
            p_oDT(0)("sImgePath") = loRow(2)
            p_oDT(0)("cForwardx") = loRow(3)
            p_oDT(0)("cPriority") = loRow(4)
            p_oDT(0)("cRecdStat") = loRow(5)
            p_oDT(0)("sPrntPath") = loRow(6)

            p_nEditMode = xeEditMode.MODE_READY
            Return True
        End If

        p_nEditMode = xeEditMode.MODE_UNKNOWN
        Return False
    End Function

    Function SearchCategory(Optional ByVal lsValue As String = "", _
                            Optional ByVal lbCode As Boolean = False) As DataRow
        Dim lsSQL As String
        Dim loDT As DataTable

        If lbCode Then
            lsSQL = AddCondition(getSQLBrowse, "sCategrCd = " & strParm(lsValue))
        Else
            lsSQL = AddCondition(getSQLBrowse, "sDescript LIKE " & strParm(lsValue & "%"))
        End If

        Debug.Print(lsSQL)
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
                        , "sCategrCd»sDescript" _
                        , "Code»Description" _
                        , "" _
                        , "sCategrCd»sDescript" _
                        , IIf(lbCode, 1, 2))
        End If
    End Function

    Function GetCategory() As DataTable
        Return p_oApp.ExecuteQuery(getSQLBrowse)
    End Function

    Function GetCategory(ByVal lsValue As String,
                            Optional ByVal lbCode As Boolean = False) As DataTable

        If lsValue = "" Then Return Nothing

        If lbCode Then
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sCategrCd = " & strParm(lsValue)))
        Else
            Return p_oApp.ExecuteQuery(AddCondition(getSQLBrowse, "sDescript = " & strParm(lsValue)))
        End If
    End Function

    Private Function getSQLBrowse() As String
        Return "SELECT" &
                    "  sCategrCd" &
                    ", sDescript" &
                    ", sImgePath" &
                    ", cForwardx" &
                    ", cPriority" &
                    ", cRecdStat" &
                    ", sPrntPath" &
                " FROM " & pxeTableNamex &
                IIf(p_nRecdStat > -1, " WHERE cRecdStat = " & strParm(p_nRecdStat), "")
    End Function

    Private Function isEntryOK() As Boolean
        If p_oDT(0)("sCategrCd") = "" Or _
            p_oDT(0)("sDescript") = "" Then Return False

        Return True
    End Function
End Class
