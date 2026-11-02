Attribute VB_Name = "Module1"
Option Explicit

Public Sub SAP_Final_Solution_V2()

    Const FIRST_DATA_ROW As Long = 2

    Dim wsInput As Worksheet
    Dim SapGuiAuto As Object
    Dim SAPGUIApp As Object
    Dim Connection As Object
    Dim Session As Object

    Dim lastRow As Long
    Dim currentRow As Long

    Dim supplierNo As String
    Dim materialNo As String
    Dim purchasingOrg As String
    Dim plantNo As String
    Dim infoRecordNo As String

    Dim priceValue As Variant
    Dim sapPrice As String
    Dim statusText As String

    On Error GoTo SAPConnectionError

    'Use the worksheet that is active when the macro starts.
    Set wsInput = ActiveSheet

    'Connect to the first active SAP GUI session.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPGUIApp = SapGuiAuto.GetScriptingEngine

    If SAPGUIApp.Children.Count = 0 Then
        Err.Raise vbObjectError + 100, , _
                  "No active SAP connection was found."
    End If

    Set Connection = SAPGUIApp.Children(0)

    If Connection.Children.Count = 0 Then
        Err.Raise vbObjectError + 101, , _
                  "No active SAP session was found."
    End If

    Set Session = Connection.Children(0)

    On Error GoTo 0

    'Column layout:
    'A = Supplier number
    'B = Description
    'C = Material number
    'D = Purchasing organization
    'E = Plant
    'F = Info record number returned by SAP
    'G = Price
    'H = Processing status

    wsInput.Range("F1").Value = "inforecord number"
    wsInput.Range("H1").Value = "SAP Status"

    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then
        MsgBox "No input data was found.", _
               vbExclamation, "SAP Update"
        Exit Sub
    End If

    Session.FindById("wnd[0]").Maximize

    For currentRow = FIRST_DATA_ROW To lastRow

        supplierNo = Trim$(CStr(wsInput.Cells(currentRow, "A").Value))
        materialNo = Trim$(CStr(wsInput.Cells(currentRow, "C").Value))
        purchasingOrg = Trim$(CStr(wsInput.Cells(currentRow, "D").Value))
        plantNo = Trim$(CStr(wsInput.Cells(currentRow, "E").Value))
        priceValue = wsInput.Cells(currentRow, "G").Value

        wsInput.Cells(currentRow, "F").NumberFormat = "@"
        wsInput.Cells(currentRow, "F").Value = vbNullString

        wsInput.Cells(currentRow, "H").Value = "Processing..."
        wsInput.Cells(currentRow, "H").Interior.Pattern = xlNone

        'Validate the Excel input.
        If supplierNo = vbNullString Then
            MarkSAPRowErrorV2 wsInput, currentRow, _
                              "Skipped: Supplier number is blank"
            GoTo NextRecord
        End If

        If materialNo = vbNullString Then
            MarkSAPRowErrorV2 wsInput, currentRow, _
                              "Skipped: Material number is blank"
            GoTo NextRecord
        End If

        If purchasingOrg = vbNullString Then
            MarkSAPRowErrorV2 wsInput, currentRow, _
                              "Skipped: Purchasing organization is blank"
            GoTo NextRecord
        End If

        If plantNo = vbNullString Then
            MarkSAPRowErrorV2 wsInput, currentRow, _
                              "Skipped: Plant is blank"
            GoTo NextRecord
        End If

        If IsEmpty(priceValue) Or _
           Trim$(CStr(priceValue)) = vbNullString Or _
           Not IsNumeric(priceValue) Then

            MarkSAPRowErrorV2 wsInput, currentRow, _
                              "Skipped: Price is blank or invalid"
            GoTo NextRecord
        End If

        'Format the price for SAP.
        sapPrice = Format$(CDbl(priceValue), "0.00")

        On Error GoTo SAPRowError

        'Open ME12.
        Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/NME12"
        Session.FindById("wnd[0]/tbar[0]/btn[0]").Press

        WaitForSAPV2 Session, 30

        'Enter the Excel values in SAP.
        Session.FindById( _
            "wnd[0]/usr/ctxtEINA-LIFNR").Text = supplierNo

        Session.FindById( _
            "wnd[0]/usr/ctxtEINA-MATNR").Text = materialNo

        Session.FindById( _
            "wnd[0]/usr/ctxtEINE-EKORG").Text = purchasingOrg

        Session.FindById( _
            "wnd[0]/usr/ctxtEINE-WERKS").Text = plantNo

        'Press Enter to open the purchasing info record.
        Session.FindById("wnd[0]/tbar[0]/btn[0]").Press

        WaitForSAPV2 Session, 30

        'Stop this row if SAP displays an error.
        statusText = GetSAPStatusTextV2(Session)

        If GetSAPStatusTypeV2(Session) = "E" Then
            Err.Raise vbObjectError + 200, , statusText
        End If

        'Extract the info record number from SAP.
        infoRecordNo = Trim$(CStr( _
            Session.FindById( _
                "wnd[0]/usr/ctxtEINA-INFNR").Text))

        If infoRecordNo = vbNullString Then
            Err.Raise vbObjectError + 201, , _
                      "SAP did not return an info record number."
        End If

        'Write the extracted number into column F.
        wsInput.Cells(currentRow, "F").NumberFormat = "@"
        wsInput.Cells(currentRow, "F").Value = infoRecordNo
        wsInput.Cells(currentRow, "F").Interior.Color = _
            RGB(255, 255, 153)

        'Open the purchasing organization data.
        Session.FindById("wnd[0]/tbar[1]/btn[8]").Press

        WaitForSAPV2 Session, 30

        'Handle the popup shown in the SAP recording.
        If SAPWindowExistsV2(Session, "wnd[1]") Then
            Session.FindById("wnd[1]/tbar[0]/btn[7]").Press
            WaitForSAPV2 Session, 30
        End If

        'Enter the price from column G.
        With Session.FindById( _
            "wnd[0]/usr/tblSAPMV13ATCTRL_D0201/" & _
            "txtKONP-KBETR[2,0]")

            .Text = sapPrice
            .SetFocus
            .CaretPosition = Len(sapPrice)
        End With

        'Validate the entered price.
        Session.FindById("wnd[0]").SendVKey 0

        WaitForSAPV2 Session, 30

        statusText = GetSAPStatusTextV2(Session)

        If GetSAPStatusTypeV2(Session) = "E" Then
            Err.Raise vbObjectError + 202, , statusText
        End If

        'Save the SAP record.
        Session.FindById("wnd[0]").SendVKey 11

        WaitForSAPV2 Session, 30

        'Handle a confirmation dialog if one appears.
        If SAPWindowExistsV2(Session, "wnd[1]") Then

            If SAPControlExistsV2( _
                Session, _
                "wnd[1]/usr/btnSPOP-OPTION1") Then

                Session.FindById( _
                    "wnd[1]/usr/btnSPOP-OPTION1").Press

                WaitForSAPV2 Session, 30
            End If

        End If

        statusText = GetSAPStatusTextV2(Session)

        If GetSAPStatusTypeV2(Session) = "E" Then
            Err.Raise vbObjectError + 203, , statusText
        End If

        If statusText = vbNullString Then
            statusText = "Updated successfully"
        End If

        wsInput.Cells(currentRow, "H").Value = statusText
        wsInput.Cells(currentRow, "F").Interior.Color = _
            RGB(198, 239, 206)
        wsInput.Cells(currentRow, "H").Interior.Color = _
            RGB(198, 239, 206)

NextRecord:

        On Error GoTo 0
        DoEvents

    Next currentRow

    MsgBox "SAP processing completed." & vbCrLf & _
           "Info record numbers are in column F." & vbCrLf & _
           "Processing results are in column H.", _
           vbInformation, "SAP Final Solution V2"

    Exit Sub

SAPRowError:

    MarkSAPRowErrorV2 wsInput, currentRow, _
        "Error " & Err.Number & ": " & Err.Description

    Err.Clear
    Resume NextRecord

SAPConnectionError:

    MsgBox "Unable to connect to SAP." & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Open SAP, log in, and confirm that " & _
           "SAP GUI scripting is enabled.", _
           vbCritical, "SAP Connection Error"

End Sub

Private Sub WaitForSAPV2( _
    ByVal Session As Object, _
    ByVal timeoutSeconds As Long)

    Dim startTime As Date
    Dim elapsedSeconds As Double

    startTime = Now

    Do
        DoEvents

        If Not Session.Busy Then Exit Do

        elapsedSeconds = DateDiff( _
            "s", startTime, Now)

        If elapsedSeconds >= timeoutSeconds Then
            Err.Raise vbObjectError + 300, , _
                      "SAP did not respond within " & _
                      timeoutSeconds & " seconds."
        End If
    Loop

End Sub

Private Function SAPWindowExistsV2( _
    ByVal Session As Object, _
    ByVal windowId As String) As Boolean

    Dim sapWindow As Object

    On Error Resume Next
    Set sapWindow = Session.FindById(windowId, False)
    SAPWindowExistsV2 = Not sapWindow Is Nothing
    Set sapWindow = Nothing
    On Error GoTo 0

End Function

Private Function SAPControlExistsV2( _
    ByVal Session As Object, _
    ByVal controlId As String) As Boolean

    Dim sapControl As Object

    On Error Resume Next
    Set sapControl = Session.FindById(controlId, False)
    SAPControlExistsV2 = Not sapControl Is Nothing
    Set sapControl = Nothing
    On Error GoTo 0

End Function

Private Function GetSAPStatusTextV2( _
    ByVal Session As Object) As String

    On Error Resume Next
    GetSAPStatusTextV2 = Trim$(CStr( _
        Session.FindById("wnd[0]/sbar").Text))
    On Error GoTo 0

End Function

Private Function GetSAPStatusTypeV2( _
    ByVal Session As Object) As String

    On Error Resume Next
    GetSAPStatusTypeV2 = UCase$(Trim$(CStr( _
        Session.FindById("wnd[0]/sbar").MessageType)))
    On Error GoTo 0

End Function

Private Sub MarkSAPRowErrorV2( _
    ByVal wsInput As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal errorMessage As String)

    wsInput.Cells(rowNumber, "H").Value = errorMessage
    wsInput.Cells(rowNumber, "H").Interior.Color = _
        RGB(255, 199, 206)

End Sub

