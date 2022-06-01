Imports FilterPalletAssy.commonClasses, System.Data.SqlClient, System.DirectoryServices

Public Class WebForm1
    Inherits System.Web.UI.Page

    Dim svr As New commonClasses.sqlInterface
    Dim emailMessage As New commonClasses.emailInterface
    Dim emailSubject As String
    Dim emailBody As String
    Dim countValid As Boolean = False
    Dim myHost As String = System.Net.Dns.GetHostName
    Dim verifyNext As Boolean = False
    Dim objRandom As New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer))

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblError.Text = ""
        If Not Page.IsPostBack Then
            addAttributes()
			' lblUserName.Text = System.Net.Dns.GetHostEntry(Request.UserHostName).HostName
		End If
    End Sub

    'SUBS / FUNCTIONS

    Private Sub addAttributes()
        txtFilterLot.Attributes.Add("NAME", "LOT")
        txtWeek.Attributes.Add("NAME", "WEEK")
        txtYear.Attributes.Add("NAME", "YEAR")
        txtPallet.Attributes.Add("NAME", "PALLET")
        lblCurrCount.Attributes.Add("NAME", "QTY")
        lblCompletedOn.Attributes.Add("NAME", "COMPLETED_ON")
    End Sub
    Private Function addBarcodeToPallet(ByVal barcode As String) As Boolean
        Dim params(5) As SqlParameter

        Try
            params(0) = New SqlParameter("@YEAR", txtYear.Text)
            params(1) = New SqlParameter("@WEEK", txtWeek.Text)
            params(2) = New SqlParameter("@PALLET", txtPallet.Text)
            params(3) = New SqlParameter("@BARCODE", barcode)
            params(4) = New SqlParameter("@FILTER_LOT", txtFilterLot.Text)
            params(5) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_ADD_BARCODE", params)

            Return True
        Catch ex As Exception
            showError(ex.Message)
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : addBarcodeToPallet "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Return False
    End Function
    Private Sub addBarcodeStringToPallet()
        Dim params(5) As SqlParameter

        Try
            params(0) = New SqlParameter("@YEAR", txtYear.Text)
            params(1) = New SqlParameter("@WEEK", txtWeek.Text)
            params(2) = New SqlParameter("@PALLET", txtPallet.Text)
            params(3) = New SqlParameter("@BARCODE", txtSerial.Text)
            params(4) = New SqlParameter("@FILTER_LOT", txtFilterLot.Text)
            params(5) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_ADD_BARCODE_STRING", params)

        Catch ex As Exception
            showError(ex.Message)
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : addBarcodeStringToPallet "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub
    Private Function checkFilterLot() As Boolean
        Dim params(1) As SqlParameter

        Try
            If ddlLoc.SelectedValue <> txtFilterLot.Text.Substring(8, 1) Then
                showError("LOT DOESN'T MATCH LINE")
                txtFilterLot.Text = ""
                Return False
            Else
                params(0) = New SqlParameter("@LOT_NUM", txtFilterLot.Text)
                params(1) = New SqlParameter("@ReturnValue", -1)
                params(1).Direction = ParameterDirection.InputOutput

                svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_CHECK_FILTER_LOT", params)

                If params(params.Length - 1).SqlValue.ToString = 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
            
        Catch ex As Exception
            showError("ERROR WITH INPUT")
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : checkFilterLot"
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            Return False
        End Try
    End Function
    Private Function checkPallet() As Boolean
        Dim params(6) As SqlParameter

        Try
            params(0) = New SqlParameter("@palletYear", txtYear.Text)
            params(1) = New SqlParameter("@palletWeek", txtWeek.Text)
            params(2) = New SqlParameter("@palletNum", txtPallet.Text)
            params(2).Direction = ParameterDirection.InputOutput
            params(2).Size = 10
            params(3) = New SqlParameter("@palletType", ddlMode.SelectedValue)
            params(4) = New SqlParameter("@LINE", ddlLoc.SelectedValue)
            params(5) = New SqlParameter("@BARCODE_COUNT", -1)
            params(5).Direction = ParameterDirection.InputOutput
            params(6) = New SqlParameter("@ReturnValue", -1)
            params(6).Direction = ParameterDirection.InputOutput

            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_CHECK_NEW_PALLET", params)

            If params(params.Length - 1).SqlValue.ToString = 1 Then
                If params(params.Length - 5).SqlValue.ToString <> "-1" Then
                    txtPallet.Text = params(params.Length - 5).SqlValue.ToString
                    Return True
                Else
                    showError("ERROR CREATING PALLET NUMBER!!")
                    Return False
                End If
            Else
                If params(params.Length - 2).SqlValue.ToString = 0 Then
                    ' PALLET EXISTS BUT IS EMPTY
                    txtPallet.Text = params(params.Length - 5).SqlValue.ToString
                    Return True
                Else
                    showError("ERROR CREATING PALLET NUMBER!!")
                    Return False
                End If
            End If

        Catch ex As Exception
            showError("ERROR WITH INPUT")
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : checkPallet "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            Return False
        End Try
    End Function
    'Private Sub configurePrinters()
    '    Dim printerList As New DataSet
    '    Dim currIndex As Integer
    '    Dim aList As ArrayList
    '    Dim dr As DataRow
    '    Dim defaultPrinterFound As Boolean = False
    '    Dim params(3) As SqlParameter
    '    Dim params2(0) As SqlParameter
    '    Dim params3(4) As SqlParameter
    '    Dim line As Integer = 0

    '    Try
    '        lblDefaultPersonalPrinter.Visible = True
    '        chkDefaultPersonal.Visible = True

    '        If myHost = "NA-CAD-WIN7-DEV1" Then
    '            ddlPrinters.DataSource = System.Drawing.Printing.PrinterSettings.InstalledPrinters
    '            ddlPrinters.DataBind()
    '        Else
    '            line += 1

    '            params2(0) = New SqlParameter("@APP_NAME", "FilterPalletAssy")

    '            printerList = svr.executeStoredProc(Application("sqlStrAS1APS"), "APPLICATION_GET_PRINTERS", params2, "PRINTER_LIST")
    '            ddlPrinters.DataSource = printerList.Tables("PRINTER_LIST")
    '            ddlPrinters.DataTextField = "printerName"
    '            ddlPrinters.DataValueField = "printerShare"
    '            ddlPrinters.DataBind()
    '            line += 1

    '            'Select Case ddlLoc.SelectedValue
    '            '    Case 1
    '            '        params(0) = New SqlParameter("@STATION", "FILTER PLT LINE1")
    '            '    Case 2
    '            '        params(0) = New SqlParameter("@STATION", "FILTER PLT LINE2")
    '            '    Case Else
    '            '        params(0) = New SqlParameter("@STATION", "FILTER PLT LINE1")
    '            'End Select

    '            'params(1) = New SqlParameter("@PRINTER", SqlDbType.VarChar)
    '            'params(1).Size = 200
    '            'params(1).Direction = ParameterDirection.Output
    '            'params(2) = New SqlParameter("@TRAY", SqlDbType.VarChar)
    '            'params(2).Size = 200
    '            'params(2).Direction = ParameterDirection.Output
    '            'params(3) = New SqlParameter("@APP_NAME", "FilterPalletAssy")


    '            params3(0) = New SqlParameter("@STATION", "")
    '            params3(1) = New SqlParameter("@PRINTER", SqlDbType.VarChar)
    '            params3(1).Size = 200
    '            params3(1).Direction = ParameterDirection.Output
    '            params3(2) = New SqlParameter("@TRAY", SqlDbType.VarChar)
    '            params3(2).Size = 200
    '            params3(2).Direction = ParameterDirection.Output
    '            params3(3) = New SqlParameter("@USERNAME", lblUserName.Text)
    '            params3(4) = New SqlParameter("@APP_NAME", "FilterPalletAssy")

    '            aList = svr.executeStoredProcNonQuery(Application("sqlStrAS1APS"), "APPLICATION_GET_DEFAULT_PRINTER_PERSONAL", params3)

    '            'If (aList.Item(1) Is DBNull.Value) Then
    '            '    aList = svr.executeStoredProcNonQuery(Application("sqlStrAS1APS"), "APPLICATION_GET_DEFAULT_PRINTER", params)
    '            'End If

    '            currIndex = 0
    '            If Not (aList.Item(1) Is DBNull.Value) Then
    '                For Each dr In printerList.Tables("PRINTER_LIST").Rows
    '                    If aList.Item(1) = dr.Item(0) Then
    '                        ddlPrinters.SelectedIndex = currIndex
    '                        defaultPrinterFound = True
    '                    Else
    '                        currIndex += 1
    '                    End If
    '                Next
    '            End If
    '        End If
    '    Catch ex As Exception
    '        showError(ex.Message)
    '        emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : configurePrinters "
    '        emailBody = ex.Message
    '        emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
    '    End Try

    'End Sub
    Private Sub deleteAllSerials()
        Dim params(3) As SqlParameter

        Try
            params(0) = New SqlParameter("@YEAR", txtYear.Text)
            params(1) = New SqlParameter("@WEEK", txtWeek.Text)
            params(2) = New SqlParameter("@PALLET", txtPallet.Text)
            params(3) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_DELETE_ALL", params)

            resetPage()
            showError("PALLET DELETED")
            grdDupBarcodes.Visible = False
        Catch ex As Exception
            showError("DELETE PALLET FAILED!")
            emailSubject = "APPLICATION ERROR: PALLET PROGRAM : btnDeleteAll_Click "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub

    Private Function deleteSerialByScan(ByRef errMsg As String) As Boolean
        Dim params(4) As SqlParameter
        Dim i As Integer
        Dim serialFound As Boolean = False
        Dim ds As New DataSet
        Dim barcodeArray() As String
        Dim barcode As String
        Dim tempSerial As String = ""
        Dim barcodeIndex As Integer = -1
        Dim lineNum As Integer = 0

        Try
            lineNum += 1
            For i = 1 To lstSerials.Items.Count
                If lstSerials.Items(i - 1).Text = txtSerialByScan.Text Then
                    serialFound = True
                    barcodeIndex = i - 1
                End If
            Next i

            lineNum += 1
            If lstSerials.Items(barcodeIndex).Text.Contains("S") Then
                barcodeArray = lstSerials.Items(barcodeIndex).Text.Split("S")

                For x = 1 To barcodeArray.Length - 1
                    barcodeArray(x) = barcodeArray(x).Substring(0, 10)
                Next
            Else
                barcodeArray = txtSerialByScan.Text.Split("*")
            End If

            lineNum += 1
            If (serialFound) Then
                For Each barcode In barcodeArray
                    removeBarcodeFromPallet(barcode)
                Next

                params(0) = New SqlParameter("@BARCODE", txtSerialByScan.Text)
                params(1) = New SqlParameter("@YEAR", txtYear.Text)
                params(2) = New SqlParameter("@WEEK", txtWeek.Text)
                params(3) = New SqlParameter("@PALLET", txtPallet.Text)
                params(4) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

                svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_DELETE_BARCODE_STRING", params)
                lineNum += 1
                lstSerials.Items.RemoveAt(barcodeIndex)

                refreshList()
                lineNum += 1
                If lstSerials.Items.Count = 0 Then
                    btnDelete.Enabled = False
                    btnDeleteAll.Enabled = False
                    btnDeleteByScan.Enabled = False
                    resetPage()
                End If
                lineNum += 1
                Return True
            Else
                errMsg = "BARCODE NOT FOUND IN LIST"
                Return False
            End If
        Catch ex As Exception
            errMsg = "ERROR WITH INPUT"
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : deleteSerialByScan - " & txtSerialByScan.Text
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Return False
    End Function


    Private Sub displaySerialPanel()
        Dim errMsg As String = ""
        Dim params(5) As SqlParameter
        Dim params2(3) As SqlParameter
        Dim params3(3) As SqlParameter
        Dim ds As New DataSet
        Dim dr As DataRow
        Dim barcodeArray() As String

        Try
            If Page.IsValid Then
                If pnlVerify.Visible = True Then

                Else
                    If btnResume.Visible = True Then
                        pnlPalletHeader.Enabled = False

                        params2(0) = New SqlParameter("@YEAR", txtYear.Text)
                        params2(1) = New SqlParameter("@WEEK", txtWeek.Text)
                        params2(2) = New SqlParameter("@PALLET", txtPallet.Text)
                        params2(3) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

                        ds = svr.executeStoredProc(Application("sqlStrFILTER"), "FILTER_PALLET_GET_PALLET_LIST", params2, "SERIALS")

                        If ds.Tables(0).Rows.Count > 0 Then
                            If ds.Tables(0).Rows(0).Item("filterLotNum").ToString.Substring(8, 1) = 1 Then
                                ddlLoc.SelectedValue = 1
                                txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                            Else
                                If ds.Tables(0).Rows(0).Item("filterLotNum").ToString.Substring(8, 1) = 3 Then
                                    ddlLoc.SelectedValue = 3
                                    txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                                Else
                                    ddlLoc.SelectedValue = 2
                                    txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                                End If
                            End If

                            Select Case ds.Tables(0).Rows(0).Item("palletID").ToString.Substring(0, 1)
                                Case "A"
                                    ddlMode.SelectedValue = 1
                                Case "S"
                                    ddlMode.SelectedValue = 2
                                Case "Q"
                                    ddlMode.SelectedValue = 3
                                Case "L"
                                    ddlMode.SelectedValue = 4
                                Case "I"
                                    ddlMode.SelectedValue = 5
                            End Select

                            ddlLoc.Visible = True
                            ddlMode.Visible = True
                            ddlLoc.Enabled = False
                            ddlMode.Enabled = False
                            pnlSerialData.Visible = True
                            pnlSerialData.Enabled = True

                            For Each dr In ds.Tables(0).Rows
                                lstSerials.Items.Add(dr.Item(5))
                                barcodeArray = dr.Item(5).ToString.Split("*")
                                If ddlMode.SelectedValue <> 4 And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                                    lblCurrCount.Text = lblCurrCount.Text + (barcodeArray.Length / 2)
                                Else
                                    lblCurrCount.Text = lblCurrCount.Text + (barcodeArray.Length)
                                End If
                            Next

                            If ddlMode.SelectedValue = 1 Or ddlMode.SelectedValue = 5 Then
                                If ddlLoc.SelectedValue = 3 Then
                                    lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 42) * 42)
                                Else
                                    lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 35) * 35)
                                End If
                            End If

                            ' CHECK FOR QTY IN PALLET
                            If lstSerials.Items.Count = 18 And ddlMode.SelectedValue = 2 And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                                txtSerial.Enabled = False
                                btnDelete.Enabled = True
                                btnDeleteAll.Enabled = True
                                btnDeleteByScan.Enabled = True
                                showError("PALLET COMPLETE:  PLEASE CLICK FINISH")
                            Else
                                btnDelete.Enabled = True
                                btnDeleteAll.Enabled = True
                                btnDeleteByScan.Enabled = True
                            End If
                        Else
                            showError("EMPTY OR INVALID PALLET")
                            params3(0) = New SqlParameter("@YEAR", txtYear.Text)
                            params3(1) = New SqlParameter("@WEEK", txtWeek.Text)
                            params3(2) = New SqlParameter("@PALLET", txtPallet.Text)
                            params3(3) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

                            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_DELETE_PALLET", params3)
                            pnlSerialData.Enabled = False
                        End If
                    Else
                        pnlPalletHeader.Enabled = False
                        pnlSerialData.Enabled = True
                        pnlSerialData.Visible = True
                        btnDeleteByScan.Enabled = True
                    End If
                End If

            End If
        Catch ex As Exception
            emailSubject = "APPLICATION ERROR: PALLET PROGRAM : btnSubmit_Click "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Page.SetFocus("txtSerial")
    End Sub
    Private Function getFieldText(ByVal fieldName As String) As String
        Dim c As System.Web.UI.Control
        Dim txtBox As System.Web.UI.WebControls.TextBox
        Dim lblBox As System.Web.UI.WebControls.Label
        Dim fieldText As String

        Try
            For Each c In pnlPalletHeader.Controls
                Select Case c.GetType.ToString
                    Case "System.Web.UI.WebControls.TextBox"
                        txtBox = c
                        If txtBox.Attributes("NAME") = fieldName Then
                            fieldText = txtBox.Text
                            Return fieldText
                        End If
                End Select
            Next

            For Each c In pnlSerialData.Controls
                Select Case c.GetType.ToString
                    Case "System.Web.UI.WebControls.TextBox"
                        txtBox = c
                        If txtBox.Attributes("NAME") = fieldName Then
                            fieldText = txtBox.Text
                            Return fieldText
                        End If
                    Case "System.Web.UI.WebControls.Label"
                        lblBox = c
                        If lblBox.Attributes("NAME") = fieldName Then
                            fieldText = lblBox.Text
                            Return fieldText
                        End If
                End Select
            Next
        Catch ex As Exception
            emailSubject = "APPLICATION ERROR: FILTER PALLET ASSY : getFieldText"
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Return Nothing
    End Function
    Private Function getFilterLot(ByVal barcode As String) As Boolean
        Dim params(1) As SqlParameter

        Try
            params(0) = New SqlParameter("@BARCODE", barcode)
            params(1) = New SqlParameter("@FILTER_LOT", "AVO0000000")
            params(1).Direction = ParameterDirection.InputOutput

            If ddlLoc.SelectedValue = 3 Then
                svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_GET_LOT_OFFLINE", params)
            Else
                svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_GET_LOT", params)
            End If



            If params(params.Length - 1).SqlValue.ToString = Nothing Then
                Return False
            Else
                If params(params.Length - 1).SqlValue.ToString = "Null" Then
                    Return False
                Else
                    txtFilterLot.Text = params(params.Length - 1).SqlValue.ToString
                    Return True
                End If
                
            End If
        Catch ex As Exception
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : getFilterLot "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            Return False
        End Try

    End Function
    Private Function getLabelTemplate(ByVal partTypeCode As String) As DataSet
        Dim ds As DataSet
        Dim params(1) As SqlParameter
        Dim currPage As Integer = 1

        Try

            params(0) = New SqlParameter("@LABEL_TYPE", partTypeCode)
            params(1) = New SqlParameter("@PAGE", currPage)

            ds = svr.executeStoredProc(Application("sqlStrAS1APS"), "LABEL_GET_TEMPLATE_LAYOUT", params, "LABEL_TEMPLATE")

            If ds Is Nothing Then
                Return Nothing
            Else
                If ds.Tables("LABEL_TEMPLATE").Rows.Count > 0 Then
                    Return ds
                Else
                    Return Nothing
                End If
            End If
        Catch ex As Exception
            emailSubject = "APPLICATION ERROR: FILTER PALLET ASSY: getLabelTemplate "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Return Nothing
    End Function
    Private Sub getNextPallet()
        Dim ds As New DataSet
        Dim params(5) As SqlParameter

        Try
            If Page.IsValid() Then
                If pnlVerify.Visible = False Then
                    params(0) = New SqlParameter("@palletYear", txtYear.Text)
                    params(1) = New SqlParameter("@palletWeek", txtWeek.Text)
                    params(2) = New SqlParameter("@palletNum", "XXXX")
                    params(2).Direction = ParameterDirection.InputOutput
                    params(3) = New SqlParameter("@palletType", ddlMode.SelectedValue)
                    params(4) = New SqlParameter("@line", ddlLoc.SelectedValue)
                    params(5) = New SqlParameter("@ReturnValue", -1)
                    params(5).Direction = ParameterDirection.InputOutput

                    svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_GET_NEW_PALLET", params)

                    If params(params.Length - 1).SqlValue.ToString = 1 Then
                        If params(params.Length - 4).SqlValue.ToString <> "-1" Then
                            txtPallet.Text = params(params.Length - 4).SqlValue.ToString
                            displaySerialPanel()
                        Else
                            showError("ERROR GETTING NEXT PALLET NUMBER!!")
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
            showError("ERROR WITH INPUT")
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : getNextPallet "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub
    Public Function getRandomNumber(Optional ByVal Low As Integer = 1, Optional ByVal High As Integer = 100) As Integer
        ' Returns a random number,
        ' between the optional Low and High parameters
        Return objRandom.Next(Low, High + 1)
    End Function
    Private Function getSerialPosition(ByVal barcode As String) As Integer
        Dim ds As New DataSet
        Dim params(0) As SqlParameter
        Dim position As Integer
        Dim dr As DataRow

        Try

            params(0) = New SqlParameter("BARCODE", barcode)

            ds = svr.executeStoredProc(Application("sqlStrFILTER"), "FILTER_PALLET_GET_BOX_POSITION", params, "POSITION")

            If ds Is Nothing Then
                Return -1
            Else
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        If dr.Item("BARCODE_STRING").ToString.Contains(barcode) Then
                            position = dr.Item("RECORD_ID")
                            Return position
                        End If
                    Next
                Else
                    Return -1
                End If
            End If

            Return -1
        Catch ex As Exception
            Return -1
        End Try
    End Function
    Private Sub processEmptyFields(ByVal pnl As Panel)
        Dim c As Control
        Dim txtBox As TextBox

        For Each c In pnl.Controls
            Select Case c.GetType.ToString
                Case "System.Web.UI.WebControls.TextBox"
                    txtBox = c
                    If txtBox.Text = "" And txtBox.Visible = True Then
                        txtBox.Enabled = True
                        If txtBox.ID = "txtPallet" And btnCreate.Visible = True Then
                            If ddlMode.SelectedValue = 1 Then
                                txtPallet.Enabled = True
                                reqFieldValPallet.Enabled = True
                                txtYear.Enabled = False
                                Page.SetFocus(txtBox.ID)
                            Else
                                Page.Validate()
                                getNextPallet()
                                txtPallet.Enabled = False
                            End If
                        Else
                            Page.SetFocus(txtBox.ID)

                            Select Case txtBox.ID
                                Case "txtWeek"
                                    reqFieldValWeek.Enabled = True
                                    cstValWeek.Enabled = True
                                Case "txtYear"
                                    reqFieldValYear.Enabled = True
                                    rngValYear.Enabled = True
                                    txtWeek.Enabled = False
                                Case "txtPallet"
                                    reqFieldValPallet.Enabled = True
                                    txtYear.Enabled = False
                                Case "txtFilterLot"
                                    If btnResume.Visible = True Then
                                        txtFilterLot.Text = "."
                                        txtFilterLot.Enabled = False
                                        processEmptyFields(pnlPalletHeader)
                                    Else
                                        If ddlLoc.SelectedValue = 2 Or ddlLoc.SelectedValue = 3 Then
                                            txtFilterLot.Text = "."
                                            txtFilterLot.Enabled = False
                                            processEmptyFields(pnlPalletHeader)
                                        End If
                                    End If

                            End Select
                        End If

                        Exit For
                    End If
            End Select
        Next

    End Sub
    Private Sub refreshList()
        Dim params2(3) As SqlParameter
        Dim ds As New DataSet
        Dim barcodeArray() As String

        Try

            lstSerials.Items.Clear()
            lblCurrCount.Text = 0

            params2(0) = New SqlParameter("@YEAR", txtYear.Text)
            params2(1) = New SqlParameter("@WEEK", txtWeek.Text)
            params2(2) = New SqlParameter("@PALLET", txtPallet.Text)
            params2(3) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

            ds = svr.executeStoredProc(Application("sqlStrFILTER"), "FILTER_PALLET_GET_PALLET_LIST", params2, "SERIALS")

            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("filterLotNum").ToString.Substring(8, 1) = 1 Then
                    ddlLoc.SelectedValue = 1
                    txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                Else
                    If ds.Tables(0).Rows(0).Item("filterLotNum").ToString.Substring(8, 1) = 3 Then
                        ddlLoc.SelectedValue = 3
                        txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                    Else
                        ddlLoc.SelectedValue = 2
                        txtFilterLot.Text = ds.Tables(0).Rows(0).Item("filterLotNum").ToString
                    End If
                End If

                Select Case ds.Tables(0).Rows(0).Item("palletID").ToString.Substring(0, 1)
                    Case "A"
                        ddlMode.SelectedValue = 1
                    Case "S"
                        ddlMode.SelectedValue = 2
                    Case "Q"
                        ddlMode.SelectedValue = 3
                    Case "L"
                        ddlMode.SelectedValue = 4
                    Case "I"
                        ddlMode.SelectedValue = 5
                End Select

                ddlLoc.Visible = True
                ddlMode.Visible = True
                ddlLoc.Enabled = False
                ddlMode.Enabled = False
                pnlSerialData.Visible = True
                pnlSerialData.Enabled = True

                For Each dr In ds.Tables(0).Rows
                    lstSerials.Items.Add(dr.Item(5))
                    barcodeArray = dr.Item(5).ToString.Split("*")
                    If ddlMode.SelectedValue <> 4 And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                        lblCurrCount.Text = lblCurrCount.Text + (barcodeArray.Length / 2)
                    Else
                        lblCurrCount.Text = lblCurrCount.Text + (barcodeArray.Length)
                    End If
                Next

                ' CHECK FOR QTY IN PALLET
                If lstSerials.Items.Count = 18 And ddlMode.SelectedValue = 2 And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                    txtSerial.Enabled = False
                    btnDelete.Enabled = True
                    btnDeleteAll.Enabled = True
                    btnDeleteByScan.Enabled = True
                Else
                    btnDelete.Enabled = True
                    btnDeleteAll.Enabled = True
                    btnDeleteByScan.Enabled = True
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub removeBarcodeFromPallet(ByVal barcode As String)
        Dim params(4) As SqlParameter

        Try
            params(0) = New SqlParameter("@BARCODE", barcode)
            params(1) = New SqlParameter("@YEAR", txtYear.Text)
            params(2) = New SqlParameter("@WEEK", txtWeek.Text)
            params(3) = New SqlParameter("@PALLET", txtPallet.Text)
            params(4) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

            svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_DELETE_BARCODE", params)
        Catch ex As Exception
            showError(ex.Message)
            emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : removeBarcodeFromPallet "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub
    Private Sub resetPage()
        btnResume.Visible = True
        btnResume.Enabled = True
        btnCreate.Visible = True
        btnCreate.Enabled = True
        ddlMode.Visible = False
        pnlPalletHeader.Visible = False
        pnlPalletHeader.Enabled = False
        ddlMode.SelectedIndex = 0
        ddlLoc.SelectedIndex = 0
        ddlLoc.Visible = False
        ddlLoc.Enabled = False
        cstValWeek.Enabled = False
        reqFieldValWeek.Enabled = False
        rngValYear.Enabled = False
        reqFieldValYear.Enabled = False
        reqFieldValPallet.Enabled = False
        btnSubmit.Visible = False
        pnlSerialData.Visible = False
        lstSerials.Items.Clear()
        txtSerial.Text = ""
        lblSerialCheck.Text = ""
        txtSerialByScan.Text = ""
        txtSerialByScan.Visible = False
        txtSerialByScan.Enabled = False
        txtSerial.Enabled = True
        txtSerial.Visible = True
        txtWeek.Text = ""
        txtYear.Text = ""
        txtPallet.Text = ""
        txtPallet.CssClass = "txtPallet"
        lblPallet.CssClass = "lblPallet"
        lblError.Text = ""
        lblCurrCount.Text = 0
        lblCurrRowCount.Text = 0
        btnDelete.Enabled = True
        btnDeleteAll.Enabled = True
        btnDeleteByScan.Enabled = True
        txtSerial.Enabled = True
        pnlSerialData.Enabled = True
        pnlSignOff.Visible = False
        cstValCountVerify.Enabled = False
        reqFieldInitials.Enabled = False
        reqFieldCountVerify.Enabled = False
        countValid = False
        txtInitials.Text = ""
        txtCountVerify.Text = ""
        pnlVerify.Visible = False
        pnlVerifyDelete.Visible = False
        txtFilterLot.Text = ""
        txtFilterLot.Enabled = False
        lblCompletedOn.Text = ""
        btnDeleteByScan.BackColor = Drawing.Color.LightGray
        grdDupBarcodes.Visible = False
        'lblDefaultPersonalPrinter.Visible = False
        'chkDefaultPersonal.Visible = False

    End Sub
    Private Sub showError(ByVal errorMsg As String)
        lblError.Text = errorMsg
        lblError.Visible = True
    End Sub
    'Private Sub updatePersonalDefaultPrinter()
    '    Dim params(4) As SqlParameter

    '    Try
    '        params(0) = New SqlParameter("@STATION", "")
    '        params(1) = New SqlParameter("@PRINTER", ddlPrinters.SelectedItem.Text)
    '        params(2) = New SqlParameter("@TRAY", "")
    '        params(3) = New SqlParameter("@USERNAME", lblUserName.Text)
    '        params(4) = New SqlParameter("@APP_NAME", "FilterPalletAssy")

    '        svr.executeStoredProc(Application("sqlStrAS1APS"), "APPLICATION_INSERT_DEFAULT_PERSONAL_PRINTER", params, "PRINTER")

    '        chkDefaultPersonal.Checked = False
    '    Catch ex As Exception
    '        emailSubject = "APPLICATION ERROR: FILTER PALLET ASSY : updatePersonalDefaultPrinter"
    '        emailBody = ex.Message
    '        emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
    '    End Try
    'End Sub
    Private Function validateSerial(ByRef errMsg As String) As Boolean
        Dim params(3) As SqlParameter
        Dim params2(1) As SqlParameter
        Dim result As Integer
        Dim i As Integer
        Dim serialFound As Boolean = False
        Dim ds As New DataSet
        Dim barcodeArray() As String
        Dim barcode As String
        Dim lstPosition As Integer = 0
        Dim dt As New DataTable
        Dim dr As DataRow

        dt = New DataTable("BARCODES")
        dt.Columns.Add("BARCODE")
        dt.Columns.Add("BOX_POSITION")
        dt.Columns.Add("PALLET_NUM")

        Dim tempSerial As String = ""

        Try
            For i = 1 To lstSerials.Items.Count
                If lstSerials.Items(i - 1).Text = txtSerial.Text Then
                    serialFound = True
                    lstSerials.SelectedIndex = i - 1
                    lstPosition = i
                    Exit For
                End If
            Next i

            If txtSerial.Text.Contains("S") Then
                barcodeArray = txtSerial.Text.Split("S")

                For x = 1 To barcodeArray.Length - 1
                    barcodeArray(x) = barcodeArray(x).Substring(0, 10)
                Next
            Else
                barcodeArray = txtSerial.Text.Split("*")
            End If

            If Not (serialFound) Then
                For Each barcode In barcodeArray
                    If barcode.StartsWith("[") Then
                        Continue For
                    End If

                    If txtFilterLot.Text = "" Then

                        If Not (getFilterLot(barcode)) Then
                            errMsg = "ERROR GETTING LOT NUMBER"
                            Return False
                        End If
                    End If

                    params(0) = New SqlParameter("@BARCODE", barcode)
                    params(1) = New SqlParameter("@FILTER_LOT", txtFilterLot.Text)
                    params(2) = New SqlParameter("@ERROR_CODE", Nothing)
                    params(2).Direction = ParameterDirection.Output
                    params(3) = New SqlParameter("@ReturnValue", Nothing)
                    params(3).Direction = ParameterDirection.Output

                    ds = svr.executeStoredProc(Application("sqlStrFILTER"), "FILTER_PALLET_CHECK_BARCODEV2", params, "PALLET")

                    result = params(params.Length - 1).SqlValue.ToString

                    Select Case result
                        Case 1
                            errMsg = ""
                        Case 2
                            errMsg = "INVALID BARCODE OR SCRAP: " & barcode
                            Return False
                        Case 3
                            If ds.Tables(0).Rows(0).Item(0).ToString = txtPallet.Text And ds.Tables(0).Rows(0).Item(1).ToString = txtWeek.Text And ds.Tables(0).Rows(0).Item(2).ToString = txtYear.Text Then
                                refreshList()
                                errMsg = "BARCODE ALREADY SCANNED: BARCODE = " & barcode
                                For i = 1 To lstSerials.Items.Count
                                    If lstSerials.Items(i - 1).Text.Contains(barcode) Then
                                        lstSerials.SelectedIndex = i - 1
                                        lstPosition = i
                                        Exit For
                                    End If
                                Next i
                                dr = dt.NewRow()
                                dr("BARCODE") = barcode
                                dr("BOX_POSITION") = lstPosition
                                dr("PALLET_NUM") = ds.Tables(0).Rows(0).Item(0).ToString
                                dt.Rows.Add(dr)
                                serialFound = True
                                pnlDuplicateError.Visible = True
                                pnlSerialData.Visible = False
                            Else
                                errMsg = "SERIAL ON ANOTHER PALLET: " & ds.Tables(0).Rows(0).Item(0).ToString & " " & ds.Tables(0).Rows(0).Item(1).ToString & " " & ds.Tables(0).Rows(0).Item(2).ToString & " / BARCODE: " & barcode
                                dr = dt.NewRow()
                                dr("BARCODE") = barcode
                                dr("BOX_POSITION") = getSerialPosition(barcode)
                                If dr("BOX_POSITION").ToString = "-1" Then
                                    dr("BOX_POSITION") = "UNKNOWN"
                                End If
                                dr("PALLET_NUM") = ds.Tables(0).Rows(0).Item(0).ToString
                                dt.Rows.Add(dr)
                                serialFound = True
                                pnlDuplicateError.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        Case 4
                            errMsg = params(2).SqlValue.ToString & " - " & barcode
                            Return False
                    End Select
                Next

                If serialFound Then
                    If dt.Rows.Count > 0 Then
                        grdDupBarcodes.DataSource = dt
                        grdDupBarcodes.DataBind()
                        grdDupBarcodes.Visible = True
                    End If
                    Return False
                Else
                    Return True
                End If

            Else
                errMsg = "BARCODE ALREADY SCANNED: BOX POSITION ON PALLET = " & lstPosition
                pnlDuplicateError.Visible = True
                pnlSerialData.Visible = False
                Return False
            End If
        Catch ex As Exception
            errMsg = "ERROR WITH INPUT"
            'emailSubject = "APPLICATION ERROR: FILTER PALLET PROGRAM : validateSerial : input = " & txtSerial.Text
            'emailBody = ex.Message
            'emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

        Return False
    End Function
    Private Function validatePallet(ByRef errMsg As String) As Boolean
        'Dim params(5) As SqlParameter
        'Dim result As Integer

        '    Try
        '        params(0) = New SqlParameter("@YEAR", txtYear.Text)
        '        params(1) = New SqlParameter("@WEEK", txtWeek.Text)
        '        params(2) = New SqlParameter("@PALLET", txtPallet.Text)
        '        params(3) = New SqlParameter("@AR", txtARNum.Text)
        '        params(4) = New SqlParameter("@RECIPE", txtRecipe.Text)
        '        params(5) = New SqlParameter("@ReturnValue", Nothing)
        '        params(5).Direction = ParameterDirection.Output

        '        result = svr.executeStoredProcRtrnValue(Application("sqlStrMASK"), "PALLET_CHECK_PALLET", params)

        '        Select Case result
        '            Case 1
        '                errMsg = "PALLET EXISTS"
        '                Return False
        '            Case 2
        '                errMsg = "INVALID RECIPE"
        '                Return False
        '            Case 3
        '                errMsg = "EMPTY PALLET"
        '                Return True
        '            Case 0
        '                errMsg = ""
        '                Return True
        '        End Select
        '    Catch ex As Exception
        '        emailSubject = "APPLICATION ERROR: PALLET PROGRAM : validateSerial "
        '        emailBody = ex.Message
        '        emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        '    End Try

        Return False
    End Function

    ' EVENT HANDLERS

    'CHECKBOX
    'Protected Sub chkDefaultPersonal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDefaultPersonal.CheckedChanged
    '    If chkDefaultPersonal.Checked = True Then
    '        updatePersonalDefaultPrinter()
    '    End If
    'End Sub

    ' BUTTONS
    Protected Sub btnCreate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        btnResume.Visible = False
        btnCreate.Enabled = False
        ddlLoc.Enabled = True
        ddlLoc.Visible = True
    End Sub
    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim params(4) As SqlParameter
        Dim barcodeArray() As String
        Dim barcode As String

        Try
            If Not (lstSerials.SelectedItem Is Nothing) Then
                If lstSerials.SelectedItem.Text.Contains("S") Then
                    barcodeArray = lstSerials.SelectedItem.Text.Split("S")

                    For x = 1 To barcodeArray.Length - 1
                        barcodeArray(x) = barcodeArray(x).Substring(0, 10)
                    Next
                Else
                    barcodeArray = lstSerials.SelectedItem.Text.Split("*")
                End If


                For Each barcode In barcodeArray
                    removeBarcodeFromPallet(barcode)
                Next

                'If ddlMode.SelectedValue <> 4 Then
                '    If ddlMode.SelectedValue = 1 Or ddlMode.SelectedValue = 5 Then
                '        If ddlLoc.SelectedValue = 3 Then
                '            lblCurrCount.Text = lblCurrCount.Text - (barcodeArray.Length)
                '        Else
                '            lblCurrCount.Text = lblCurrCount.Text - (barcodeArray.Length / 2)
                '        End If

                '        If ddlMode.SelectedValue = 1 Or ddlMode.SelectedValue = 5 Then
                '            If ddlLoc.SelectedValue = 3 Then
                '                lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 42) * 42)
                '            Else
                '                lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 35) * 35)
                '            End If
                '        End If
                '    Else
                '        If ddlLoc.SelectedValue = 2 Then
                '            lblCurrCount.Text = lblCurrCount.Text - (barcodeArray.Length / 2)
                '        Else
                '            lblCurrCount.Text = lblCurrCount.Text - (barcodeArray.Length)

                '            refreshList()
                '        End If
                '    End If

                'Else
                '    lblCurrCount.Text = lblCurrCount.Text - (barcodeArray.Length)

                '    refreshList()
                'End If


                params(0) = New SqlParameter("@BARCODE", lstSerials.SelectedItem.Text)
                params(1) = New SqlParameter("@YEAR", txtYear.Text)
                params(2) = New SqlParameter("@WEEK", txtWeek.Text)
                params(3) = New SqlParameter("@PALLET", txtPallet.Text)
                params(4) = New SqlParameter("@LINE", ddlLoc.SelectedValue)

                svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_DELETE_BARCODE_STRING", params)
                'lblCurrCount.Text = CType(lblCurrCount.Text, Integer) - 1
                lstSerials.Items.RemoveAt(lstSerials.SelectedIndex)

                refreshList()

                If lstSerials.Items.Count = 0 Then
                    btnDelete.Enabled = False
                    btnDeleteAll.Enabled = False
                    btnDeleteByScan.Enabled = False
                    resetPage()
                End If

            Else
                showError("No Serial Selected")
            End If
            grdDupBarcodes.Visible = False
        Catch ex As Exception
            showError("DELETE SERIAL FAILED!")
            emailSubject = "APPLICATION ERROR: PALLET PROGRAM : btnDelete_Click "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

    End Sub
    Private Sub btnDeleteAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteAll.Click
        pnlVerifyDelete.Visible = True
        pnlSerialData.Visible = False
    End Sub
    Private Sub btnDeleteByScan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteByScan.Click
        If btnDeleteByScan.BackColor = Drawing.Color.Red Then
            btnDeleteByScan.BackColor = Drawing.Color.LightGray
            txtSerial.Enabled = True
            txtSerial.Visible = True
            txtSerialByScan.Visible = False
            txtSerialByScan.Enabled = False
            Page.SetFocus("txtSerial")
        Else
            btnDeleteByScan.BackColor = Drawing.Color.Red
            txtSerial.Enabled = False
            txtSerial.Visible = False
            txtSerialByScan.Visible = True
            txtSerialByScan.Enabled = True
            Page.SetFocus("txtSerialByScan")
        End If
        grdDupBarcodes.Visible = False
    End Sub
    Private Sub btnDupOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDupOK.Click
        pnlDuplicateError.Visible = False
        pnlSerialData.Visible = True
        Page.SetFocus("txtSerial")
        grdDupBarcodes.Visible = False
    End Sub
    Protected Sub btnResume_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnResume.Click
        ddlLoc.Visible = True
        ddlLoc.Enabled = True
    End Sub
    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        resetPage()
    End Sub
    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim barcodeArray() As String
        Dim barcode As String

        displaySerialPanel()

        If pnlVerify.Visible = True Then

        Else
            If lstSerials.Items.Count > 0 Then
                barcodeArray = lstSerials.Items(0).Text.Split("*")
                barcode = barcodeArray(0)
                getFilterLot(barcode)
            End If
            ' configurePrinters()
            'ddlPrinters.Visible = True
        End If
        
    End Sub
    Protected Sub btnFinish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinish.Click
        Dim params(3) As SqlParameter

        Try
            If pnlSignOff.Visible = True Then
                If countValid = True Then
                    params(0) = New SqlParameter("@YEAR", txtYear.Text)
                    params(1) = New SqlParameter("@WEEK", txtWeek.Text)
                    params(2) = New SqlParameter("@PALLET", txtPallet.Text)
                    params(3) = New SqlParameter("@INITIALS", txtInitials.Text)

                    svr.executeStoredProcOutput(Application("sqlStrFILTER"), "FILTER_PALLET_FINISH_PALLET", params)
                    lblCompletedOn.Text = System.DateTime.Now()
                    printDocument()
                    resetPage()
                End If
            Else
                pnlSignOff.Visible = True
                txtSerial.Enabled = False
                btnDelete.Enabled = False
                btnDeleteAll.Enabled = False
                btnDeleteByScan.Enabled = False
                reqFieldInitials.Enabled = True
                cstValCountVerify.Enabled = True
                reqFieldCountVerify.Enabled = True
                Page.SetFocus("txtInitials")
            End If
        Catch ex As Exception
            showError("FINISH FAILED")
            emailSubject = "APPLICATION ERROR: PALLET PROGRAM : btnFinish_Click "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try

    End Sub
    Protected Sub btnNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNo.Click
        resetPage()
    End Sub
    Protected Sub btnYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnYes.Click
        pnlSerialData.Visible = True
        pnlVerify.Visible = False
        cstValWeek.Enabled = False
        Dim barcodeArray() As String
        Dim barcode As String

        If btnResume.Visible = True Then
            displaySerialPanel()

            If lstSerials.Items.Count > 0 Then
                barcodeArray = lstSerials.Items(0).Text.Split("*")
                barcode = barcodeArray(0)
                getFilterLot(barcode)
            End If
            'configurePrinters()
            'ddlPrinters.Visible = True
        Else
            If ddlMode.SelectedValue = 1 Then
            Else
                getNextPallet()
            End If

        End If
    End Sub
    Private Sub btnYesDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnYesDelete.Click
        pnlVerifyDelete.Visible = False
        pnlSerialData.Visible = True
        deleteAllSerials()
    End Sub
    Private Sub btnNoDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNoDelete.Click
        pnlVerifyDelete.Visible = False
        pnlSerialData.Visible = True
    End Sub

    ' DROP DOWN LISTS
    Protected Sub ddlLoc_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlLoc.SelectedIndexChanged
        If ddlLoc.SelectedIndex <> 0 Then
            If btnResume.Visible = True Then
                ddlLoc.Enabled = False
                btnCreate.Visible = False
                btnResume.Enabled = False
                pnlPalletHeader.Enabled = True
                pnlPalletHeader.Visible = True
                lblWeek.Visible = True
                txtWeek.Visible = True
                cstValWeek.Enabled = True
                reqFieldValWeek.Enabled = True
                processEmptyFields(pnlPalletHeader)
            Else
                'configurePrinters()
                'ddlPrinters.Visible = True
                ddlMode.Visible = True
                ddlMode.Enabled = True
                ddlLoc.Enabled = False
            End If
        Else

        End If
    End Sub
    Protected Sub ddlMode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlMode.SelectedIndexChanged
        If ddlMode.SelectedIndex <> 0 Then
            ddlMode.Enabled = False
            pnlPalletHeader.Enabled = True
            pnlPalletHeader.Visible = True
            lblWeek.Visible = True
            txtWeek.Visible = True
            btnSubmit.Visible = True
            btnReset.Focus()
            processEmptyFields(pnlPalletHeader)
        Else
            resetPage()
        End If
    End Sub

    ' TEXTBOX
    Private Sub txtFilterLot_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterLot.TextChanged
        If checkFilterLot() Then
            txtFilterLot.Enabled = False
            txtWeek.Text = txtFilterLot.Text.Substring(5, 2)
            txtWeek.Enabled = False
            txtYear.Text = "20" & txtFilterLot.Text.Substring(3, 2)
            txtYear.Enabled = False
            processEmptyFields(pnlPalletHeader)
        Else
            processEmptyFields(pnlPalletHeader)
        End If
    End Sub
    Private Sub txtPallet_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPallet.TextChanged
        If txtPallet.Text <> "" Then
            If ddlMode.SelectedValue = 1 And btnCreate.Visible = True Then
                If txtPallet.Text.Length <= 3 Then
                    If IsNumeric(txtPallet.Text) Then
                        If checkPallet() Then
                            txtPallet.Enabled = False
                            Page.Validate()
                            displaySerialPanel()
                        Else
                            txtPallet.Text = ""
                            showError("ERROR CREATING PALLET!! TRY RESUME")
                            processEmptyFields(pnlPalletHeader)
                        End If
                    Else
                        txtPallet.Text = ""
                        showError("PALLET VALUE NEEDS TO BE NUMERIC ONLY!!")
                        processEmptyFields(pnlPalletHeader)
                    End If
                Else
                    txtPallet.Text = ""
                    showError("PALLET VALUE NEEDS TO BE NUMERIC ONLY!!")
                    processEmptyFields(pnlPalletHeader)
                End If
            Else
                txtPallet.Text = txtPallet.Text.ToUpper
                txtPallet.Enabled = False
                btnSubmit.Enabled = True
                btnSubmit.Visible = True
                Page.SetFocus("btnSubmit")
            End If
            
        End If
    End Sub
    Private Sub txtWeek_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeek.TextChanged
        txtWeek.Text = txtWeek.Text.PadLeft(2, "0")
        
        processEmptyFields(pnlPalletHeader)
    End Sub
    Private Sub txtYear_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtYear.TextChanged
        processEmptyFields(pnlPalletHeader)
    End Sub
    Protected Sub txtSerial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSerial.TextChanged
        ' Runs when you enter a serial number
        Dim errMsg As String = ""
        Dim barcode As String
        Dim barcodeArray() As String
        Dim filterLotValid As Boolean = True
        Dim barcodeStringCount As Integer = 0
        Dim validInput As Boolean = True
        Dim barcodeIndex As Integer = 0

        Try
            ' BYPASS FOR 2ND SCAN - DAVID VAUGHAN 8/26/11 VIA HELPDESK TICKET

            'Changes the serial characters to upper
            txtSerial.Text = txtSerial.Text.ToUpper()

            'sets lblSerialCheck.Text to the upper cased serial barcode
            lblSerialCheck.Text = txtSerial.Text


            grdDupBarcodes.Visible = False

            'If the mode is set to Assembly or In Process QA Run this code
            If ddlMode.SelectedValue = 1 Or ddlMode.SelectedValue = 5 Then

                ' This just validates that the serial is not blank
                If lblSerialCheck.Text = "" Then

                    lblSerialCheck.Text = txtSerial.Text
                    txtSerial.Text = ""
                    showError("ENTER SERIAL AGAIN PLEASE")
                    Page.SetFocus("txtSerial")
                    Exit Sub

                Else
                    'If the serial is not blank check that the second scan is a match? This makes no sense, as only one scan is necessary. This may be depricated code that was not removed
                    If txtSerial.Text <> lblSerialCheck.Text Then
                        showError("SECOND SCAN NOT A MATCH:" & txtSerial.Text & " / " & lblSerialCheck.Text)
                        txtSerial.Text = ""
                        lblSerialCheck.Text = ""
                        Page.SetFocus("txtSerial")
                        Exit Sub
                    End If

                End If
            End If

            'If this serial entered contains the letter S, then the barcode Array, needs to add however many serials delimited with an S 
            If txtSerial.Text.Contains("S") Then
                barcodeArray = txtSerial.Text.Split("S")
                'For each barcode entered into the barcode array, only give me the substring starting with position 0 and ending with position 10. This is the serial extracted from the barcode
                For x = 1 To barcodeArray.Length - 1
                    barcodeArray(x) = barcodeArray(x).Substring(0, 10)
                Next
            Else
                'If there is no S in the serials split the barcodes scanned using a * and add each to the barcode array
                barcodeArray = txtSerial.Text.Split("*")
            End If

            'CHECK DUPLICATE HERE

            'Checks the length of the selected barcode array to see if it is anything other than a spare. If not it can only be entered in pairs. If a spare you can add as many serials to the array as you like. 
            Select Case ddlMode.SelectedValue
                Case 1
                    If barcodeArray.Length > 2 Then
                        validInput = False
                        showError("ONLY FILTER PAIRS ALLOWED FOR ASSEMBLY")
                    End If
                Case 3
                    If barcodeArray.Length <> 2 And ddlLoc.SelectedValue = 2 Then
                        validInput = False
                        showError("ONLY FILTER PAIRS ALLOWED FOR QAR")
                    End If
                Case 4
                    If barcodeArray.Length > 2 Then
                        validInput = False
                        showError("ONLY SINGLE FILTERS ALLOWED FOR LAB")
                    End If
                Case 5
                    If barcodeArray.Length > 2 Then
                        validInput = False
                        showError("ONLY FILTER PAIRS ALLOWED FOR IN PROCESS QA")
                    End If
            End Select


            'If the length of the array is correct for the mode then...
            If validInput Then
                'Check to see if the filter lot number is correct because it was not entered manually
                If txtFilterLot.Text = "." Then
                    If barcodeArray(barcodeIndex).StartsWith("[") Then
                        barcodeIndex = 1
                    End If
                    If getFilterLot(barcodeArray(barcodeIndex)) Then
                        If txtFilterLot.Text.Substring(5, 2) <> txtWeek.Text Then
                            filterLotValid = False
                            showError("FILTER LOT / BARCODE WEEK DOESN'T MATCH PALLET WEEK -->" & txtFilterLot.Text)
                            txtFilterLot.Text = "."
                        End If
                    Else
                        showError("INVALID SERIALS OR FAILED TO GET FILTER LOT -->" & txtFilterLot.Text)
                        txtFilterLot.Text = "."
                        txtSerial.Text = ""
                        Exit Sub
                    End If
                End If

                If filterLotValid Then
                    If validateSerial(errMsg) Then
                        For Each barcode In barcodeArray
                            If barcode.StartsWith("[") Then
                                Continue For
                            End If
                            If addBarcodeToPallet(barcode) Then
                                btnDelete.Enabled = True
                                btnDeleteAll.Enabled = True
                                btnDeleteByScan.Enabled = True
                                barcodeStringCount += 1

                                If lstSerials.Items.Count = 18 And ddlMode.SelectedItem.Text = "Spares" And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                                    txtSerial.Enabled = False
                                    showError("SPARES PALLET COMPLETE")
                                End If
                            Else
                                showError("Serial could not be added to pallet")
                            End If
                        Next

                        If pnlSerialData.CssClass = "pnlSerialData" Then
                            pnlSerialData.CssClass = "pnlSerialDataChange"
                        Else
                            pnlSerialData.CssClass = "pnlSerialData"
                        End If

                        If barcodeStringCount <> 0 Then
                            If ddlMode.SelectedValue <> 4 And ddlLoc.SelectedValue <> 3 And ddlLoc.SelectedValue <> 1 Then
                                barcodeStringCount = barcodeStringCount / 2
                            End If
                        End If

                        lblCurrCount.Text = lblCurrCount.Text + barcodeStringCount

                        If ddlMode.SelectedValue = 1 Or ddlMode.SelectedValue = 5 Then
                            If ddlLoc.SelectedValue = 3 Then
                                lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 42) * 42)
                            Else
                                lblCurrRowCount.Text = lblCurrCount.Text - (Int(lblCurrCount.Text / 35) * 35)
                            End If
                        End If
                        addBarcodeStringToPallet()
                        lstSerials.Items.Add(txtSerial.Text)

                        If ddlMode.SelectedValue <> 2 Then
                            showError("BARCODE LAST SCANNED : " & txtSerial.Text)
                        End If

                        'REFRESH LIST / COUNT

                        refreshList()
                    Else
                        showError(errMsg)
                    End If
                End If
            End If



            txtSerial.Text = ""
            lblSerialCheck.Text = ""
            Page.SetFocus("txtSerial")
        Catch ex As Exception
            showError(ex.Message)
            'emailSubject = "APPLICATION ERROR: LABEL PRINTING : checkFilterLotStatus "
            'emailBody = ex.Message
            'emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub
    Private Sub txtSerialByScan_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSerialByScan.TextChanged
        Dim errMsg As String = ""
        Dim barcodeStringCount As Integer = 0

        Try
            txtSerialByScan.Text = txtSerialByScan.Text.ToUpper()
            If deleteSerialByScan(errMsg) Then
                txtSerialByScan.Text = ""
                Page.SetFocus("txtSerialByScan")
                'lblCurrCount.Text = lblCurrCount.Text + barcodeStringCount
                refreshList()
            Else
                showError(errMsg)
            End If
        Catch ex As Exception
            showError(ex.Message)
            emailSubject = "APPLICATION ERROR: LABEL PRINTING : txtSerialByScan "
            emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub

    'VALIDATORS
    Protected Sub cstValWeek_ServerValidate(ByVal source As Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles cstValWeek.ServerValidate
        Dim currWeek As Integer
        Dim currYear As Integer

        currWeek = DatePart(DateInterval.WeekOfYear, System.DateTime.Now)
        currYear = DatePart(DateInterval.Year, System.DateTime.Now)

        If System.Math.Abs(txtWeek.Text - currWeek) <= 52 And currYear = txtYear.Text Then
            args.IsValid = True
            If currWeek = txtWeek.Text And currYear = txtYear.Text Then

            Else
                Select Case currWeek
                    Case 1
                        If txtWeek.Text >= 49 And txtYear.Text = currYear - 1 Then

                        Else
                            If args.IsValid Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        End If
                    Case 2
                        If txtWeek.Text >= 50 And txtYear.Text = currYear - 1 Then

                        Else
                            If args.IsValid Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        End If
                    Case 3
                        If txtWeek.Text >= 51 And txtYear.Text = currYear - 1 Then

                        Else
                            If args.IsValid Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        End If
                    Case 4
                        If txtWeek.Text >= 52 And txtYear.Text = currYear - 1 Then

                        Else
                            If args.IsValid Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        End If
                    Case Else
                        If txtWeek.Text > currWeek And txtYear.Text = currYear Then
                            If txtWeek.Text = currWeek + 1 And txtYear.Text = currYear Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            Else
                                args.IsValid = False
                                txtWeek.Text = ""
                                processEmptyFields(pnlPalletHeader)
                            End If

                        Else
                            If txtWeek.Text >= currWeek - 5 And txtYear.Text = currYear Then

                            Else
                                If args.IsValid Then
                                    verifyNext = True
                                    pnlVerify.Visible = True
                                    pnlSerialData.Visible = False
                                End If
                            End If

                        End If

                End Select

            End If
        Else
            'NOT CURRENT YEAR
            If (txtWeek.Text = 31 Or txtWeek.Text = 51 Or txtWeek.Text >= 15 Or txtWeek.Text = 35) And txtYear.Text > currYear - 3 Then
                args.IsValid = True
                verifyNext = True
                pnlVerify.Visible = True
                pnlSerialData.Visible = False
            Else
                If currWeek <= 2 Then
                    ' TEST FOR PRIOR YEAR ROLL OVER
                    If txtWeek.Text >= 50 And txtYear.Text = currYear - 1 Then
                        args.IsValid = True
                        If currWeek = txtWeek.Text And currYear = txtYear.Text Then

                        Else
                            If args.IsValid Then
                                verifyNext = True
                                pnlVerify.Visible = True
                                pnlSerialData.Visible = False
                            End If
                        End If
                    Else
                        args.IsValid = False
                        txtWeek.Text = ""
                    End If
                Else
                    args.IsValid = False
                    txtWeek.Text = ""
                End If
            End If

        End If

    End Sub
    Protected Sub cstValCountVerify_ServerValidate(ByVal source As Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles cstValCountVerify.ServerValidate
        If IsNumeric(txtCountVerify.Text) Then
            If txtCountVerify.Text = lblCurrCount.Text Then
                args.IsValid = True
                countValid = True
            Else
                showError("COUNT DOES NOT MATCH")
                args.IsValid = False
            End If
        Else
            showError("NOT VALID NUMBER")
            args.IsValid = False
        End If
    End Sub

    ' PRINTING FUNCTIONS
    Private Sub printDocument()
        'Dim labelType As String = ""
        'Dim copies As Integer = 1

        Try
            'Dim printDoc As New System.Drawing.Printing.PrintDocument()
            'printDoc.PrintController = New System.Drawing.Printing.StandardPrintController

            'printDoc.PrinterSettings.PrinterName = ddlPrinters.Text
            'printDoc.PrinterSettings.PrinterName = "APSZebra13"
            'printDoc.PrinterSettings.PrinterName = "APS-PR-048"

            'Select Case ddlMode.SelectedValue
            '    Case 1, 2
            '        printDoc.PrinterSettings.Copies = 1
            '    Case Else
            '        printDoc.PrinterSettings.Copies = 1
            'End Select

            'AddHandler printDoc.PrintPage, AddressOf Me.printDoc_PrintPage
            printLabel()
            'printDoc.Print()

            'RemoveHandler printDoc.PrintPage, AddressOf Me.printDoc_PrintPage
            'Catch ex As System.AccessViolationException
            '    emailSubject = "APPLICATION ERROR SYSTEM ACCESS VIOLATION: FILTER PALLET ASSY: PrintDocument"
            '    emailBody = ex.Message

            '    Dim stkTrace As String

            '    Try
            '        stkTrace = ex.StackTrace
            '        emailBody = emailBody & stkTrace
            '    Catch stkEx As Exception
            '        emailBody = emailBody & " unable to acquire stack trace data"
            '    End Try

            '    emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)

            '    repairAppPool()

        Catch ex As Exception
            emailSubject = "APPLICATION ERROR: FILTER PALLET ASSY: PrintDocument"
            emailBody = ex.Message

            'Dim stkTrace As String

            'Try
            '    stkTrace = ex.StackTrace
            '    emailBody = emailBody & stkTrace
            'Catch stkEx As Exception
            '    emailBody = emailBody & " unable to acquire stack trace data"
            'End Try

            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        End Try
    End Sub
	'Private Sub printDoc_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
	'    Dim ds As New DataSet
	'    Dim dr As DataRow
	'    Dim prntData As New commonPrinting.printingClass
	'    Dim aList As New ArrayList
	'    Dim partTypeCode As String

	'    partTypeCode = "FILTER_PLT"

	'    Try
	'        ds = getLabelTemplate(partTypeCode)

	'        If ds Is Nothing Then
	'            showError("NO LABEL TEMPLATE FOR:" & partTypeCode)
	'        Else
	'            If ds.Tables(0).Rows.Count > 0 Then
	'                For Each dr In ds.Tables("LABEL_TEMPLATE").Rows
	'                    Select Case dr.Item("fieldType").ToString
	'                        Case "TEXT"
	'                            If dr.Item("fieldName") Is DBNull.Value Then
	'                                ' FIELD TITLE PRINT
	'                                prntData.renderTextData(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFont").ToString, dr.Item("fieldFontSize").ToString, dr.Item("fieldText").ToString, e)
	'                            Else
	'                                ' FIELD TEXT PRINT
	'                                prntData.renderTextData(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFont").ToString, dr.Item("fieldFontSize").ToString, getFieldText(dr.Item("fieldName")).ToUpper, e)
	'                            End If
	'                        Case "BARCODE"
	'                            Select Case dr.Item("fieldName").ToString
	'                                Case "LOT"
	'                                    prntData.renderBarcodeDataLabel(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFontSize").ToString, txtFilterLot.Text.ToUpper, CType(dr.Item("fieldWidth"), Single), CType(dr.Item("fieldRatio"), Single), CType(dr.Item("fieldHRFontSize"), Integer), e)
	'                                Case "WEEK"
	'                                    prntData.renderBarcodeDataLabel(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFontSize").ToString, txtWeek.Text.ToUpper, CType(dr.Item("fieldWidth"), Single), CType(dr.Item("fieldRatio"), Single), CType(dr.Item("fieldHRFontSize"), Integer), e)
	'                                Case "YEAR"
	'                                    prntData.renderBarcodeDataLabel(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFontSize").ToString, txtYear.Text.ToUpper, CType(dr.Item("fieldWidth"), Single), CType(dr.Item("fieldRatio"), Single), CType(dr.Item("fieldHRFontSize"), Integer), e)
	'                                Case "PALLET"
	'                                    prntData.renderBarcodeDataLabel(CType(dr.Item("fieldXCoord").ToString, Single), CType(dr.Item("fieldYCoord"), Single), dr.Item("fieldFontSize").ToString, txtPallet.Text.ToUpper, CType(dr.Item("fieldWidth"), Single), CType(dr.Item("fieldRatio"), Single), CType(dr.Item("fieldHRFontSize"), Integer), e)
	'                            End Select
	'                    End Select
	'                Next
	'            Else
	'                lblError.Text = "NO LABEL TEMPLATE FOR TYPE: " & partTypeCode
	'                lblError.Visible = True
	'            End If
	'        End If
	'    Catch ex As Exception
	'        emailSubject = "APPLICATION ERROR: FILTER PALLET ASSY: printDoc_PrintPage"
	'        emailBody = ex.Message
	'        emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
	'    End Try
	'End Sub
	Private Sub printLabel()
        Dim w As System.IO.StreamWriter
        Dim fileLoc As String = ""
        Dim barcodeString As String = ""
        Dim dateText As String
        Dim randomText As String = ""
        Dim num As Integer
		Dim dataString As String = ""
		Dim lineNum As Integer = 0

		Try
            While randomText = ""
                For i = 1 To 4  ' Random Text will have 8 characters
                    num = getRandomNumber()
                    randomText = randomText & CStr(num)
                Next
            End While



			dateText = CStr(System.DateTime.Today)
			dateText = dateText.Replace("/"c, "-"c)


			fileLoc = Server.MapPath("~\PrintDir\FILTER_PALLET_ASSY" & dateText & "_" & randomText & ".dd")

			dataString = "LOT,PLT,WEEK,YEAR,QTY,DATE" & vbCrLf
			dataString = dataString & txtFilterLot.Text.ToUpper & "," & txtPallet.Text.ToUpper & "," & txtWeek.Text & "," & txtYear.Text & "," & lblCurrCount.Text & "," & CStr(DateAdd(DateInterval.Hour, 1, System.DateTime.Now))

			w = New System.IO.StreamWriter(fileLoc)
            w.Write(dataString)
        Catch ex As Exception
			emailSubject = "APPLICATION ERROR: Filter Pallet Program: printLabel() "
			emailBody = ex.Message
            emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
        Finally
            w.Flush()
            w.Close()
        End Try
    End Sub

    ' REPAIR APPLICATION POOL
    'Private Sub repairAppPool()
    '    Dim appPoolID As String
    '    If checkAppPoolRunning() Then
    '        appPoolID = getAppPoolID()
    '        If appPoolID <> "" Then
    '            recycleAppPool(appPoolID)
    '        End If
    '    End If
    'End Sub
    'Private Sub recycleAppPool(ByVal appPoolID As String)
    '    Dim appPoolPath As String
    '    Dim appPoolEntry As DirectoryEntry

    '    appPoolPath = "IIS://localhost/W3SVC/AppPools/" + appPoolID
    '    appPoolEntry = New DirectoryEntry(appPoolPath)
    '    appPoolEntry.Invoke("Recycle")

    'End Sub
    'Private Function checkAppPoolRunning() As Boolean
    '    If Not (AppDomain.CurrentDomain.FriendlyName.StartsWith("/LM/")) Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function
    'Private Function getAppPoolID() As String
    '    Dim virDirPath As String
    '    Dim index As Integer
    '    Dim virDirEntry As DirectoryEntry

    '    virDirPath = AppDomain.CurrentDomain.FriendlyName
    '    virDirPath = virDirPath.Substring(4)
    '    index = virDirPath.Length + 1
    '    index = virDirPath.LastIndexOf("-", index - 1, index - 1)
    '    index = virDirPath.LastIndexOf("-", index - 1, index - 1)
    '    virDirPath = "IIS://localhost/" + virDirPath.Remove(index)

    '    virDirEntry = New DirectoryEntry(virDirPath)

    '    Return virDirEntry.Properties("AppPoolId").Value.ToString
    'End Function
    
    
   
End Class