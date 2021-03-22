Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Web.Http
Imports System.Xml
Imports Newtonsoft.Json
Imports System.Web

Namespace Controllers
    Public Class FBOrderCreationController
        Inherits ApiController
        Protected dbad As New OleDbDataAdapter
        Dim Connection As [String] = ConfigurationManager.AppSettings("dbconn.ConnectionString")
        Dim fbtestredirectingurl As String = System.Configuration.ConfigurationManager.AppSettings("fbtestredirectingurl")
        Protected Shared AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
        Public dbadc As New OleDbCommand
        Dim conn As New OleDbConnection(Connection)
        Dim cmd As New OleDbCommand()
        Public Function Post(request As HttpRequestMessage) As HttpResponseMessage
            Dim result As String = String.Empty
            Dim resp = New HttpResponseMessage(HttpStatusCode.OK)
            dbad.SelectCommand = New OleDbCommand
            dbad.UpdateCommand = New OleDbCommand
            Dim res As Int32 = 0
            Dim content = request.Content
            Dim jsonContent As String = content.ReadAsStringAsync().Result

            Try
                'Saving XML
                Dim path As String = "C:\Omegacube_ERP\PORTALS\POSTED_XML\"
                If Not Directory.Exists(path) Then
                    Directory.CreateDirectory(path)
                End If
                path = path & DateTime.Now.ToString("yyyyMMddHHmmss") & ".XML"
               
                Using writer As New StreamWriter(path, True)
                    writer.WriteLine(Replace(HttpContext.Current.Server.UrlDecode(jsonContent), "cxml-urlencoded=", ""))

                End Using

                Dim Header_Id, From_Id, TOTAL, FROM_IDENTITY, FROM_SHARED_SECRET, To_Id, TO_IDENTITY, TO_SHARED_SECRET, Sender_Id, SENDER_IDENTITY, SENDER_SHARED_SECRET, USER_AGENT, Message_Id, PunchOutOrderMessage_Id, BUYER_COOKIE, PunchOutOrderMessageHeader_Id, Total_Id, TOTAL_CURRENCY, TOTAL_MONEY, SHIPPING_CURRENCY, SHIPPING_MONEY, SHIPPING_DESCP, TAX_CURRENCY, TAX_MONEY, TAX_DESCP, Qty_detail, LINE_NUMBER, SUPPLIERPARTID, SUPPLIER_AUX_ID, ITEM_QTY, ITEM_UOM, UNIT_PRICE_MONEY, Item_detail_Id, ITEM_DECSC, MANUFACTURERPART_ID, ITEM_CLASSIFICATION, SHIPPING_ITEM_CURRENCY, SHIPPING_ITEM_MONEY, SHIPPING_ITEM_DESCP, TAX_ITEM_CURRENCY, TAX_ITEM_MONEY, TAX_ITEM_DESCP, Postal_ID As String
                '   Call Delete_query("DELETE FROM TEMP_MRO_REQUISITION_ORDER")
                ' ---
                Dim MRO_REQ_ORDER_NO_ds As New DataSet
                MRO_REQ_ORDER_NO_ds = Return_record_set("SELECT SEQ_MRO_REQ_ORDER.NEXTVAL FROM DUAL")
                Dim MRO_REQ_ORDER_NO As String
                MRO_REQ_ORDER_NO = MRO_REQ_ORDER_NO_ds.Tables(0).Rows(0)(0)
                Dim ds As New Data.DataSet
                ds.Clear()
                ds.ReadXml(path)
                For Each dr In ds.Tables("Message").Rows
                    Message_Id = dr("Message_Id").ToString.Trim
                    Dim dr_Message() As DataRow = ds.Tables("Message").Select("Message_Id Is Not null")
                    For Each dr_Message_ID As DataRow In dr_Message
                        Message_Id = dr("Message_Id").ToString.Trim
                        Dim dr_PunnchOutOrdertable() As DataRow = ds.Tables("PunchOutOrderMessage").Select("Message_Id='" & Message_Id & "'")
                        For Each dr_Punchout_Post As DataRow In dr_PunnchOutOrdertable
                            PunchOutOrderMessage_Id = dr_Punchout_Post("PunchOutOrderMessage_Id").ToString.Trim
                            Try
                                If Not IsDBNull(dr_Punchout_Post("BuyerCookie").ToString.Trim) Then
                                    BUYER_COOKIE = dr_Punchout_Post("BuyerCookie").ToString.Trim
                                Else
                                    BUYER_COOKIE = ""
                                End If
                            Catch ex As Exception
                                BUYER_COOKIE = ""
                            End Try
                            'PunchoutMessageHeader
                            Dim dr_PunchOutOrderMessageHeader() As DataRow = ds.Tables("PunchOutOrderMessageHeader").Select("PunchOutOrderMessage_Id='" & PunchOutOrderMessage_Id & "'")
                            'ItemIn
                            Dim dr_Item_In() As DataRow = ds.Tables("ItemIn").Select("PunchOutOrderMessage_Id='" & PunchOutOrderMessage_Id & "'")
                            'Dim Qty_detail, LINE_NUMBER, SUPPLIERPARTID, SUPPLIER_AUX_ID, QTY, UOM, UNIT_PRICE_MONEY, Item_detail_Id, ITEM_DECSC, MANUFACTURERPART_ID, ITEM_CLASSIFICATION, SHIPPING_ITEM_CURRENCY, SHIPPING_ITEM_MONEY, SHIPPING_ITEM_DESCP, TAX_ITEM_CURRENCY, TAX_ITEM_MONEY, TAX_ITEM_DESCP, Postal_ID As String
                            For Each dr_I_O As DataRow In dr_Item_In
                                If Not IsDBNull(dr_I_O("ItemIn_Id").ToString) Then
                                    Dim ItemIn_Id As String = dr_I_O("ItemIn_Id").ToString.Trim
                                    Qty_detail = dr_I_O("quantity").ToString.Trim
                                    LINE_NUMBER = dr_I_O("LineNumber").ToString.Trim

                                    Dim dr_ItemId() As DataRow = ds.Tables("ItemID").Select("ItemIn_Id='" & ItemIn_Id & "'")
                                    For Each dr_it As DataRow In dr_ItemId
                                        If Not IsDBNull(dr_it("SupplierPartID").ToString) Then

                                            SUPPLIERPARTID = UCase(dr_it("SupplierPartID").ToString.Trim)
                                            SUPPLIERPARTID = SUPPLIERPARTID.Replace("'", "''")

                                        End If
                                        If Not IsDBNull(dr_it("SupplierPartAuxiliaryID").ToString) Then

                                            SUPPLIER_AUX_ID = UCase(dr_it("SupplierPartAuxiliaryID").ToString.Trim)
                                            SUPPLIER_AUX_ID = SUPPLIER_AUX_ID.Replace("'", "''")

                                        End If

                                        Dim dr_ItemDetail() As DataRow = ds.Tables("ItemDetail").Select("ItemIn_Id='" & ItemIn_Id & "'")
                                        For Each dr_UOM As DataRow In dr_ItemDetail
                                            If Not IsDBNull(dr_UOM("UnitOfMeasure").ToString) Then
                                                ' QTY = Convert.ToInt32(Qty_detail) * Convert.ToInt32(dr_UOM("UnitOfMeasure"))
                                                ITEM_QTY = Convert.ToInt32(Qty_detail)
                                                'UOM = dr_UOM("UnitOfMeasure").ToString.Trim
                                                ITEM_UOM = UCase(dr_UOM("UnitOfMeasure").ToString.Trim)

                                                MANUFACTURERPART_ID = UCase(dr_UOM("ManufacturerPartID").ToString.Trim)
                                            End If
                                            Item_detail_Id = dr_UOM("ItemDetail_Id").ToString
                                        Next
                                        'Unit Price
                                        Dim dr_UP() As DataRow = ds.Tables("UnitPrice").Select("ItemDetail_Id='" & ItemIn_Id & "'")
                                        For Each dr_UP1 As DataRow In dr_UP
                                            Dim UnitP As String = dr_UP1("UnitPrice_Id").ToString
                                            Dim dr_MNY() As DataRow = ds.Tables("Money").Select("UnitPrice_Id='" & UnitP & "'")
                                            For Each dr_MN As DataRow In dr_MNY
                                                If Not IsDBNull(dr_MN("Money_Text").ToString) Then
                                                    UNIT_PRICE_MONEY = dr_MN("Money_Text").ToString.Trim
                                                End If
                                            Next

                                        Next



                                        'Description
                                        Dim dr_ItemDesc() As DataRow = ds.Tables("Description").Select("ItemDetail_Id='" & ItemIn_Id & "'")
                                        For Each dr_Desc As DataRow In dr_ItemDesc
                                            If Not IsDBNull(dr_Desc("Description_Text").ToString) Then
                                                'ITEM_DECSC = Replace(dr_Desc("Description_Text").ToString, ",", "-")
                                                ITEM_DECSC = Replace(dr_Desc("Description_Text").ToString, "'", "''")
                                            Else
                                                ITEM_DECSC = ""
                                            End If
                                        Next

                                        'Classification
                                        Dim dr_Classification() As DataRow = ds.Tables("Classification").Select("ItemDetail_Id='" & ItemIn_Id & "'")
                                        For Each dr_Clasif As DataRow In dr_Classification
                                            If Not IsDBNull(dr_Clasif("Classification_Text").ToString) Then
                                                ITEM_CLASSIFICATION = Replace(dr_Clasif("Classification_Text").ToString, "'", "''")
                                                'ITEM_DECSC = ITEM_DECSC.Substring(0, 100)
                                            Else
                                                ITEM_CLASSIFICATION = ""
                                            End If
                                        Next

                                        'Shipping Item
                                        Dim dr_ShippingItem() As DataRow = ds.Tables("Shipping").Select("ItemIn_Id='" & ItemIn_Id & "'")
                                        For Each dr_UP1 As DataRow In dr_ShippingItem
                                            Dim Shipping_Item_Id As String = dr_UP1("Shipping_Id").ToString
                                            Dim dr_MNY() As DataRow = ds.Tables("Money").Select("Shipping_Id='" & Shipping_Item_Id & "'")
                                            For Each dr_MN As DataRow In dr_MNY
                                                Try
                                                    If Not IsDBNull(dr_MN("currency").ToString.Trim) Then
                                                        SHIPPING_ITEM_CURRENCY = dr_MN("currency").ToString.Trim
                                                    Else
                                                        SHIPPING_ITEM_CURRENCY = ""
                                                    End If
                                                Catch ex As Exception
                                                    SHIPPING_ITEM_CURRENCY = ""
                                                End Try
                                                Try
                                                    If Not IsDBNull(dr_MN("Money_Text").ToString.Trim) Then
                                                        SHIPPING_ITEM_MONEY = dr_MN("Money_Text").ToString.Trim
                                                    Else
                                                        SHIPPING_ITEM_MONEY = ""
                                                    End If
                                                Catch ex As Exception
                                                    SHIPPING_MONEY = ""
                                                End Try
                                            Next
                                            Dim dr_ShipItemDesc() As DataRow = ds.Tables("Description").Select("Shipping_Id='" & Shipping_Item_Id & "'")
                                            For Each dr_Desc As DataRow In dr_ShipItemDesc
                                                If Not IsDBNull(dr_Desc("Description_Text").ToString) Then
                                                    'SHIPPING_ITEM_DESCP = Replace(dr_Desc("Description_Text").ToString, ",", "-")
                                                    SHIPPING_ITEM_DESCP = Replace(dr_Desc("Description_Text").ToString, "'", "''")
                                                    'Disc_Text = TAG_LINE.Substring(0, 100)
                                                Else
                                                    SHIPPING_ITEM_DESCP = ""
                                                End If
                                            Next

                                        Next
                                        'end Shipping Item
                                        'Shipping Item
                                        Dim dr_TaxItem() As DataRow = ds.Tables("Tax").Select("ItemIn_Id='" & ItemIn_Id & "'")
                                        For Each dr_UP1 As DataRow In dr_TaxItem
                                            Dim Tax_Item_Id As String = dr_UP1("Tax_Id").ToString
                                            Dim dr_MNY() As DataRow = ds.Tables("Money").Select("Tax_Id='" & Tax_Item_Id & "'")
                                            For Each dr_MN As DataRow In dr_MNY
                                                Try
                                                    If Not IsDBNull(dr_MN("currency").ToString.Trim) Then
                                                        TAX_ITEM_CURRENCY = dr_MN("currency").ToString.Trim
                                                    Else
                                                        TAX_ITEM_CURRENCY = ""
                                                    End If
                                                Catch ex As Exception
                                                    TAX_ITEM_CURRENCY = ""
                                                End Try
                                                Try
                                                    If Not IsDBNull(dr_MN("Money_Text").ToString.Trim) Then
                                                        TAX_ITEM_MONEY = dr_MN("Money_Text").ToString.Trim
                                                    Else
                                                        TAX_ITEM_MONEY = ""
                                                    End If
                                                Catch ex As Exception
                                                    TAX_ITEM_MONEY = ""
                                                End Try
                                            Next
                                            Dim dr_ShipItemDesc() As DataRow = ds.Tables("Description").Select("Tax_Id='" & Tax_Item_Id & "'")
                                            For Each dr_Desc As DataRow In dr_ShipItemDesc
                                                If Not IsDBNull(dr_Desc("Description_Text").ToString) Then
                                                    'TAX_ITEM_DESCP = Replace(dr_Desc("Description_Text").ToString, ",", "-")
                                                    TAX_ITEM_DESCP = Replace(dr_Desc("Description_Text").ToString, "'", "''")
                                                    'Disc_Text = TAG_LINE.Substring(0, 100)
                                                Else
                                                    TAX_ITEM_DESCP = ""
                                                End If
                                            Next

                                        Next
                                        'end Shipping Item

                                    Next
                                End If

                                'End ItemIn

                                For Each dr_PunchOutOrderMessageHeader_Post As DataRow In dr_PunchOutOrderMessageHeader
                                    PunchOutOrderMessageHeader_Id = dr_PunchOutOrderMessageHeader_Post("PunchOutOrderMessageHeader_Id").ToString.Trim
                                    'total
                                    'PunchOutOrderMessageHeader_Id,Total_Id
                                    Dim dr_Total() As DataRow = ds.Tables("Total").Select("PunchOutOrderMessageHeader_Id='" & PunchOutOrderMessageHeader_Id & "'")
                                    For Each dr_Total_Post As DataRow In dr_Total
                                        Total_Id = dr_Total_Post("Total_Id").ToString.Trim
                                        'money
                                        Dim dr_Money() As DataRow = ds.Tables("Money").Select("Total_Id='" & Total_Id & "'")
                                        For Each dr_Money_Post As DataRow In dr_Money
                                            'CURRENCY,MONEY_TEXT,SHIPPING_ID,TAX_ID,UNIT_PRICE_ID
                                            Try
                                                If Not IsDBNull(dr_Money_Post("currency").ToString.Trim) Then
                                                    TOTAL_CURRENCY = dr_Money_Post("currency").ToString.Trim
                                                Else
                                                    TOTAL_CURRENCY = ""
                                                End If
                                            Catch ex As Exception
                                                TOTAL_CURRENCY = ""
                                            End Try

                                            Try
                                                If Not IsDBNull(dr_Money_Post("Money_Text").ToString.Trim) Then
                                                    TOTAL_MONEY = dr_Money_Post("Money_Text").ToString.Trim
                                                Else
                                                    TOTAL_MONEY = ""
                                                End If
                                            Catch ex As Exception
                                                TOTAL_MONEY = ""
                                            End Try
                                        Next
                                        'end money
                                    Next
                                    'end total
                                    'shipping
                                    Dim dr_Shipping() As DataRow = ds.Tables("Shipping").Select("PunchOutOrderMessageHeader_Id='" & PunchOutOrderMessageHeader_Id & "'")
                                    For Each dr_UP1 As DataRow In dr_Shipping
                                        Dim Shipping_Id As String = dr_UP1("Shipping_Id").ToString
                                        Dim dr_MNY() As DataRow = ds.Tables("Money").Select("Shipping_Id='" & Shipping_Id & "'")
                                        For Each dr_MN As DataRow In dr_MNY
                                            Try
                                                If Not IsDBNull(dr_MN("currency").ToString.Trim) Then
                                                    SHIPPING_CURRENCY = dr_MN("currency").ToString.Trim
                                                Else
                                                    SHIPPING_CURRENCY = ""
                                                End If
                                            Catch ex As Exception
                                                SHIPPING_CURRENCY = ""
                                            End Try
                                            Try
                                                If Not IsDBNull(dr_MN("Money_Text").ToString.Trim) Then
                                                    SHIPPING_MONEY = dr_MN("Money_Text").ToString.Trim
                                                Else
                                                    SHIPPING_MONEY = ""
                                                End If
                                            Catch ex As Exception
                                                SHIPPING_MONEY = ""
                                            End Try
                                        Next
                                        Dim dr_ItemDesc() As DataRow = ds.Tables("Description").Select("Shipping_Id='" & Shipping_Id & "'")
                                        For Each dr_Desc As DataRow In dr_ItemDesc
                                            If Not IsDBNull(dr_Desc("Description_Text").ToString) Then
                                                'SHIPPING_DESCP = Replace(dr_Desc("Description_Text").ToString, ",", "-")
                                                SHIPPING_DESCP = Replace(dr_Desc("Description_Text").ToString, "'", "''")
                                                'Disc_Text = TAG_LINE.Substring(0, 100)
                                            Else
                                                SHIPPING_DESCP = ""
                                            End If
                                        Next

                                    Next
                                    'end Shipping
                                    'tax
                                    Dim dr_Tax() As DataRow = ds.Tables("Tax").Select("PunchOutOrderMessageHeader_Id='" & PunchOutOrderMessageHeader_Id & "'")
                                    For Each dr_tax11 As DataRow In dr_Tax
                                        Dim Tax_Id As String = dr_tax11("Tax_Id").ToString
                                        Dim dr_MNY() As DataRow = ds.Tables("Money").Select("Tax_Id='" & Tax_Id & "'")
                                        For Each dr_MN As DataRow In dr_MNY
                                            Try
                                                If Not IsDBNull(dr_MN("currency").ToString.Trim) Then
                                                    TAX_CURRENCY = dr_MN("currency").ToString.Trim
                                                Else
                                                    TAX_CURRENCY = ""
                                                End If
                                            Catch ex As Exception
                                                TAX_CURRENCY = ""
                                            End Try
                                            Try
                                                If Not IsDBNull(dr_MN("Money_Text").ToString.Trim) Then
                                                    TAX_MONEY = dr_MN("Money_Text").ToString.Trim
                                                Else
                                                    TAX_MONEY = ""
                                                End If
                                            Catch ex As Exception
                                                TAX_MONEY = ""
                                            End Try
                                        Next
                                        Dim dr_ItemDesc() As DataRow = ds.Tables("Description").Select("Tax_Id='" & Tax_Id & "'")
                                        For Each dr_Desc As DataRow In dr_ItemDesc
                                            If Not IsDBNull(dr_Desc("Description_Text").ToString) Then
                                                'TAX_DESCP = Replace(dr_Desc("Description_Text").ToString, ",", "-")
                                                TAX_DESCP = Replace(dr_Desc("Description_Text").ToString, "'", "''")
                                                'Disc_Text = TAG_LINE.Substring(0, 100)
                                            Else
                                                TAX_DESCP = ""
                                            End If
                                        Next

                                    Next
                                    'end tax
                                    Call Insert_query("Insert Into TEMP_MRO_REQUISITION_ORDER(MRO_REQ_ORDER_NO,BUYER_COOKIE,PUNCH_ORDER_MSG_HEADER_ID,TOTAL_ID,TOTAL_CURRENCY,TOTAL_MONEY,SHIPPING_CURRENCY,SHIPPING_MONEY,SHIPPING_DESCP,TAX_CURRENCY,TAX_MONEY,TAX_DESCP,QTY_DETAIL,LINE_NUMBER,SUPPLIER_PART_ID,SUPPLIER_AUX_ID,ITEM_QTY,ITEM_UOM,UNIT_PRICE_MONEY,ITEM_DETAIL_ID,ITEM_DECSC,MANUFACTURER_PART_ID,ITEM_CLASSIFICATION,SHIP_ITEM_CURRENCY,SHIP_ITEM_MONEY,SHIP_ITEM_DESCP ,TAX_ITEM_CURRENCY,TAX_ITEM_MONEY,TAX_ITEM_DESCP,CREATED_DATE,CREATED_BY) Values('" & MRO_REQ_ORDER_NO & "','" & BUYER_COOKIE & "','" & PunchOutOrderMessageHeader_Id & "','" & Total_Id & "','" & TOTAL_CURRENCY & "','" & TOTAL_MONEY & "','" & SHIPPING_CURRENCY & "','" & SHIPPING_MONEY & "','" & SHIPPING_DESCP & "','" & TAX_CURRENCY & "','" & TAX_MONEY & "','" & TAX_DESCP & "','" & Qty_detail & "','" & LINE_NUMBER & "','" & SUPPLIERPARTID & "','" & SUPPLIER_AUX_ID & "','" & ITEM_QTY & "','" & ITEM_UOM & "','" & UNIT_PRICE_MONEY & "','" & Item_detail_Id & "','" & ITEM_DECSC & "','" & MANUFACTURERPART_ID & "','" & ITEM_CLASSIFICATION & "','" & SHIPPING_ITEM_CURRENCY & "','" & SHIPPING_ITEM_MONEY & "','" & SHIPPING_ITEM_DESCP & "','" & TAX_ITEM_CURRENCY & "','" & TAX_ITEM_MONEY & "','" & TAX_ITEM_DESCP & "',sysdate,'OCT')")
                                Next
                            Next

                        Next
                        'End PunchoutMessage Header
                    Next

                    'Call Insert_query("Insert Into TEMP_MRO_REQUISITION_ORDER(HEADER_ID,FROM_ID,TOTAL,FROM_IDENTITY,FROM_SHARED_SECRET,TO_ID,TO_IDENTITY,TO_SHARED_SECRET,SENDER_ID,SENDER_IDENTITY,SENDER_SHARED_SECRET,USER_AGENT,MESSAGE_ID,PUNCH_ORDER_MSG_ID,PUNCH_ORDER_MSG_HEADER_ID,TOTAL_ID,TOTAL_CURRENCY,TOTAL_MONEY,SHIPPING_CURRENCY,SHIPPING_MONEY,SHIPPING_DESCP,TAX_CURRENCY,TAX_MONEY,TAX_DESCP,QTY_DETAIL,LINE_NUMBER,SUPPLIER_PART_ID,SUPPLIER_AUX_ID,ITEM_QTY,ITEM_UOM,UNIT_PRICE_MONEY,ITEM_DETAIL_ID,ITEM_DECSC,MANUFACTURER_PART_ID,ITEM_CLASSIFICATION,SHIP_ITEM_CURRENCY,SHIP_ITEM_MONEY,HIP_ITEM_DESCP ,TAX_ITEM_CURRENCY,AX_ITEM_MONEY,TAX_ITEM_DESCP,CREATED_DATE,CREATED_BY) Values('" & HEADER_ID & "','" & FROM_ID & "','" & TOTAL & "','" & FROM_IDENTITY & "','" & FROM_SHARED_SECRET & "','" & TO_ID & "','" & TO_IDENTITY & "','" & TO_SHARED_SECRET & "','" & SENDER_ID & "','" & SENDER_IDENTITY & "','" & SENDER_SHARED_SECRET & "','" & USER_AGENT & "','" & MESSAGE_ID & "','" & PUNCHOUTORDERMESSAGE_ID & "','" & BUYER_COOKIE & "','" & PUNCHOUTORDERMESSAGEHEADER_ID & "','" & TOTAL_ID & "','" & TOTAL_CURRENCY & "','" & TOTAL_MONEY & "','" & SHIPPING_CURRENCY & "','" & SHIPPING_MONEY & "','" & SHIPPING_DESCP & "','" & TAX_CURRENCY & "','" & TAX_MONEY & "','" & TAX_DESCP & "','" & QTY_DETAIL & "','" & LINE_NUMBER & "','" & SUPPLIERPARTID & "','" & SUPPLIER_AUX_ID & "','" & ITEM_QTY & "','" & ITEM_UOM & "','" & UNIT_PRICE_MONEY & "','" & ITEM_DETAIL_ID & "','" & ITEM_DECSC & "','" & MANUFACTURERPART_ID & "','" & ITEM_CLASSIFICATION & "','" & SHIPPING_ITEM_CURRENCY & "','" & SHIPPING_ITEM_MONEY & "','" & SHIPPING_ITEM_DESCP & "','" & TAX_ITEM_CURRENCY & "','" & TAX_ITEM_MONEY & "','" & TAX_ITEM_DESCP & "',sysdate,'" & Session("USER_ID") & "')")









                    'err_msg.Text = "<strong><font color='#42C120' size='3' face='Arial'>Loaded Successfully</font></strong>"
                Next
                Call execute_storeProcedure("SP_IMPORT_MRO_REQUISITION_XML", "OCT#" & MRO_REQ_ORDER_NO, "P_USER_ID#P_MRO_REQ_ORDER_NO", "C#C")
                Dim DSEDGE As Data.DataSet
                Dim EdgePRNo As String
                DSEDGE = Return_record_set("select distinct REQUISITION_NO AS PR_NO from TEMP_MRO_REQUISITION_ORDER WHERE MRO_REQ_ORDER_NO='" & MRO_REQ_ORDER_NO & "'")
                DSEDGE.Clear()
                dbad.Fill(DSEDGE)

                If (DSEDGE.Tables(0).Rows.Count > 0) Then

                    If Not (Equals(DSEDGE.Tables(0).Rows(0)("PR_NO"), System.DBNull.Value)) Then
                        EdgePRNo = DSEDGE.Tables(0).Rows(0)("PR_NO")

                        'result = "<strong><font color='#229954' size='3' face='Arial'>Purchase requisition has been Created. PR No #:</font></strong> <a href=""PO_REQUISITION.aspx?REQUISITION_NO=" & EdgePRNo & """ target='_blank'>" & EdgePRNo & " </a> <strong><font color='#229954' size='3' face='Arial'></font></strong>"
                        ' Dim response = request.CreateResponse(HttpStatusCode.Moved)
                        'Dim response1 = response1.redirect(Url, status)
                        'response.Headers.Location = New Uri("http://192.168.8.44:82/PO_REQUISITION.aspx?REQUISITION_NO=" & EdgePRNo & "")
                        'wanted code
                        'fbtestredirectingurl
                        Dim response = request.CreateResponse(HttpStatusCode.Moved)
                        ' response.Headers.Location = New Uri("http://192.168.8.44:82/PO_REQUISITION.aspx?REQUISITION_NO=" & EdgePRNo & ""
                        response.Headers.Location = New Uri("" & fbtestredirectingurl & "PO_REQUISITION.aspx?REQUISITION_NO=" & EdgePRNo & "")
                        Return response

                    Else
                        result = "<strong><font color='#FF0000' size='3' face='Arial'>PR creation failed, reload the correct file again </font></strong>"
                    End If
                End If
                ' ----


                Dim line As String
                Dim encoding = ASCIIEncoding.ASCII
                Using reader As New StreamReader(path, System.Text.Encoding.UTF8)
                    line = reader.ReadLine()
                End Using
                Console.WriteLine(line)
                result = result & "successfully inserted"
                resp.Content = New StringContent(result, System.Text.Encoding.UTF8, "text/plain")
                Return resp
                ' Return resp
            Catch ex As Exception
                Call insert_sys_log("FB", ex.Message)
            End Try


            'path = path & DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
            'If File.Exists(path) Then
            '    File.Delete(path)
            'End If
            'Dim sw As StreamWriter = New StreamWriter(path)
            'sw.WriteLine(requestHeaders)
            'sw.Close()
            'Dim sr1 As StreamReader = New StreamReader(path)
            'Console.WriteLine(sr1.ReadToEnd())
            'sr1.Close()



        End Function
        Public Sub Update_query(ByVal str_update As String)
            Try
                dbad.UpdateCommand = New OleDbCommand
                dbad.UpdateCommand.Connection = conn
                dbad.UpdateCommand.CommandText = str_update
                dbad.UpdateCommand.ExecuteNonQuery()
                dbad.UpdateCommand.Connection.Close()
            Catch ex As Exception
                Call insert_sys_log("Update_query", ex.Message)
            End Try
        End Sub

        Public Sub Insert_query(ByVal str_insert As String)
            Try
                dbad.InsertCommand = New OleDbCommand
                dbad.InsertCommand.Connection = conn
                dbad.InsertCommand.Connection.Open()
                dbad.InsertCommand.CommandText = str_insert
                dbad.InsertCommand.ExecuteNonQuery()
                dbad.InsertCommand.Connection.Close()
            Catch ex As Exception
                Call insert_sys_log("Insert_query", str_insert)
            End Try
        End Sub
        Public Sub Delete_query(ByVal str_delete As String)
            Try
                dbad.DeleteCommand = New OleDbCommand
                dbad.DeleteCommand.Connection = conn
                dbad.DeleteCommand.Connection.Open()
                dbad.DeleteCommand.CommandText = str_delete
                dbad.DeleteCommand.ExecuteNonQuery()
                dbad.DeleteCommand.Connection.Close()
            Catch ex As Exception
                Call insert_sys_log("Delete_query", ex.Message)
            End Try
        End Sub
        Public Function Return_record_count(ByVal str_select As String) As Integer
            Try
                Dim ds_new As New Data.DataSet
                dbad.SelectCommand = New OleDbCommand
                dbad.SelectCommand.Connection = conn
                dbad.SelectCommand.CommandText = str_select
                dbad.Fill(ds_new)
                dbad.SelectCommand.Connection.Close()
                If (ds_new.Tables(0).Rows.Count > 0) Then
                    Return 1
                Else
                    Return 0
                End If
                Return 0
            Catch ex As Exception
                Call insert_sys_log("Return_record_count", ex.Message)
            End Try
        End Function
        Public Function Return_record_set(ByVal str_select As String) As Data.DataSet
            Try
                Dim ds_new As New Data.DataSet
                dbad.SelectCommand = New OleDbCommand
                dbad.SelectCommand.Connection = conn
                dbad.SelectCommand.CommandText = str_select
                dbad.Fill(ds_new)
                dbad.SelectCommand.Connection.Close()
                Return ds_new
            Catch ex As Exception
                Call insert_sys_log("Return_record_set", ex.Message)
            End Try
        End Function


        Public Sub execute_storeProcedure(ByVal fname As String, ByVal plist As String, ByVal plist1 As String, ByVal plist2 As String)
            Dim s As String
            s = ""
            Dim pu1, pu2, pu3
            Dim rct, pp As Integer
            If (plist <> "") Then
                pu1 = Split(plist, "#")
                pu2 = Split(plist1, "#")
                pu3 = Split(plist2, "#")
                If (UBound(pu1) > 0) Then
                    rct = UBound(pu1)
                Else
                    rct = 0
                End If
            Else
                rct = -1
            End If
            Dim rvalue As String
            dbad.SelectCommand.Connection = conn
            dbad.SelectCommand.Parameters.Clear()
            dbad.SelectCommand.CommandType = Data.CommandType.Text
            If (rct = -1) Then
                dbad.SelectCommand.CommandText = ("{call " & fname & "()}")
            Else
                If (rct = 0) Then
                    dbad.SelectCommand.CommandText = ("{call " & fname & "(?)}")
                Else
                    Dim rrr As String
                    rrr = ""
                    For pp = 0 To rct
                        rrr = rrr & "?" & ","
                    Next
                    rrr = Mid(rrr, 1, Len(rrr) - 1)
                    dbad.SelectCommand.CommandText = ("{call " & fname & "(" & rrr & ")}")
                End If
            End If
            If (rct >= 0) Then
                If (rct = 0) Then
                    If (UCase(plist2) = "N") Then
                        dbad.SelectCommand.Parameters.Add(New Data.OleDb.OleDbParameter(plist1, System.Data.OleDb.OleDbType.Double, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", Data.DataRowVersion.Current, plist))
                    Else
                        dbad.SelectCommand.Parameters.Add(New Data.OleDb.OleDbParameter(plist1, System.Data.OleDb.OleDbType.VarChar, 2000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", Data.DataRowVersion.Current, plist))
                    End If
                Else
                    For pp = 0 To rct
                        If (UCase(pu3(pp)) = "N") Then
                            dbad.SelectCommand.Parameters.Add(New Data.OleDb.OleDbParameter(pu2(pp), System.Data.OleDb.OleDbType.Double, 20, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", Data.DataRowVersion.Current, pu1(pp)))
                        Else
                            dbad.SelectCommand.Parameters.Add(New Data.OleDb.OleDbParameter(pu2(pp), System.Data.OleDb.OleDbType.VarChar, 2000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", Data.DataRowVersion.Current, pu1(pp)))
                        End If
                    Next
                End If
            End If
            If (dbad.SelectCommand.Connection.State = Data.ConnectionState.Closed) Then
                dbad.SelectCommand.Connection.Open()
            End If
            Try
                dbad.SelectCommand.ExecuteNonQuery()
            Catch exException As Exception
                dbad.SelectCommand.Connection.Close()
            End Try
            dbad.SelectCommand.Connection.Close()
        End Sub
        Public Sub insert_sys_log(ByVal str1 As String, ByVal message As String)
            Dim Connection As [String] = ConfigurationManager.AppSettings("dbconn.ConnectionString")
            Dim con As New OleDbConnection(Connection)
            dbad.InsertCommand = New OleDbCommand
            Dim sterr1, sterr2, sterr3, sterr4, sterr As String
            sterr = Replace(message, "'", "''")
            If (Len(sterr) > 4000) Then
                sterr1 = Mid(sterr, 1, 4000)
                If (Len(sterr) > 8000) Then
                    sterr2 = Mid(sterr, 4000, 8000)
                    If (Len(sterr) > 12000) Then
                        sterr3 = Mid(sterr, 8000, 12000)
                        If (Len(sterr) > 16000) Then
                            sterr4 = Mid(sterr, 12000, 16000)
                        Else
                            sterr4 = Mid(sterr, 12000, Len(sterr))
                        End If
                    Else
                        sterr3 = Mid(sterr, 8000, Len(sterr))
                        sterr4 = ""
                    End If
                Else
                    sterr2 = Mid(sterr, 4000, Len(sterr))
                    sterr3 = ""
                    sterr3 = ""
                    sterr4 = ""
                End If
            Else
                sterr1 = sterr
                sterr2 = ""
                sterr3 = ""
                sterr4 = ""
            End If
            Try
                dbad.InsertCommand.Connection = con
                If (dbad.InsertCommand.Connection.State = Data.ConnectionState.Closed) Then
                    dbad.InsertCommand.Connection.Open()
                End If
                dbad.InsertCommand.CommandText = "Insert into SYS_ACTIVATE_STATUS_LOG (LINE_NO, CHANGE_REQUEST_NO,  OBJECT_TYPE, OBJECT_NAME, ERROR_TEXT, STATUS,LOG_DATE,ERROR_TEXT1, ERROR_TEXT2, ERROR_TEXT3) values ((select nvl(max(to_number(line_no)),0)+1 from SYS_ACTIVATE_STATUS_LOG),'','FB_GRAINGER_ORDER_CREATION','" & str1 & "','" & sterr1 & "','N',sysdate,'" & sterr2 & "','" & sterr3 & "','" & sterr4 & "')"
                dbad.InsertCommand.ExecuteNonQuery()
                dbad.InsertCommand.Connection.Close()
            Catch ex As Exception

            End Try
        End Sub
    End Class
End Namespace