Public Class ReFactura

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double
    Dim ContOBNK, AORIN As Integer


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    Public Function Valid(ByVal DocNum As String, ByVal CardCode As String, ByVal psDirectory As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim DocEntry As String
        Dim ContP, ContORIN As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""DocEntry"" from OINV where (""DocNum""=" & DocNum & ")"
            oRecSetH.DoQuery(stQueryH)

            DocEntry = oRecSetH.Fields.Item("DocEntry").Value

            ContP = Payment(DocEntry)

            ContORIN = ORINA(DocNum)

            If ContP > 0 Then

                cSBOApplication.MessageBox("Cancelación de Pagos exitosa")
                ContP = 0

            End If

            If ContORIN > 0 Then

                cSBOApplication.MessageBox("Creación de NC exitosa")
                ContORIN = 0
                ORDR(DocNum, CardCode, psDirectory)

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al validar Factura: " & ex.Message)

        End Try

    End Function


    Public Function Payment(ByVal DocEntry As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim vPay As SAPbobsCOM.Payments
        Dim DocEntryP As String
        Dim contador As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        contador = 0

        Try

            stQueryH = "Select T1.""DocEntry"" from RCT2 T0 inner join ORCT T1 on T1.""DocEntry""=T0.""DocNum"" where T0.""DocEntry""=" & DocEntry & " and T0.""InvType""=13"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    vPay = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    DocEntryP = oRecSetH.Fields.Item("DocEntry").Value
                    vPay.GetByKey(DocEntryP)

                    If vPay.Cancel() = 0 Then

                        contador = contador + 1

                    End If

                    oRecSetH.MoveNext()

                Next

            End If

            Return contador

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Cancelar el Pago: " & ex.Message)

        End Try

    End Function


    Public Function ORINA(ByVal DocNum As String)

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim oORINA As SAPbobsCOM.Documents
        Dim DocEntry, DocCur, Folio, OINVSeries As String
        Dim llError As Long
        Dim lsError As String
        Dim contador As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        contador = 0

        Try

            stQueryH = "Select T0.""DocEntry"",T0.""Series"",T0.""CardCode"",T0.""SlpCode"",T0.""Project"",T0.""DocTotal"",T0.""DocCur"",T0.""DocType"",T1.""ReportID"",""VatSum"",T2.""U_B1SYS_MainUsage"" from OINV T0 Left Outer Join ECM2 T1 on T1.""SrcObjAbs""=T0.""DocEntry"" and T1.""SrcObjType""=T0.""ObjType"" where T0.""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oORINA = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oRecSetH.MoveFirst()

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                OINVSeries = oRecSetH.Fields.Item("Series").Value
                DocCur = oRecSetH.Fields.Item("DocCur").Value

                oORINA.Series = 6
                oORINA.CardCode = oRecSetH.Fields.Item("CardCode").Value
                oORINA.DocDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORINA.DocDueDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORINA.DocTotal = oRecSetH.Fields.Item("DocTotal").Value
                oORINA.SalesPersonCode = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.DocumentsOwner = 38
                oORINA.Indicator = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.UserFields.Fields.Item("U_WhsCodeC").Value = oRecSetH.Fields.Item("Project").Value
                oORINA.UserFields.Fields.Item("U_B1SYS_MainUsage").Value = "G02"

                If oRecSetH.Fields.Item("DocType").Value = "I" Then

                    oORINA.DocType = 0

                ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                    oORINA.DocType = 1

                End If

                Folio = oRecSetH.Fields.Item("ReportID").Value

                If Folio = Nothing Or Folio = "" Then

                    oORINA.EDocGenerationType = 2

                Else

                    oORINA.EDocGenerationType = 0
                    'oORINA.EDocExportFormat = 0
                    'oORINA.ElectronicProtocols.MappingID = 49

                End If

                stQueryH2 = "Select T0.""ObjType"",T0.""LineNum"",T0.""ItemCode"",T0.""Dscription"",T0.""Price"",T0.""Quantity"",T0.""TaxCode"",T0.""WhsCode"",T0.""Project"",T0.""DiscPrcnt"" from INV1 T0 where ""DocEntry""=" & DocEntry & " order by T0.""LineNum"""
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    oRecSetH2.MoveFirst()

                    For l = 0 To oRecSetH2.RecordCount - 1

                        If oRecSetH.Fields.Item("DocType").Value = "I" Then

                            oORINA.Lines.ItemCode = oRecSetH2.Fields.Item("ItemCode").Value

                        ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                            oORINA.Lines.ItemDescription = oRecSetH2.Fields.Item("Dscription").Value

                        End If

                        oORINA.Lines.BaseType = oRecSetH2.Fields.Item("ObjType").Value
                        oORINA.Lines.BaseLine = oRecSetH2.Fields.Item("LineNum").Value
                        oORINA.Lines.BaseEntry = DocEntry
                        oORINA.Lines.UnitPrice = oRecSetH2.Fields.Item("Price").Value
                        oORINA.Lines.Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        oORINA.Lines.TaxCode = oRecSetH2.Fields.Item("TaxCode").Value
                        oORINA.Lines.WarehouseCode = oRecSetH2.Fields.Item("WhsCode").Value
                        oORINA.Lines.ProjectCode = oRecSetH2.Fields.Item("Project").Value
                        oORINA.Lines.DiscountPercent = oRecSetH2.Fields.Item("DiscPrcnt").Value
                        oORINA.Lines.Currency = DocCur

                        stQueryH3 = "Select T1.""BatchNum"",T1.""Quantity"" from IBT1 T1 where T1.""BaseType""=" & oRecSetH2.Fields.Item("ObjType").Value & " and T1.""BaseEntry""=" & DocEntry & " And T1.""BaseLinNum""=" & oRecSetH2.Fields.Item("LineNum").Value & " And T1.""ItemCode""='" & oRecSetH2.Fields.Item("ItemCode").Value & "'"
                        oRecSetH3.DoQuery(stQueryH3)

                        If oRecSetH3.RecordCount > 0 Then

                            oRecSetH3.MoveFirst()

                            For z = 0 To oRecSetH3.RecordCount - 1

                                oORINA.Lines.BatchNumbers.BatchNumber = oRecSetH3.Fields.Item("BatchNum").Value
                                oORINA.Lines.BatchNumbers.Quantity = oRecSetH3.Fields.Item("Quantity").Value
                                oORINA.Lines.BatchNumbers.Notes = oRecSetH3.Fields.Item("BatchNum").Value
                                oORINA.Lines.BatchNumbers.BaseLineNumber = oRecSetH2.Fields.Item("LineNum").Value

                                oORINA.Lines.BatchNumbers.Add()

                                oRecSetH3.MoveNext()

                            Next

                        End If

                        oORINA.Lines.Add()

                        oRecSetH2.MoveNext()

                    Next

                End If


                If oORINA.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                Else

                    AORIN = cSBOCompany.GetNewObjectKey().ToString()
                    contador = contador + 1
                    'ContOBNK = OBNK(DocNum)

                End If

            End If

            Return contador

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Crear Nota de Crédito: " & ex.Message)

        End Try

    End Function


    Public Function ORDR(ByVal DocNum As String, ByVal CardCode As String, ByVal psDirectory As String)

        Dim stQueryH, stQueryH2, stQueryH3 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3 As SAPbobsCOM.Recordset
        Dim oORDR As SAPbobsCOM.Documents
        Dim DocEntry, DocCur, OINVSeries, MainUsage As String
        Dim llError As Long
        Dim lsError, OrderDocNum As String
        Dim contador, Order As Integer
        Dim DocTotal As Double

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        contador = 0

        Try

            stQueryH = "Select T0.""DocEntry"",T3.""Series"",T4.""U_B1SYS_MainUsage"",T0.""TrnspCode"",T0.""CardCode"",T0.""SlpCode"",T0.""Project"",T0.""DocTotal"",T0.""DocCur"",T0.""DocType"",T1.""ReportID"",T0.""NumAtCard"" from OINV T0 Left Outer Join ECM2 T1 on T1.""SrcObjAbs""=T0.""DocEntry"" and T1.""SrcObjType""=T0.""ObjType"" Inner Join NNM1 T2 on T2.""Series""=T0.""Series"" and T2.""ObjectCode""=T0.""ObjType"" Inner Join NNM1 T3 on T3.""SeriesName""=T2.""SeriesName"" and T3.""ObjectCode""=17 Inner Join OCRD T4 on T4.""CardCode""=T0.""CardCode"" where T0.""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oORDR = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                oRecSetH.MoveFirst()

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                OINVSeries = oRecSetH.Fields.Item("Series").Value
                MainUsage = oRecSetH.Fields.Item("U_B1SYS_MainUsage").Value
                DocCur = oRecSetH.Fields.Item("DocCur").Value
                DocTotal = oRecSetH.Fields.Item("DocTotal").Value

                oORDR.Series = OINVSeries
                oORDR.CardCode = CardCode
                oORDR.DocDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORDR.DocDueDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORDR.DocTotal = DocTotal
                oORDR.NumAtCard = oRecSetH.Fields.Item("NumAtCard").Value

                oORDR.SalesPersonCode = oRecSetH.Fields.Item("SlpCode").Value
                oORDR.DocumentsOwner = 60
                oORDR.Indicator = oRecSetH.Fields.Item("SlpCode").Value
                oORDR.UserFields.Fields.Item("U_WhsCodeC").Value = oRecSetH.Fields.Item("Project").Value
                oORDR.UserFields.Fields.Item("U_B1SYS_MainUsage").Value = MainUsage
                oORDR.TransportationCode = 3

                If oRecSetH.Fields.Item("DocType").Value = "I" Then

                    oORDR.DocType = 0

                ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                    oORDR.DocType = 1

                End If

                stQueryH2 = "Select T0.""ObjType"",T0.""LineNum"",T0.""ItemCode"",T0.""Dscription"",T0.""Price"",T0.""Quantity"",T0.""TaxCode"",T0.""WhsCode"",T0.""Project"",T0.""DiscPrcnt"" from INV1 T0 where ""DocEntry""=" & DocEntry & " order by T0.""LineNum"""
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    oRecSetH2.MoveFirst()

                    For l = 0 To oRecSetH2.RecordCount - 1

                        If oRecSetH.Fields.Item("DocType").Value = "I" Then

                            oORDR.Lines.ItemCode = oRecSetH2.Fields.Item("ItemCode").Value

                        ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                            oORDR.Lines.ItemDescription = oRecSetH2.Fields.Item("Dscription").Value

                        End If

                        oORDR.Lines.UnitPrice = oRecSetH2.Fields.Item("Price").Value
                        oORDR.Lines.Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        oORDR.Lines.TaxCode = oRecSetH2.Fields.Item("TaxCode").Value
                        oORDR.Lines.WarehouseCode = oRecSetH2.Fields.Item("WhsCode").Value
                        oORDR.Lines.ProjectCode = oRecSetH2.Fields.Item("Project").Value
                        oORDR.Lines.DiscountPercent = oRecSetH2.Fields.Item("DiscPrcnt").Value
                        oORDR.Lines.Currency = DocCur

                        oORDR.Lines.Add()

                        oRecSetH2.MoveNext()

                    Next

                End If

                'oORDR.DocTotal = oRecSetH.Fields.Item("DocTotal").Value

                If oORDR.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                Else

                    Order = cSBOCompany.GetNewObjectKey().ToString()

                    stQueryH3 = "Select ""DocNum"" from ORDR where ""DocEntry""=" & Order
                    oRecSetH3.DoQuery(stQueryH3)

                    OrderDocNum = oRecSetH3.Fields.Item("DocNum").Value

                    cSBOApplication.MessageBox("Se creo con exito la Orden " & OrderDocNum)

                    UpdateSN(CardCode, DocTotal)
                    UpdateORINA(OrderDocNum)
                    CreateOINV(DocNum, Order, CardCode, psDirectory)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Crear Orden de Venta: " & ex.Message)

        End Try

    End Function

    Public Function CreateOINV(ByVal DocNum As String, ByVal Order As String, ByVal CardCode As String, ByVal psDirectory As String)

        Dim stQueryH, stQueryH2, stQueryH3, stQueryH4 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim oOINV As SAPbobsCOM.Documents
        Dim DocEntry, DocCur, Folio, OINVSeries, MainUsage As String
        Dim llError As Long
        Dim lsError As String
        Dim contador, AOINV As Integer
        Dim oED As FrmtekEDocuments

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH4 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        contador = 0

        Try

            stQueryH = "Select T0.""DocEntry"",T2.""Series"",T3.""U_B1SYS_MainUsage"",T0.""TrnspCode"",T0.""CardCode"",T0.""SlpCode"",T0.""Project"",T0.""DocTotal"",T0.""DocCur"",T0.""DocType"",T1.""ReportID"",T0.""NumAtCard"" from OINV T0 Left Outer Join ECM2 T1 on T1.""SrcObjAbs""=T0.""DocEntry"" and T1.""SrcObjType""=T0.""ObjType"" Inner Join NNM1 T2 on T2.""Series""=T0.""Series"" and T2.""ObjectCode""=T0.""ObjType"" Inner Join OCRD T3 on T3.""CardCode""=T0.""CardCode"" where T0.""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oOINV = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oRecSetH.MoveFirst()

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                OINVSeries = oRecSetH.Fields.Item("Series").Value
                MainUsage = oRecSetH.Fields.Item("U_B1SYS_MainUsage").Value
                DocCur = oRecSetH.Fields.Item("DocCur").Value

                oOINV.Series = OINVSeries
                oOINV.CardCode = CardCode
                oOINV.DocDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oOINV.DocDueDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oOINV.DocTotal = oRecSetH.Fields.Item("DocTotal").Value
                oOINV.PaymentGroupCode = 61
                oOINV.NumAtCard = oRecSetH.Fields.Item("NumAtCard").Value

                oOINV.SalesPersonCode = oRecSetH.Fields.Item("SlpCode").Value
                oOINV.DocumentsOwner = 60
                oOINV.Indicator = oRecSetH.Fields.Item("SlpCode").Value
                oOINV.UserFields.Fields.Item("U_WhsCodeC").Value = oRecSetH.Fields.Item("Project").Value
                oOINV.UserFields.Fields.Item("U_B1SYS_MainUsage").Value = MainUsage
                oOINV.TransportationCode = 3

                If oRecSetH.Fields.Item("DocType").Value = "I" Then

                    oOINV.DocType = 0

                ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                    oOINV.DocType = 1

                End If

                Folio = oRecSetH.Fields.Item("ReportID").Value

                'If Folio = Nothing Or Folio = "" Then

                '    oORINA.EDocGenerationType = 2

                'Else

                '    oORINA.EDocGenerationType = 0
                '    oORINA.EDocExportFormat = 0
                '    oORINA.ElectronicProtocols.MappingID = 49

                'End If

                stQueryH2 = "Select T0.""ObjType"",T0.""LineNum"",T0.""ItemCode"",T0.""Dscription"",T0.""Price"",T0.""Quantity"",T0.""TaxCode"",T0.""WhsCode"",T0.""Project"",T0.""DiscPrcnt"" from INV1 T0 where ""DocEntry""=" & DocEntry & " order by T0.""LineNum"""
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    oRecSetH2.MoveFirst()

                    For l = 0 To oRecSetH2.RecordCount - 1

                        If oRecSetH.Fields.Item("DocType").Value = "I" Then

                            oOINV.Lines.ItemCode = oRecSetH2.Fields.Item("ItemCode").Value

                        ElseIf oRecSetH.Fields.Item("DocType").Value = "S" Then

                            oOINV.Lines.ItemDescription = oRecSetH2.Fields.Item("Dscription").Value

                        End If

                        oOINV.Lines.BaseType = 17
                        oOINV.Lines.BaseLine = oRecSetH2.Fields.Item("LineNum").Value
                        oOINV.Lines.BaseEntry = Order
                        oOINV.Lines.UnitPrice = oRecSetH2.Fields.Item("Price").Value
                        oOINV.Lines.Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        oOINV.Lines.TaxCode = oRecSetH2.Fields.Item("TaxCode").Value
                        oOINV.Lines.WarehouseCode = oRecSetH2.Fields.Item("WhsCode").Value
                        oOINV.Lines.ProjectCode = oRecSetH2.Fields.Item("Project").Value
                        oOINV.Lines.DiscountPercent = oRecSetH2.Fields.Item("DiscPrcnt").Value
                        oOINV.Lines.Currency = DocCur

                        stQueryH3 = "Select T1.""BatchNum"",T1.""Quantity"" from IBT1 T1 where T1.""BaseType""=" & oRecSetH2.Fields.Item("ObjType").Value & " and T1.""BaseEntry""=" & DocEntry & " And T1.""BaseLinNum""=" & oRecSetH2.Fields.Item("LineNum").Value & " And T1.""ItemCode""='" & oRecSetH2.Fields.Item("ItemCode").Value & "'"
                        oRecSetH3.DoQuery(stQueryH3)

                        If oRecSetH3.RecordCount > 0 Then

                            For z = 0 To oRecSetH3.RecordCount - 1

                                oOINV.Lines.BatchNumbers.BatchNumber = oRecSetH3.Fields.Item("BatchNum").Value
                                oOINV.Lines.BatchNumbers.Quantity = oRecSetH3.Fields.Item("Quantity").Value
                                oOINV.Lines.BatchNumbers.Notes = oRecSetH3.Fields.Item("BatchNum").Value
                                oOINV.Lines.BatchNumbers.BaseLineNumber = oRecSetH2.Fields.Item("LineNum").Value

                                oOINV.Lines.BatchNumbers.Add()

                                oRecSetH3.MoveNext()

                            Next

                        End If

                        oOINV.Lines.Add()

                        oRecSetH2.MoveNext()

                    Next

                End If

                If oOINV.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                Else

                    AOINV = cSBOCompany.GetNewObjectKey().ToString()

                    stQueryH4 = "Select ""DocNum"" from OINV where ""DocEntry""=" & AOINV
                    oRecSetH4.DoQuery(stQueryH4)

                    cSBOApplication.MessageBox("Se creo con exito la Factura " & oRecSetH4.Fields.Item("DocNum").Value)

                    DeactivateSN(CardCode)

                    oED = New FrmtekEDocuments
                    oED.openForm(psDirectory, AORIN, AOINV)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Crear la Factura: " & ex.Message)

        End Try

    End Function


    Public Function UpdateSN(ByVal CardCode As String, ByVal DocTotal As Double)

        Dim oOCRD As SAPbobsCOM.BusinessPartners
        Dim llError As Long
        Dim lsError As String

        Try

            oOCRD = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If (oOCRD.GetByKey(CardCode) = True) Then

                oOCRD.CreditLimit = DocTotal
                oOCRD.MaxCommitment = DocTotal

                If oOCRD.Update <> 0 Then
                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)
                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al actualizar SN: " & ex.Message)

        End Try

    End Function


    Public Function DeactivateSN(ByVal CardCode As String)

        Dim oOCRD As SAPbobsCOM.BusinessPartners
        Dim llError As Long
        Dim lsError As String

        Try

            oOCRD = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If (oOCRD.GetByKey(CardCode) = True) Then

                oOCRD.CreditLimit = 0
                oOCRD.MaxCommitment = 0
                oOCRD.UserFields.Fields.Item("U_EcommerceAct").Value = "N"

                If oOCRD.Update <> 0 Then
                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)
                Else
                    cSBOApplication.MessageBox("Cliente desactivado.")
                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al desactivar SN: " & ex.Message)

        End Try

    End Function


    Public Function UpdateORINA(ByVal OrderDocNum As String)

        Dim oORINA As SAPbobsCOM.Documents
        Dim llError As Long
        Dim lsError As String

        Try

            oORINA = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)

            If (oORINA.GetByKey(AORIN) = True) Then

                oORINA.UserFields.Fields.Item("U_ORDR").Value = OrderDocNum

                If oORINA.Update() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Actualizar Nota de Crédito: " & ex.Message)

        End Try

    End Function


End Class
