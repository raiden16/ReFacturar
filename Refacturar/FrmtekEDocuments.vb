Public Class FrmtekEDocuments

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String, ByVal AORIN As Integer, ByVal AOINV As Integer)
        ', ByVal CardCode As String, ByVal DocNum As String
        Try

            csFormUID = "ElectronicDocuments"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")

            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            AgregarEDocuments(csFormUID, AORIN, AOINV)

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            oGrid = coForm.Items.Item("2").Specific
            oDataTable = coForm.DataSources.DataTables.Add("EDORIN")
            oGrid.DataTable = oDataTable

            oGrid = coForm.Items.Item("3").Specific
            oDataTable = coForm.DataSources.DataTables.Add("EDOINV")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '----- carga los ultimos precios de entrega segun el proveedor
    Public Function AgregarEDocuments(ByVal psFormUID As String, ByVal AORIN As Integer, ByVal AOINV As Integer)
        Dim oGrid, oGrid2 As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim oColumn, oColumn2 As SAPbouiCOM.GridColumn

        Try

            coForm = cSBOApplication.Forms.Item(psFormUID)

            '///////////////////////////////////////GRID ORIN
            oGrid = coForm.Items.Item("2").Specific
            oGrid.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select ""DocEntry"",""DocNum"",""CardCode"",""CardName"",""DocTotal"" from ORIN where ""DocEntry""=" & AORIN
            oGrid.DataTable.ExecuteQuery(stQuery)

            oColumn = oGrid.Columns.Item("DocEntry")
            oColumn.LinkedObjectType = 14

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False
            oGrid.Columns.Item(4).Editable = False

            '///////////////////////////////////////GRID OINV
            oGrid2 = coForm.Items.Item("3").Specific
            oGrid2.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select ""DocEntry"",""DocNum"",""CardCode"",""CardName"",""DocTotal"" from OINV where ""DocEntry""=" & AOINV
            oGrid2.DataTable.ExecuteQuery(stQuery)

            oColumn2 = oGrid2.Columns.Item("DocEntry")
            oColumn2.LinkedObjectType = 13

            oGrid2.Columns.Item(0).Editable = False
            oGrid2.Columns.Item(1).Editable = False
            oGrid2.Columns.Item(2).Editable = False
            oGrid2.Columns.Item(3).Editable = False
            oGrid2.Columns.Item(4).Editable = False

            Return 0

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function


End Class
