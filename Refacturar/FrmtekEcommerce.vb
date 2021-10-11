Public Class FrmtekEcommerce

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
    Public Function openForm(ByVal psDirectory As String)
        ', ByVal CardCode As String, ByVal DocNum As String
        Try

            csFormUID = "BusinessPartner"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")

            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            cargarComboSNE()

            'AgregarPrecios(CardCode, csFormUID)

            'HideOrShowFormItems(DocNum, csFormUID)

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
        Dim oComboBox As SAPbouiCOM.ComboBox

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("dsSNE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT) 'Creo el datasources
            oComboBox = coForm.Items.Item("2").Specific  'identifico mi combobox
            oComboBox.DataBind.SetBound(True, "", "dsSNE")   ' uno mi userdatasources a mi combobox

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


    Public Function cargarComboSNE()

        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecSet As SAPbobsCOM.Recordset

        Try
            cargarComboSNE = 0
            '--- referencia de combo 
            oCombo = coForm.Items.Item("2").Specific
            coForm.Freeze(True)
            '---- SI YA SE TIENEN VALORES, SE ELIMMINAN DEL COMBO
            If oCombo.ValidValues.Count > 0 Then
                Do
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Loop While oCombo.ValidValues.Count > 0
            End If
            '--- realizar consulta
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSet.DoQuery("Call BusinessPartnerEcommerce")
            '---- cargamos resultado
            oRecSet.MoveFirst()
            Do While oRecSet.EoF = False
                oCombo.ValidValues.Add(oRecSet.Fields.Item(0).Value, oRecSet.Fields.Item(1).Value)
                oRecSet.MoveNext()
            Loop
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            coForm.Freeze(False)


        Catch ex As Exception
            coForm.Freeze(False)
            MsgBox("FrmTratamientoPedidos. cargarComboPorcentaje: " & ex.Message)
        Finally
            oCombo = Nothing
            oRecSet = Nothing
        End Try
    End Function


    Public Function closeC(ByVal psFormUID As String) As Integer
        closeC = 0
        coForm = cSBOApplication.Forms.Item(psFormUID)
        coForm.Close()
    End Function


End Class
