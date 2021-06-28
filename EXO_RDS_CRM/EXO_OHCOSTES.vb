Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OHCOSTES
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
            'ParametrizacionGeneral()
        End If
    End Sub
#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

    Private Sub cargaCampos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            'Campos de usuario en Factura de clientes
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OHCOSTES.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_OHCOSTES", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
            'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub

    'Private Sub ParametrizacionGeneral()
    '    If Not objGlobal.refDi.OGEN.existeVariable("EXO_PATH_EDI_FACTURAS") Then
    '        objGlobal.refDi.OGEN.fijarValorVariable("EXO_PATH_EDI_FACTURAS", "\\" & objGlobal.compañia.Server.Split(CChar(":"))(0) & "\B1_SHF\EDIFACT\" & objGlobal.compañia.CompanyDB)
    '    End If
    'End Sub

#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnOHC"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_OHCOSTES")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OHCOSTES"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OHCOSTES"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    CargaCombo_Lote(oForm, "")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OHCOSTES"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OHCOSTES"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OHCOSTES"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OHCOSTES"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_Before = False
        Dim bCamposObligatoriosSuministros As Boolean = False
        Dim bCamposObligatoriosSrvPropios As Boolean = False
        Dim bCamposObligatoriosOtros As Boolean = False
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                'oForm.PaneLevel = 69
                'oForm.Items.Item("fldREDES").Specific.select()
                If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                    'Sumatorios de cabecera
                    Dim dCosteTotal As Double = 0 : Dim dPVPTotal As Double = 0
                    bCamposObligatoriosSuministros = MatrixToNet_Suministros(oForm, dCosteTotal, dPVPTotal)
                    bCamposObligatoriosSrvPropios = MatrixToNet_SRVPropios(oForm, dCosteTotal, dPVPTotal)
                    bCamposObligatoriosOtros = MatrixToNet_Otros(oForm, dCosteTotal, dPVPTotal)
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_CTOTAL", 0, EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dCosteTotal, EXO_GLOBALES.FuenteInformacion.Otros))
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_PVP", 0, EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPVPTotal, EXO_GLOBALES.FuenteInformacion.Otros))
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_MARGEN", 0, EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPVPTotal - dCosteTotal, EXO_GLOBALES.FuenteInformacion.Otros))
                    Dim dPmargen As Double = 0
                    If dPVPTotal = 0 Then
                        dPmargen = 0
                    Else
                        dPmargen = ((dPVPTotal - dCosteTotal) * 100) / dPVPTotal
                        dPmargen = CDbl(String.Format("{0:0.0000}", dPmargen))
                    End If
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_PMARGEN", 0, EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPmargen, EXO_GLOBALES.FuenteInformacion.Otros))

                    'Grabamos en el control de proyectos
                    Dim sCode As String = oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").GetValue("Code", 0)
                    Dim sCodigos() As String = sCode.Split("_")
                    sSQL = "UPDATE ""@EXO_CNTRLPRL"" "
                    sSQL &= " SET ""U_EXO_HOJACOST""='" & sCode & "' "
                    sSQL &= " , ""U_EXO_IMPORTE""=" & EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPVPTotal, EXO_GLOBALES.FuenteInformacion.Otros)
                    sSQL &= " WHERE ""Code""='" & sCodigos(0) & "'  And ""U_EXO_LOTE""=" & sCodigos(1)
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRs.DoQuery(sSQL)

                    If bCamposObligatoriosOtros = False Or bCamposObligatoriosSrvPropios = False Or bCamposObligatoriosSuministros = False Then
                        EventHandler_ItemPressed_Before = False
                    Else
                        EventHandler_ItemPressed_Before = True
                    End If
                Else
                    EventHandler_ItemPressed_Before = True
                End If
            Else
                EventHandler_ItemPressed_Before = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function MatrixToNet_Suministros(ByRef oForm As SAPbouiCOM.Form, ByRef dCosteTotal As Double, ByRef dPVPTotal As Double) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing

        Dim sValor As String = ""
        Dim dCoste As Double = 0 : Dim dPVP As Double = 0
        Dim bCamposObligatorios As Boolean = True
        Dim sMatrixUID As String = ""

        MatrixToNet_Suministros = False

        Try
            sXML = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")


                'Aqui inicializamos los datos de registro si hace falta

                '___________________________________________________________

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "C_0_6" Then 'Coste Total
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        dCoste = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)

                    ElseIf oXmlNodeField.InnerXml = "C_0_11" Then 'PVP
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        dPVP = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)
                    ElseIf oXmlNodeField.InnerXml = "C_0_1" Then 'Capitulo
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El capítulo en la pestaña ""Suministros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("El capítulo en la pestaña ""Suministros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    ElseIf oXmlNodeField.InnerXml = "C_0_2" Then 'partida
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La partida en la pestaña ""Suministros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("La partida en la pestaña ""Suministros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    ElseIf oXmlNodeField.InnerXml = "C_0_3" Then 'Descripción
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La descripción en la pestaña ""Suministros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("La Descripción en la pestaña ""Suministros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    ElseIf oXmlNodeField.InnerXml = "C_0_7" Then 'Fabricante
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El fabricante en la pestaña ""Suministros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("El fabricante en la pestaña ""Suministros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    ElseIf oXmlNodeField.InnerXml = "C_0_9" Then 'Carencia de pago
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La carencia de pago en la pestaña ""Suministros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("La carencia de pago en la pestaña ""Suministros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    End If
                Next
                dCosteTotal += dCoste
                dPVPTotal += dPVP
            Next

            If bCamposObligatorios = False Then
                MatrixToNet_Suministros = False
            Else
                MatrixToNet_Suministros = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Function MatrixToNet_SRVPropios(ByRef oForm As SAPbouiCOM.Form, ByRef dCosteTotal As Double, ByRef dPVPTotal As Double) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing

        Dim sValor As String = ""
        Dim dCoste As Double = 0 : Dim dPVP As Double = 0

        Dim sMatrixUID As String = ""
        Dim bCamposObligatorios As Boolean = True
        MatrixToNet_SRVPropios = False

        Try
            sXML = CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")

                'Aqui inicializamos los datos de registro si hace falta

                '___________________________________________________________

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "C_1_5" Then 'Coste Total
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        dCoste = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)
                    ElseIf oXmlNodeField.InnerXml = "C_1_7" Then 'PVP
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        dPVP = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)
                    ElseIf oXmlNodeField.InnerXml = "C_1_1" Then 'Departamento
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Departamento en la pestaña ""Servicios"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("El concepto en la pestaña ""Servicios"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    ElseIf oXmlNodeField.InnerXml = "C_1_3" Then 'Descripción
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La Descripción en la pestaña ""Servicios"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("La Descripción en la pestaña ""Servicios"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    End If
                Next
                dCosteTotal += dCoste
                dPVPTotal += dPVP
            Next

            If bCamposObligatorios = False Then
                MatrixToNet_SRVPropios = False
            Else
                MatrixToNet_SRVPropios = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Function MatrixToNet_Otros(ByRef oForm As SAPbouiCOM.Form, ByRef dCosteTotal As Double, ByRef dPVPtotal As Double) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing

        Dim sValor As String = ""
        Dim dCoste As Double = 0 : Dim dPVP As Double = 0
        Dim sMatrixUID As String = ""
        Dim bCamposObligatorios As Boolean = True
        MatrixToNet_Otros = False

        Try
            sXML = CType(oForm.Items.Item("2_U_G").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")


                'Aqui inicializamos los datos de registro si hace falta

                '___________________________________________________________

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "C_2_2" Then 'Coste Total
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        dCoste = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)
                        dPVP = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oXmlNodeField.InnerText)
                    ElseIf oXmlNodeField.InnerXml = "C_2_1" Then 'Concepto
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")
                        sValor = oXmlNodeField.InnerText
                        'If sValor = "" Then
                        '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El concepto en la pestaña ""Otros"", no puede estar vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objGlobal.SBOApp.MessageBox("El concepto en la pestaña ""Otros"", no puede estar vacío.")
                        '    bCamposObligatorios = False
                        'End If
                    End If
                Next
                dCosteTotal += dCoste
                dPVPtotal += dPVP
            Next

            If bCamposObligatorios = False Then
                MatrixToNet_Otros = False
            Else
                MatrixToNet_Otros = True
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim dCantidad As Double = 0 : Dim dCoste As Double = 0 : Dim dCosteT As Double = 0 : Dim dMargen As Double = 0 : Dim dMargenT As Double = 0
        Dim sMayorista As String = "" : Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_VALIDATE_After = False
        Try
            oForm.Freeze(True)
            If pVal.ItemUID = "0_U_G" And (pVal.ColUID = "C_0_4" Or pVal.ColUID = "C_0_5" Or pVal.ColUID = "C_0_10") Then
                dCantidad = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)
                dCoste = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)

                dMargen = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)

                dCosteT = dCantidad * dCoste : dMargenT = dCosteT + (dCosteT * dMargen) / 100
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_11").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dMargenT.ToString, EXO_GLOBALES.FuenteInformacion.Otros)
                CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dCosteT.ToString, EXO_GLOBALES.FuenteInformacion.Otros)
                oForm.PaneLevel = 2 : oForm.PaneLevel = 1
                Select Case pVal.ColUID
                    Case "C_0_4" : CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                    Case "C_0_5" : CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                    Case "C_0_10" : CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                End Select

            ElseIf pVal.ItemUID = "0_U_G" And (pVal.ColUID = "C_0_8") Then
                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_8").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString <> "" Then
                    sMayorista = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_8").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                    'Carencia de pago
                    sSQL &= "SELECT G.""PymntGroup"" ""CodPago"" FROM ""OCTG"" G INNER JOIN ""OCRD"" C ON G.""GroupNum""=C.""GroupNum"" and C.""CardName""='" & sMayorista & "' "
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 And oRs.Fields.Item("CodPago").Value.ToString <> "" Then
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Select(oRs.Fields.Item("CodPago").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            ElseIf pVal.ItemUID = "1_U_G" And (pVal.ColUID = "C_1_4" Or pVal.ColUID = "C_1_1" Or pVal.ColUID = "C_1_6") Then
                Calcular_SRVPropios(oForm, pVal)
                oForm.PaneLevel = 1 : oForm.PaneLevel = 2
            End If
            EventHandler_VALIDATE_After = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub Calcular_SRVPropios(ByRef oForm As SAPbouiCOM.Form, ByRef pVal As ItemEvent)
        Dim dCantidad As Double = 0 : Dim dCoste As Double = 0 : Dim dCosteT As Double = 0 : Dim dMargen As Double = 0 : Dim dMargenT As Double = 0
        Try
            oForm.Freeze(True)
            dCantidad = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)
            dCoste = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)
            dMargen = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_6").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)

            If dMargen < 15 Then
                'Se visualiza un mensaje
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El margen es inferior a 15%. Debe estar sometido a la dirección", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox("El margen es inferior a 15%. Debe estar sometido a la dirección")
            End If

            dCosteT = dCantidad * dCoste : dMargenT = dCosteT + (dCosteT * dMargen) / 100
            CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_7").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dMargenT.ToString, EXO_GLOBALES.FuenteInformacion.Otros)
            CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_5").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dCosteT.ToString, EXO_GLOBALES.FuenteInformacion.Otros)

        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sOportunidad As String = "" : Dim sLote As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oForm.Freeze(True)
            If pVal.ItemUID = "cbLote" Then 'Combo Lote
                sOportunidad = oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").GetValue("U_EXO_IDOP", 0).ToString
                sLote = oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").GetValue("U_EXO_LOTE", 0).ToString

                'Ponemos la descripción
                sSQL = "Select DISTINCT ""U_EXO_DES"" FROM ""@EXO_CNTRLPRL"" WHERE ""Code""='" & sOportunidad & "' and ""U_EXO_LOTE""=" & sLote
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_DES", 0, oRs.Fields.Item("U_EXO_DES").Value.ToString)
                End If

                ' Tenemos que poner el código de la hoja de coste
                If sOportunidad <> "" And sLote <> "" Then
                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("Code", 0, sOportunidad & "_" & sLote)
                End If
            ElseIf pVal.ItemUID = "1_U_G" And pVal.ColUID = "C_1_1" Then 'Combo Departamento en la pestaña de Servicios Propios
                If CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    Dim sDepartamento As String = CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    Dim sfechaActual As String = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
                    sSQL = "SELECT  L.""U_EXO_COSTE"" FROM ""@EXO_COSTEDPTOL"" L INNER JOIN ""@EXO_COSTEDPTO"" D ON D.""Code""=L.""Code"" WHERE D.""Code""='" & sDepartamento & "'"
                    sSQL &= " and ""U_EXO_FINI""<='" & sfechaActual & "' and ifnull(""U_EXo_FFIN"",'20501231')>='" & sfechaActual & "' "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        Dim dCoste As Double = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, oRs.Fields.Item("U_EXO_COSTE").Value.ToString)

                        CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dCoste, EXO_GLOBALES.FuenteInformacion.Otros)
                    Else
                        CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = 0
                    End If
                    Calcular_SRVPropios(oForm, pVal)
                    oForm.PaneLevel = 1
                    oForm.PaneLevel = 2
                    CType(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim oFolder As SAPbouiCOM.Folder = Nothing
        Dim oNewItem As SAPbouiCOM.Item = Nothing
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = Nothing
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams = objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        Dim oCFL As SAPbouiCOM.ChooseFromList = Nothing
        Dim sSQL As String = ""
        EventHandler_FORM_VISIBLE_After = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            'objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)

#Region "Inicializamos valores"
            'Inicializamos valores
            If oForm.Visible = True Then
                'El lote no se puede modificar
                oForm.Items.Item("cbLote").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("cbLote").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oForm.Items.Item("cbLote").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
#Region "Pestaña otros"
                'Departamento
                sSQL = " SELECT  T0.""Code"", T0.""Name"" FROM ""@EXO_CONCEPTOS"" T0 "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("2_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_2_1").ValidValues, sSQL)
#End Region
#Region "Pestaña SERVICIOS Propios"
                'Departamento
                sSQL = "SELECT 'No Procede' AS ""Departamento"" FROM DUMMY "
                sSQL &= " UNION ALL "
                sSQL &= " SELECT  T0.""Name"" FROM OUDP T0 "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_1").ValidValues, sSQL)
                CType(oForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_1_1").ExpandType = BoExpandType.et_DescriptionOnly
#End Region
#Region "Pestaña SUM./SERV. Externos"
                'Fabricante
                sSQL = " SELECT * FROM ("
                sSQL &= " (SELECT '- No identificado –' ""Fabricantes"" FROM DUMMY ) "
                sSQL &= " UNION ALL "
                sSQL &= "(SELECT ""FirmName"" ""Fabricantes"" FROM ""OMRC"" ORDER BY ""FirmName"") "
                'sSQL &= " UNION ALL "
                'sSQL &= " (SELECT ""CardName""  ""Fabricantes"" FROM ""OCRD"" WHERE ""CardType""='S' ORDER BY ""CardName"") "
                sSQL &= ")T ORDER BY ""Fabricantes"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_7").ValidValues, sSQL)
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_7").ExpandType = BoExpandType.et_DescriptionOnly

                ''Mayorista
                'sSQL = " SELECT * FROM ("
                'sSQL &= " (SELECT ' ' ""Mayorista"" FROM DUMMY ) "
                'sSQL &= " UNION ALL "
                'sSQL &= " (SELECT ""CardName""  ""Mayorista"" FROM ""OCRD"" WHERE ""CardType""='S' ORDER BY ""CardName"") "
                'sSQL &= ")T ORDER BY ""Mayorista"" "
                'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_8").ValidValues, sSQL)
                'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_8").ExpandType = BoExpandType.et_DescriptionOnly

                'Carencia de pago
                sSQL = " (SELECT ""PymntGroup"" FROM ""OCTG""  ORDER BY ""PymntGroup"") "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").ValidValues, sSQL)
                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").ExpandType = BoExpandType.et_DescriptionOnly
#End Region


                'Activo Id Oportunidad
                CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Active = True
            End If
#End Region
            EventHandler_FORM_VISIBLE_After = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try

            If pVal.ItemUID = "13_U_E" Then
                Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "97"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("OpprId", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_IDNOM", 0, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").SetValue("U_EXO_IDNOM", 0, sDes)
                                End Try
                                'cargamos los lotes
                                CargaCombo_Lote(oForm, sCod)
                        End Select
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub CargaCombo_Lote(ByRef oForm As SAPbouiCOM.Form, ByVal sOportunidad As String)
        Dim sSQL As String = ""

        Try
            If sOportunidad = "" Then
                sOportunidad = oForm.DataSources.DBDataSources.Item("@EXO_OHCOSTES").GetValue("U_EXO_IDOP", 0).ToString
            End If
            CType(oForm.Items.Item("cbLote").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
            sSQL = "Select DISTINCT ""U_EXO_LOTE"" ""Codigo"",""U_EXO_LOTE"" ""Lote"" FROM ""@EXO_CNTRLPRL"" WHERE ""Code""='" & sOportunidad & "' "
            sSQL &= " ORDER BY ""U_EXO_LOTE"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbLote").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
