Imports System.Xml
Imports SAPbouiCOM

Public Class EXO_CNTRLPR
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Shared _sOportunidad As String = ""
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
            'ParametrizacionGeneral()
        End If
        cargamenu()
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
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")

        Try
            objGlobal.SBOApp.LoadBatchActions(menuXML)
            Dim res As String = objGlobal.SBOApp.GetLastBatchResults
            'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub cargaCampos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            'Campos de usuario en Factura de clientes
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CNTRLPR.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_CNTRLPR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                    Case "EXO-MnOCP"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CNTRLPR")
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
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CNTRLPR"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

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
                        Case "UDO_FT_EXO_CNTRLPR"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CNTRLPR"
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
                        Case "UDO_FT_EXO_CNTRLPR"
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
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_CNTRLPR"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_CNTRLPR"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    Dim sIntermediario As String = "" : Dim sSQL As String = ""
                                    sIntermediario = CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                                    If sIntermediario <> "" Then
                                        'Cargamos combo de contactos
                                        sSQL = "SELECT ""CntctCode"", ""Name"" FROM ""OCPR"" WHERE ""CardCode""='" & sIntermediario & "' "
                                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                    End If
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
    Private Function EventHandler_VALIDATE_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sIntermediario As String = ""
        Dim bOpTriangulada As Boolean = False
        Dim sSQL As String = ""
        EventHandler_VALIDATE_Before = False
        Try
            bOpTriangulada = CType(oForm.Items.Item("chkOP").Specific, SAPbouiCOM.CheckBox).Checked
            If pVal.ItemUID = "16_U_E" Then
                sIntermediario = CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                If sIntermediario <> "" Then
                    'Cargamos combo de contactos
                    sSQL = "SELECT ""CntctCode"", ""Name"" FROM ""OCPR"" WHERE ""CardCode""='" & sIntermediario & "' "
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                End If
                If sIntermediario = "" And bOpTriangulada = True Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Intermediario no puede estar vacío al ser una operación triangulada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El Intermediario no puede estar vacío al ser una operación triangulada.")
                    EventHandler_VALIDATE_Before = False
                Else
                    EventHandler_VALIDATE_Before = True
                End If
            ElseIf pVal.ItemUID = "cbContact" Then
                If CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sIntermediario = CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                Else
                    sIntermediario = ""
                End If
                If sIntermediario = "" And bOpTriangulada = True Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Contacto del Intermediario no puede estar vacío al ser una operación triangulada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El Contacto del Intermediario no puede estar vacío al ser una operación triangulada.")
                    EventHandler_VALIDATE_Before = False
                Else
                    EventHandler_VALIDATE_Before = True
                End If
            Else
                EventHandler_VALIDATE_Before = True
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sOportunidad As String = "" : Dim sNomOportunidad As String = ""
        Dim sLote As String = "" : Dim sDes As String = ""
        Dim sHojaCoste As String = ""
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos  
        EventHandler_VALIDATE_After = False
        Try
            If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_2" Then
                sOportunidad = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                sNomOportunidad = CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                sLote = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                sDes = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                sHojaCoste = sOportunidad & "_" & sLote
                'Buscamos para ver si existe
                sSQL = "SELECT ""Code"" FROM ""@EXO_OHCOSTES"" WHERE ""Code""='" & sHojaCoste & "' "
                oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sHojaCoste
                    End If
                Else
#Region "Creación Hoja de coste"
                    'Creamos cabecera de la hoja de coste
                    oDI_COM = objGlobal.refDi.dameEXO_UDOEntity("EXO_OHCOSTES")  'UDO de Campos de SAP
                    oDI_COM.GetNew()
                    oDI_COM.SetValue("Code") = sHojaCoste
                    oDI_COM.SetValue("U_EXO_IDOP") = sOportunidad
                    oDI_COM.SetValue("U_EXO_IDNOM") = sNomOportunidad
                    oDI_COM.SetValue("U_EXO_LOTE") = sLote
                    oDI_COM.SetValue("U_EXO_DES") = sDes
                    If oDI_COM.UDO_Add = False Then
                        Throw New Exception("(EXO) - Error al añadir Hoja de Coste """ & sHojaCoste & """. " & oDI_COM.GetLastError)
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se añade la Hoja de coste " & sHojaCoste, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_9").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sHojaCoste
                    End If
#End Region
                End If
                'Buscamos para ver si existe
                sSQL = "SELECT ""Code"" FROM ""@EXO_ORAECO"" WHERE ""Code""='" & sHojaCoste & "' "
                oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sHojaCoste
                    End If
                Else

#Region "Creación Análisis Económico"
                    'Creamos cabecera de la hoja de coste
                    oDI_COM = objGlobal.refDi.dameEXO_UDOEntity("EXO_ORAECO")  'UDO de Campos de SAP
                    oDI_COM.GetNew()
                    oDI_COM.SetValue("Code") = sHojaCoste
                    oDI_COM.SetValue("U_EXO_IDOP") = sOportunidad
                    oDI_COM.SetValue("U_EXO_IDNOM") = sNomOportunidad
                    oDI_COM.SetValue("U_EXO_LOTE") = sLote
                    oDI_COM.SetValue("U_EXO_DES") = sDes
                    If oDI_COM.UDO_Add = False Then
                        Throw New Exception("(EXO) - Error al añadir Análisis Económico """ & sHojaCoste & """. " & oDI_COM.GetLastError)
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se añade Análisis Económico " & sHojaCoste, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_10").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sHojaCoste
                    End If
#End Region
                End If
            End If
            EventHandler_VALIDATE_After = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
#Region "Variables"
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sDes As String = ""
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
#End Region

        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            Select Case pVal.ItemUID
                Case "0_U_E"
                    If oDataTable IsNot Nothing Then
                        Try
                            Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                                Case "97"
                                    Try
                                        sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString
                                        oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("Name", 0, sDes)
                                    Catch ex As Exception
                                        oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("Name", 0, sDes)
                                    End Try
                                    'Buscamos el nombre del cliente
                                    If oForm.Visible = True Then
                                        oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                                        sSQL = "SELECT ""CardCode"",""CardName"" FROM ""OOPR"" T0  where ""OpprId"" ='" & oDataTable.SelectedObjects.GetValue("OpprId", 0).ToString & "' "
                                        oRs.DoQuery(sSQL)
                                        If oRs.RecordCount > 0 Then
                                            oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("U_EXO_ICCOD", 0, oRs.Fields.Item("CardCode").Value.ToString)
                                            oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("U_EXO_ICNOM", 0, oRs.Fields.Item("CardName").Value.ToString)
                                        End If
                                    End If
                            End Select
                            'If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                        Catch ex As Exception
                            Throw ex
                        End Try
                    End If
                Case "16_U_E"
                    If oDataTable IsNot Nothing Then
                        Try
                            Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                                Case "2"
                                    Try
                                        sDes = oDataTable.SelectedObjects.GetValue("CardName", 0).ToString
                                        oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("U_EXO_INTERNOM", 0, sDes)

                                    Catch ex As Exception
                                        oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPR").SetValue("U_EXO_INTERNOM", 0, sDes)
                                    End Try
                            End Select
                            If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                        Catch ex As Exception
                            Throw ex
                        End Try
                    End If
            End Select
            If pVal.ItemUID = "0_U_E" Then

            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim bOpTriangulada As Boolean = False
        Dim sIntermediario As String = ""
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            bOpTriangulada = CType(oForm.Items.Item("chkOP").Specific, SAPbouiCOM.CheckBox).Checked
            Select Case pVal.ItemUID
                Case "1" 'Validamos los campos 
                    If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                        If bOpTriangulada = True Then
                            sIntermediario = CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                            If sIntermediario = "" Then
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Intermediario no puede estar vacío al ser una operación triangulada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox("El Intermediario no puede estar vacío al ser una operación triangulada.")
                                EventHandler_ItemPressed_Before = False
                            Else
                                EventHandler_ItemPressed_Before = True
                            End If
                            If CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                                sIntermediario = CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                            Else
                                sIntermediario = ""
                            End If
                            If sIntermediario = "" Then
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Contacto del Intermediario no puede estar vacío al ser una operación triangulada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox("El Contacto del Intermediario no puede estar vacío al ser una operación triangulada.")
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

                Case "0_U_G"
                    ' Escribimos un Lote por defecto
                    If pVal.ColUID = "C_0_1" Then
                        If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String = "" Then
                            Dim sLinea As String = oForm.DataSources.DBDataSources.Item("@EXO_CNTRLPRL").GetValue("LineId", pVal.Row - 1).Trim
                            CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String = sLinea
                        End If
                    End If
                Case Else
                    EventHandler_ItemPressed_Before = True
            End Select

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
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


        EventHandler_FORM_VISIBLE_After = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            If oForm.Visible = True Then

#Region "Inicializamos valores"
                If oForm.Mode = BoFormMode.fm_ADD_MODE And _sOportunidad <> "" Then
                    ' Si no existe el control de proyecto
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = _sOportunidad
                ElseIf oForm.Mode = BoFormMode.fm_OK_MODE Then
                    Dim sIntermediario As String = "" : Dim sSQL As String = ""
                    sIntermediario = CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                    If sIntermediario <> "" Then
                        'Cargamos combo de contactos
                        sSQL = "SELECT ""CntctCode"", ""Name"" FROM ""OCPR"" WHERE ""CardCode""='" & sIntermediario & "' "
                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbContact").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    End If
                End If
                    oForm.Items.Item("cbContact").DisplayDesc = True
                'Inicializamos valores
                Try
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").DisplayDesc = True
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ExpandType = BoExpandType.et_DescriptionOnly
                    'CargarCombos(oForm)
                Catch ex As Exception
                    'objGlobal.SBOApp.StatusBar.SetText("No se puede inicializar el campo ""Etapa de Anexos"". " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
#End Region
            End If

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
    Private Sub CargarCombos(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""

        Try

            sSQL = "SELECT ""Descript"",""Descript"" FROM ""OOST"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
