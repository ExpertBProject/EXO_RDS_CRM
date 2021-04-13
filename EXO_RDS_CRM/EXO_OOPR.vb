Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OOPR
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OOPR.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OOPR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "320"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "320"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
                        Case "320"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "320"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "chkMarco"
                    If CType(oForm.Items.Item("chkMarco").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                        oForm.Items.Item("cbOVIN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    Else
                        oForm.Items.Item("cbOVIN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If
                Case "btnCP"
                    sSQL = "SELECT ""Code"" FROM ""@EXO_CNTRLPR"" WHERE ""Code""='" & CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        EXO_CNTRLPR._sOportunidad = ""
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_CNTRLPR", CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString)
                    Else
                        EXO_CNTRLPR._sOportunidad = CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CNTRLPR")
                    End If
                Case "btnEA"
                    sSQL = "SELECT ""Code"" FROM ""@EXO_OPORANEXOS"" WHERE ""Code""='" & CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        EXO_OEANEXOS._sOportunidad = ""
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_OEANEXOS", CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString)
                    Else
                        EXO_OEANEXOS._sOportunidad = CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_OEANEXOS")
                    End If
                Case "btnDE"
                    sSQL = "SELECT ""Code"" FROM ""@EXO_OPORDOCO"" WHERE ""Code""='" & CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        EXO_OOBLI._sOportunidad = ""
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_OOBLI", CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString)
                    Else
                        EXO_OOBLI._sOportunidad = CType(oForm.Items.Item("74").Specific, SAPbouiCOM.EditText).Value.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_OOBLI")
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim oFolder As SAPbouiCOM.Folder = Nothing
        Dim oNewItem As SAPbouiCOM.Item = Nothing
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = Nothing
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams = objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
        Dim oCFL As SAPbouiCOM.ChooseFromList = Nothing

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Visible = False
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
#Region "Campos en cabecera y en pie"
#Region "Tipo Oportunidad"
            oNewItem = oForm.Items.Add("cbTOP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.LinkTo = "137"
            oNewItem.Top = oForm.Items.Item("234000001").Top
            oNewItem.Left = oForm.Items.Item("137").Left
            oNewItem.Height = oForm.Items.Item("137").Height
            oNewItem.Width = oForm.Items.Item("137").Width
            oNewItem.Enabled = True
            oNewItem.DisplayDesc = True
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            CType(oNewItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OOPR", "U_EXO_OPOTIP")
            oNewItem = oForm.Items.Add("lblTOP", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("cbTOP").Top
            oNewItem.Left = oForm.Items.Item("136").Left
            oNewItem.Height = oForm.Items.Item("136").Height
            oNewItem.Width = oForm.Items.Item("136").Width
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            oNewItem.LinkTo = "cbTOP"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Tipo de Oportunidad"
#End Region
#Region "Clasificación"
            oNewItem = oForm.Items.Add("cbClasi", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.LinkTo = "137"
            oNewItem.Top = oForm.Items.Item("86").Top + 2
            oNewItem.Left = oForm.Items.Item("137").Left
            oNewItem.Height = oForm.Items.Item("137").Height
            oNewItem.Width = oForm.Items.Item("137").Width
            oNewItem.Enabled = True
            oNewItem.DisplayDesc = True
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            CType(oNewItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OOPR", "U_EXO_CLASIFICACION")
            oNewItem = oForm.Items.Add("lblClasi", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("cbClasi").Top
            oNewItem.Left = oForm.Items.Item("136").Left
            oNewItem.Height = oForm.Items.Item("136").Height
            oNewItem.Width = oForm.Items.Item("136").Width
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            oNewItem.LinkTo = "cbClasi"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Clasificación"
#End Region
#Region "Acuerdo Marco"
            oNewItem = oForm.Items.Add("chkMarco", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oNewItem.LinkTo = "137"
            oNewItem.Top = oForm.Items.Item("cbClasi").Top + oForm.Items.Item("cbClasi").Height + 2
            oNewItem.Left = oForm.Items.Item("86").Left
            oNewItem.Height = oForm.Items.Item("86").Height
            oNewItem.Width = oForm.Items.Item("86").Width
            oNewItem.Enabled = True
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            CType(oNewItem.Specific, SAPbouiCOM.CheckBox).Caption = "Acuerdo Marco"
            CType(oNewItem.Specific, SAPbouiCOM.CheckBox).DataBind.SetBound(True, "OOPR", "U_EXO_AMARCO")
#End Region
#Region "Oportunidad Vinculada"
            oNewItem = oForm.Items.Add("cbOVIN", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oNewItem.LinkTo = "chkMarco"
            oNewItem.Top = oForm.Items.Item("chkMarco").Top
            oNewItem.Left = oForm.Items.Item("137").Left
            oNewItem.Height = oForm.Items.Item("137").Height
            oNewItem.Width = oForm.Items.Item("137").Width
            oNewItem.Enabled = False
            oNewItem.DisplayDesc = True
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            CType(oNewItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OOPR", "U_EXO_OPVIN")
            oNewItem = oForm.Items.Add("lblOVIN", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("cbOVIN").Top
            oNewItem.Left = oForm.Items.Item("136").Left
            oNewItem.Height = oForm.Items.Item("136").Height
            oNewItem.Width = oForm.Items.Item("136").Width
            oNewItem.FromPane = 0 : oNewItem.ToPane = 0
            oNewItem.LinkTo = "cbOVIN"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Oportunidad Vinculada"
            oNewItem = oForm.Items.Add("lnkOV", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oNewItem.Top = oForm.Items.Item("cbOVIN").Top
            oNewItem.Left = oForm.Items.Item("cbOVIN").Left - 20
            oNewItem.LinkTo = "cbOVIN"
            CType(oNewItem.Specific, SAPbouiCOM.LinkedButton).LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_SalesOpportunity


#End Region

            oNewItem = oForm.Items.Add("lblF", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("24").Top + oForm.Items.Item("24").Height + 5
            oNewItem.Left = oForm.Items.Item("18").Left
            oNewItem.Height = oForm.Items.Item("18").Height
            oNewItem.Width = oForm.Items.Item("19").Width
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            oNewItem.LinkTo = "24"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Fechas:"

#Region "Presentación oferta"
            oNewItem = oForm.Items.Add("txtFPO", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.LinkTo = "lblF"
            oNewItem.Top = oForm.Items.Item("lblF").Top + oForm.Items.Item("lblF").Height + 2
            oNewItem.Left = oForm.Items.Item("24").Left
            oNewItem.Height = oForm.Items.Item("24").Height
            oNewItem.Width = oForm.Items.Item("24").Width
            oNewItem.Enabled = True
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            CType(oNewItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OOPR", "U_EXO_FPO")
            oNewItem = oForm.Items.Add("lblFPO", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("txtFPO").Top
            oNewItem.Left = oForm.Items.Item("18").Left
            oNewItem.Height = oForm.Items.Item("18").Height
            oNewItem.Width = oForm.Items.Item("30").Width
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            oNewItem.LinkTo = "txtFPO"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Presentación Oferta"
#End Region
#Region "Adjudicación o aceptación oferta"
            oNewItem = oForm.Items.Add("txtFAO", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.LinkTo = "lblF"
            oNewItem.Top = oForm.Items.Item("txtFPO").Top + oForm.Items.Item("txtFPO").Height + 2
            oNewItem.Left = oForm.Items.Item("24").Left
            oNewItem.Height = oForm.Items.Item("24").Height
            oNewItem.Width = oForm.Items.Item("24").Width
            oNewItem.Enabled = True
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            CType(oNewItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OOPR", "U_EXO_FAO")
            oNewItem = oForm.Items.Add("lblFAO", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("txtFAO").Top
            oNewItem.Left = oForm.Items.Item("18").Left
            oNewItem.Height = oForm.Items.Item("18").Height
            oNewItem.Width = oForm.Items.Item("30").Width
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            oNewItem.LinkTo = "txtFAO"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Adj. o aceptación oferta"
#End Region
#Region "Firma de contrato"
            oNewItem = oForm.Items.Add("txtFFC", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.LinkTo = "lblF"
            oNewItem.Top = oForm.Items.Item("txtFAO").Top + oForm.Items.Item("txtFAO").Height + 2
            oNewItem.Left = oForm.Items.Item("24").Left
            oNewItem.Height = oForm.Items.Item("24").Height
            oNewItem.Width = oForm.Items.Item("24").Width
            oNewItem.Enabled = True
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            CType(oNewItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OOPR", "U_EXO_FFC")
            oNewItem = oForm.Items.Add("lblFFC", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("txtFFC").Top
            oNewItem.Left = oForm.Items.Item("18").Left
            oNewItem.Height = oForm.Items.Item("18").Height
            oNewItem.Width = oForm.Items.Item("30").Width
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            oNewItem.LinkTo = "txtFFC"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Firma de contrato"
#End Region
#Region "Finalización del proyecto"
            oNewItem = oForm.Items.Add("txtFFP", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oNewItem.LinkTo = "lblF"
            oNewItem.Top = oForm.Items.Item("txtFFC").Top + oForm.Items.Item("txtFFC").Height + 2
            oNewItem.Left = oForm.Items.Item("24").Left
            oNewItem.Height = oForm.Items.Item("24").Height
            oNewItem.Width = oForm.Items.Item("24").Width
            oNewItem.Enabled = True
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            CType(oNewItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OOPR", "U_EXO_FFP")
            oNewItem = oForm.Items.Add("lblFFP", BoFormItemTypes.it_STATIC)
            oNewItem.Top = oForm.Items.Item("txtFFP").Top
            oNewItem.Left = oForm.Items.Item("18").Left
            oNewItem.Height = oForm.Items.Item("18").Height
            oNewItem.Width = oForm.Items.Item("30").Width
            oNewItem.FromPane = 1 : oNewItem.ToPane = 1
            oNewItem.LinkTo = "txtFFP"
            CType(oNewItem.Specific, SAPbouiCOM.StaticText).Caption = "Finalización del proyecto"
#End Region
#Region "Botón de Control de proyecto"
            oNewItem = oForm.Items.Add("btnCP", BoFormItemTypes.it_BUTTON)
            oNewItem.Top = oForm.Items.Item("2").Top
            oNewItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oNewItem.Height = oForm.Items.Item("2").Height
            oNewItem.Width = (2 * oForm.Items.Item("2").Width)
            oNewItem.LinkTo = "2"
            CType(oNewItem.Specific, SAPbouiCOM.Button).Caption = "Control de proyectos"
            oForm.Items.Item("btnCP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnCP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnCP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
#Region "Botón Enlace a Anexos"
            oNewItem = oForm.Items.Add("btnEA", BoFormItemTypes.it_BUTTON)
            oNewItem.Top = oForm.Items.Item("2").Top
            oNewItem.Left = oForm.Items.Item("btnCP").Left + oForm.Items.Item("btnCP").Width + 5
            oNewItem.Height = oForm.Items.Item("2").Height
            oNewItem.Width = (2 * oForm.Items.Item("2").Width)
            oNewItem.LinkTo = "2"
            CType(oNewItem.Specific, SAPbouiCOM.Button).Caption = "Enlace a Anexos"
            oForm.Items.Item("btnEA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnEA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnEA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
#Region "Botón Documentos a Entregar"
            oNewItem = oForm.Items.Add("btnDE", BoFormItemTypes.it_BUTTON)
            oNewItem.Top = oForm.Items.Item("2").Top
            oNewItem.Left = oForm.Items.Item("btnEA").Left + oForm.Items.Item("btnEA").Width + 5
            oNewItem.Height = oForm.Items.Item("2").Height
            oNewItem.Width = (2 * oForm.Items.Item("2").Width)
            oNewItem.LinkTo = "2"
            CType(oNewItem.Specific, SAPbouiCOM.Button).Caption = "Doc. a Entregar"
            oForm.Items.Item("btnDE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnDE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oForm.Items.Item("btnDE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region
#End Region

#Region "Inicializamos valores"
            'Inicializamos valores
            Try
                CargarCombos(oForm)
                'CType(oForm.Items.Item("44").Specific, SAPbouiCOM.ComboBox).Select("2", BoSearchKey.psk_ByValue)
            Catch ex As Exception
                objGlobal.SBOApp.StatusBar.SetText("No se puede inicializar el campo ""Oportunidades"". " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
#End Region
        Catch exCOM As System.Runtime.InteropServices.COMException
                Throw exCOM
            Catch ex As Exception
                Throw ex
            Finally
                ' oForm.Freeze(False)
                If oForm IsNot Nothing Then oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub CargarCombos(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""

        Try
            sSQL = "SELECT ""OpprId"",""Name"" FROM ""OOPR"" WHERE ""U_EXO_CLASIFICACION""<>'NR' "

            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbOVIN").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "320"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "320"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    If CType(oForm.Items.Item("chkMarco").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                                        oForm.Items.Item("cbOVIN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                    Else
                                        oForm.Items.Item("cbOVIN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
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
#End Region
End Class
