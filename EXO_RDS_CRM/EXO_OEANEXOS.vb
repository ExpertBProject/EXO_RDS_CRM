Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OEANEXOS
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OEANEXOS.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_OEANEXOS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                    Case "EXO-MnEAO"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_OEANEXOS")
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
                        Case "UDO_FT_EXO_OEANEXOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                    If EventHandler_DOUBLE_CLICK_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OEANEXOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
                        Case "UDO_FT_EXO_OEANEXOS"
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
                        Case "UDO_FT_EXO_OEANEXOS"
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
    Private Function EventHandler_DOUBLE_CLICK_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCarpeta As String
        EventHandler_DOUBLE_CLICK_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "13_U_E" 'Guardamos el nombre de la carpeta
                    sCarpeta = CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value
                    If sCarpeta <> "" Then
                        Process.Start(sCarpeta)
                    End If
            End Select

            EventHandler_DOUBLE_CLICK_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_VALIDATE_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sCarpeta As String = ""
        EventHandler_VALIDATE_Before = False
        Try
            If pVal.ItemUID = "13_U_E" Then
                sCarpeta = CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value.ToString
                If sCarpeta <> "" Then
                    If System.IO.Directory.Exists(sCarpeta) = True Then
                        EventHandler_VALIDATE_Before = True
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - La carpeta seleccionada no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        EventHandler_VALIDATE_Before = False
                    End If
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
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False
        Dim sDes As String = ""
        Try

            If pVal.ItemUID = "0_U_E" Then
                Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "97"
                                Try
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@EXO_OPORANEXOS").SetValue("Name", 0, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@EXO_OPORANEXOS").SetValue("Name", 0, sDes)
                                End Try
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCarpeta(0) As String : Dim sFichero As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnCar" 'Guardamos el nombre de la carpeta
                    If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                        sFichero = objGlobal.SBOApp.GetFileFromBrowser()
                        sCarpeta = IO.Directory.GetDirectories(sFichero)
                    Else
                        sFichero = objGlobal.funciones.OpenDialogFolder("Búsqueda de Carpeta")
                        sCarpeta(0) = sFichero
                    End If
                    CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).Value = sCarpeta(0)
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
            If oForm.Visible = True Then
                'objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oForm.Freeze(True)

#Region "Inicializamos valores"
                'Inicializamos valores
                Try
                    If oForm.Mode = BoFormMode.fm_ADD_MODE And _sOportunidad <> "" Then
                        ' Si no existe el control de proyecto
                        CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value = _sOportunidad
                    End If

                    CargarCombos(oForm)
                    'Se quita la columna de etapas
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Visible = False

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
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").DisplayDesc = True
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ExpandType = BoExpandType.et_DescriptionOnly
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
