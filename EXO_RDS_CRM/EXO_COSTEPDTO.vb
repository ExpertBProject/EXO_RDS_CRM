Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_COSTEPDTO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
            'ParametrizacionGeneral()
        End If
        'cargamenu()
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_COSTEDPTO.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_COSTEDPTO", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                    Case "EXO-MnCDP"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_COSTEDPTO")
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
                        Case "UDO_FT_EXO_COSTEDPTO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
                        Case "UDO_FT_EXO_COSTEDPTO"
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
                        Case "UDO_FT_EXO_COSTEDPTO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    'If EventHandler_FORM_VISIBLE_After(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_COSTEDPTO"
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
    Private Function EventHandler_VALIDATE_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        Dim iFilaActual As Integer = 0 : Dim iFilaAnterior As Integer = 0
        Dim sFechaINI As String = "" : Dim sFechaFIn As String = "" : Dim dFechaINI As Date = Now.Date : Dim dFechaFIn As Date = Now.Date
        Dim sFechaINIAnt As String = "" : Dim sFechaFInAnt As String = "" : Dim dFechaINIAnt As Date = Now.Date : Dim dFechaFInAnt As Date = Now.Date
        EventHandler_VALIDATE_Before = False
        Try
            iFilaActual = pVal.Row : iFilaAnterior = iFilaActual - 1

            Select Case iFilaActual
                Case 1
                    'Sólo controlamos la columna FFIN
                    Select Case pVal.ColUID
                        Case "C_0_1" : EventHandler_VALIDATE_Before = True
                        Case "C_0_2"
                            'Controlamos que la fecha Fin sea posterior o vacía
                            sFechaINI = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaINI <> "" Then
                                dFechaINI = CDate(Left(sFechaINI, 4) & "-" & Mid(sFechaINI, 5, 2) & "-" & Right(sFechaINI, 2))
                            End If

                            sFechaFIn = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaFIn <> "" Then
                                dFechaFIn = CDate(Left(sFechaFIn, 4) & "-" & Mid(sFechaFIn, 5, 2) & "-" & Right(sFechaFIn, 2))
                                If dFechaFIn < dFechaINI Then
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La fecha Fin no puede ser menor que la fecha de inicio. Compruebe los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objGlobal.SBOApp.MessageBox("La fecha Fin no puede ser menor que la fecha de inicio. Compruebe los datos.")
                                    EventHandler_VALIDATE_Before = False
                                Else
                                    EventHandler_VALIDATE_Before = True
                                End If
                            Else
                                EventHandler_VALIDATE_Before = True
                            End If
                        Case Else
                            EventHandler_VALIDATE_Before = True
                    End Select
                Case Else
                    Select Case pVal.ColUID
                        Case "C_0_1"
                            'Tenemos que mirar que no se solape con la linea anterior
                            sFechaINI = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iFilaActual).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaINI <> "" Then
                                dFechaINI = CDate(Left(sFechaINI, 4) & "-" & Mid(sFechaINI, 5, 2) & "-" & Right(sFechaINI, 2))
                            End If
                            sFechaFIn = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iFilaAnterior).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaFIn <> "" Then
                                dFechaFIn = CDate(Left(sFechaFIn, 4) & "-" & Mid(sFechaFIn, 5, 2) & "-" & Right(sFechaFIn, 2))
                                If dFechaINI <= dFechaFIn Then
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La fecha de inicio se solapa con la fecha fin de la línea anterior. Compruebe los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objGlobal.SBOApp.MessageBox("La fecha de inicio se solapa con la fecha fin de la línea anterior. Compruebe los datos.")
                                    EventHandler_VALIDATE_Before = False
                                Else
                                    EventHandler_VALIDATE_Before = True
                                End If
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - La fecha de inicio se solapa con la fecha fin de la línea anterior. Marque una fecha fin en la línea anterior.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox("La fecha de inicio se solapa con la fecha fin de la línea anterior. Marque una fecha fin en la línea anterior.")
                                EventHandler_VALIDATE_Before = False
                            End If
                        Case "C_0_2"
                            'Controlamos que la fecha Fin sea posterior o vacía
                            sFechaINI = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iFilaActual).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaINI <> "" Then
                                dFechaINI = CDate(Left(sFechaINI, 4) & "-" & Mid(sFechaINI, 5, 2) & "-" & Right(sFechaINI, 2))
                            End If

                            sFechaFIn = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iFilaActual).Specific, SAPbouiCOM.EditText).Value.ToString
                            If sFechaFIn <> "" Then
                                dFechaFIn = CDate(Left(sFechaFIn, 4) & "-" & Mid(sFechaFIn, 5, 2) & "-" & Right(sFechaFIn, 2))
                                If dFechaFIn < dFechaINI Then
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - La fecha Fin no puede ser menor que la fecha de inicio. Compruebe los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objGlobal.SBOApp.MessageBox("La fecha Fin no puede ser menor que la fecha de inicio. Compruebe los datos.")
                                    EventHandler_VALIDATE_Before = False
                                Else
                                    EventHandler_VALIDATE_Before = True
                                End If
                            Else
                                EventHandler_VALIDATE_Before = True
                            End If
                        Case Else
                            EventHandler_VALIDATE_Before = True
                    End Select
            End Select

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing

        EventHandler_FORM_VISIBLE_After = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oForm.Freeze(True)

#Region "Inicializamos valores"
                'Inicializamos valores
                Try
                    CargarCombos(oForm)
                Catch ex As Exception
                    Throw ex
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
            oForm.Items.Item("cbDPTO").DisplayDesc = True
            sSQL = "SELECT 'No Procede' AS ""Departamento"" FROM DUMMY "
            sSQL &= " UNION ALL "
            sSQL &= " SELECT  T0.""Name"" FROM OUDP T0 "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbDPTO").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
