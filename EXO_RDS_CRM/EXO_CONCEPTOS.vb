Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_CONCEPTOS
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CONCEPTOS.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_CONCEPTOS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
            'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Carga_Valores()
        End If
    End Sub
    Private Sub Carga_Valores()
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos    

        Try
            oDI_COM = objGlobal.refDi.dameEXO_UDOEntity("EXO_CONCEPTOS") 'New EXO_DIAPI.EXO_UDOEntity("EXO_XRTP") 'UDO de Campos de SAP
            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Aduanas"
            oDI_COM.SetValue("CodEntry") = "01"
            oDI_COM.SetValue("Name") = "Aduanas"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Aduanas"". " & oDI_COM.GetLastError)
            End If

            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Coste Financiero"
            oDI_COM.SetValue("CodEntry") = "02"
            oDI_COM.SetValue("Name") = "Coste Financiero"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Coste Financiero"". " & oDI_COM.GetLastError)
            End If

            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Mano de Obra"
            oDI_COM.SetValue("CodEntry") = "03"
            oDI_COM.SetValue("Name") = "Mano de Obra"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Mano de Obra"". " & oDI_COM.GetLastError)
            End If

            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Costes BOC"
            oDI_COM.SetValue("CodEntry") = "04"
            oDI_COM.SetValue("Name") = "Costes BOC"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Costes BOC"". " & oDI_COM.GetLastError)
            End If

            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Gastos de viajes"
            oDI_COM.SetValue("CodEntry") = "05"
            oDI_COM.SetValue("Name") = "Gastos de viajes"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Gastos de viajes"". " & oDI_COM.GetLastError)
            End If

            oDI_COM.GetNew()
            oDI_COM.SetValue("Code") = "Aval"
            oDI_COM.SetValue("CodEntry") = "06"
            oDI_COM.SetValue("Name") = "Aval"
            If oDI_COM.UDO_Add = False Then
                Throw New Exception("(EXO) - Error al añadir Concepto ""Aval"". " & oDI_COM.GetLastError)
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))

        End Try
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
                    Case "EXO-MnCNC"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CONCEPTOS")
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
                        Case "UDO_FT_EXO_CONCEPTOS"
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
                        Case "UDO_FT_EXO_CONCEPTOS"
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
                        Case "UDO_FT_EXO_CONCEPTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CONCEPTOS"
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


#End Region
End Class
