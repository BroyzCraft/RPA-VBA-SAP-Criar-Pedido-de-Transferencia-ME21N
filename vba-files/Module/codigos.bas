Attribute VB_Name = "codigos"
Sub CriaPedidoTransferencia()

    'Desativa atualização de tela e alertas
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    'Salva os dados atuais
    ActiveWorkbook.Save
   
    'Dim sapAberto As Variant
    sapAberto = Shell("taskkill /IM saplogon.exe", vbNormalFocus)
    sapAberto = Shell("taskkill /IM saplogon.exe", vbNormalFocus)
    
    Dim usuario, senha
    usuario = UCase(InputBox("Digite seu login SAP: "))
    senha = InputBox("Digite sua senha SAP: ")
    'debug
    'usuario = "BOMARQUES"
    'senha = "321654987Leo!"
    
    'Valida se usuario preencheu os dados de login
    If usuario = "" Or senha = "" Then
        MsgBox ("Dados de Login Incompletos!")
        End
    End If
    
    Dim SapGui, Applic, connection, session, WSHShell
    
    'Abre o Sap instalado na sua máquina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    'Inicia a variável com o objeto SAP
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    
    Set WSHShell = Nothing
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.OpenConnection("14 - ECC PRD - EP1", True)
    Set session = connection.Children(0)
    
    session.findById("wnd[0]").maximize
    'DADOS PARA FAZER O LOGIN NO SISTEMA

    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "500" 'client do sistema
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario 'usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha 'senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"  'idioma do sistema
    session.findById("wnd[0]").sendVKey 0 'botão enter para entrar no sistema
    session.findById("wnd[0]").maximize
    
    'parametriza a tela de pedidos como transferencia
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ME21N"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "ZUB"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "1000"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").SetFocus
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").caretPosition = 4
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").Text = "1000"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").caretPosition = 4
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").Text = "999"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    
    'insere itens
    linha = Range("A100000").End(xlUp).Row + 1
    x = 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").Text = "1"
    session.findById("wnd[0]").sendVKey 0
    linhaSap = 0
    Do While x < linha - 2
        'item
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & linhaSap & "]").SetFocus
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & linhaSap & "]").caretPosition = 7
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & linhaSap & "]").Text = Range("A" & x + 2).Value
        session.findById("wnd[0]").sendVKey 0
        'destino
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11," & linhaSap & "]").Text = Range("C" & x + 2).Value
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11," & linhaSap & "]").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
        'qtde
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & linhaSap & "]").Text = Range("B" & x + 2).Value
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & linhaSap & "]").Text = Range("B" & x + 2).Value
        session.findById("wnd[0]").sendVKey 0
        If x = 0 Then
            linhaSap = 1
        Else
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = x
        End If
        x = x + 1
    Loop
    'salvar
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    transf = InputBox("Digite o Numero da Transferencia: ")
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "VL10B"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").Text = transf
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").SetFocus
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").caretPosition = 6
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'gerar remessa
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AMPEL"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WADAT"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LPRIO"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KUNWE"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ROUTE"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VBELV"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BRGEW"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "GEWEI"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLUM"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLEH"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VSTEL"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AUART"
    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VKBUR"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/tbar[1]/btn[20]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AMPEL"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WADAT"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LPRIO"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KUNWE"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ROUTE"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VBELV"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VBELN"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BRGEW"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "GEWEI"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLUM"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VOLEH"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VKORG"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "VSTEL"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AUART"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 1, "VBELN"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
    session.findById("wnd[0]/tbar[1]/btn[25]").press
    
    'valida qdte itens
    x = 0
    Do While x < linha - 2
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").Text = Range("B" & x + 2).Value
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").SetFocus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").caretPosition = 3
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").Text = Range("B" & x + 2).Value
        session.findById("wnd[0]").sendVKey 0
    Loop
    
    session.findById("wnd[0]/tbar[1]/btn[20]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "VBELV"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14").Select
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").currentCellColumn = "BELNR"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/tbar[1]/btn[24]").press
    
    'Imprimir NF
    'NF = "003818568" 'InputBox("Digite o numero de 9º da NF: ")
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "J1B3N"
    session.findById("wnd[0]").sendVKey 0
    '------------------------CORTAR PROCESSO E IMPRIMIR DIRETO
    'session.findById("wnd[0]").sendVKey 4
    'session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").Text = NF
    'session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").SetFocus
    'session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").caretPosition = 9
    'session.findById("wnd[1]/tbar[0]/btn[0]").press
    'session.findById("wnd[1]/tbar[0]/btn[0]").press
    '------------------------CORTAR PROCESSO E IMPRIMIR DIRETO
    session.findById("wnd[0]/mbar/menu[0]/menu[8]").Select
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    'Set session = Nothing
    'app.Wait Now + TimeValue("0:00:05")
    'connection.CloseSession ("ses[0]")
    'Set connection = Nothing
    'Set sap = Nothing
    
    MsgBox ("Transferencia realizada e NF Impressa!")
    
End Sub
