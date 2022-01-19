Attribute VB_Name = "codigos"
Sub CriaPedidoTransferencia(gerar As Integer)

'Desativa atualização de tela e alertas
Application.ScreenUpdating = True
Application.DisplayAlerts = True

'Salva os dados atuais
ActiveWorkbook.Save
   
login = 0
If login = 1 Then
sap_login:
    Call SAP_Logon
End If

If Not IsObject(app) Then
    On Error GoTo sap_login
    Set SapGuiAuto = GetObject("SAPGUI")
    Set app = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
    Set connection = app.Children(0)
End If

If Not IsObject(session) Then
    Set session = connection.Children(0)
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject app, "on"
End If

If gerar = 1 Then
    GoTo gerar_faturar
End If
    
'parametriza a tela de pedidos como transferencia
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nME21N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").Key = "ZUB"
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
    'session.findById("wnd[0]").sendVKey 0
    'destino
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11," & linhaSap & "]").Text = Range("C" & x + 2).Value
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11," & linhaSap & "]").caretPosition = 4
    'session.findById("wnd[0]").sendVKey 0
    'qtde
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & linhaSap & "]").Text = Range("B" & x + 2).Value
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & linhaSap & "]").Text = Range("B" & x + 2).Value
    session.findById("wnd[0]").sendVKey 0
    'identifica se precisa rolar tela para baixo
    linhaBranco = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & linhaSap & "]").Text
    'pos = x
    If x = 0 Then
        linhaSap = 1
        GoTo pula
    Else
        Do While linhaBranco <> ""
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = x
            linhaBranco = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & linhaSap & "]").Text
            'pos = pos + 1
        Loop
    End If
pula:
    x = x + 1
Loop
'salvar
session.findById("wnd[0]/tbar[0]/btn[11]").press

opcao = MsgBox("Caso tenha apresentado algum erro referente ao material, a macro não irá funcionar a partir deste ponto. Por tanto, faça os ajustes necessarios e clique em 'Sim' se deseja continuar ou 'Não' se deseja cancelar.", vbYesNo)
If opcao = vbNo Then
    End
End If

'colocar opção se deseja continuar
gerar_faturar:
transf = InputBox("Digite o Nº da Transferencia: ")
If transf = "" Then
    End
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVL10B"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5").Select
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").Text = transf
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").SetFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/ctxtST_EBELN-LOW").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press

'gerar remessa
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[20]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 1, "VBELN"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
session.findById("wnd[0]/tbar[1]/btn[25]").press

'valida qdte itens
x = 0
linha = Range("A100000").End(xlUp).Row + 1
Do While x < linha - 2
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").Text = Range("B" & x + 2).Value
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").SetFocus
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & x & "]").Text = Range("B" & x + 2).Value
    session.findById("wnd[0]").sendVKey 0
    x = x + 1
Loop

session.findById("wnd[0]/tbar[1]/btn[20]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/tbar[0]/btn[3]").press

'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "VBELV"
'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").clickCurrentCell
'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14").Select
'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").currentCellColumn = "BELNR"
'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").selectedRows = "0"
'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1316/ssubPO_HISTORY:SAPLMMHIPO:0100/cntlMEALV_GRID_CONTROL_MMHIPO/shellcont/shell").clickCurrentCell
'session.findById("wnd[0]/tbar[1]/btn[24]").press

'Imprimir NF
session.findById("wnd[0]/tbar[0]/okcd").Text = "/nJ1B3N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[0]/menu[8]").Select
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[3]").press

MsgBox ("Transferencia realizada e NF Impressa!")
    
End Sub

Sub GerarRemessaFaturar()
    CriaPedidoTransferencia (1)
End Sub
Sub Criar()
    CriaPedidoTransferencia (0)
End Sub
