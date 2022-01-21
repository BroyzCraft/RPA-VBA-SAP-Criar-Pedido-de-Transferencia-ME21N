Attribute VB_Name = "SAPLogon"
Sub SAP_Logon()
    
    Dim usuario, senha
    usuario = UCase(InputBox("Digite seu login SAP: "))
    senha = InputBox("Digite sua senha SAP: ")
    
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

End Sub
