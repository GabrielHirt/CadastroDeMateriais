Attribute VB_Name = "CadastroDeMateriais"
Option Explicit

Sub Cadastrar_material()
    
    Dim Application, SapGuiAuto, connection, session, WScript
    Dim TargetRange                                             As Range
    Dim CountCopyCells                                          As Integer
    
    If Not IsObject(Application) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = Application.Children(0)
    End If
    If Not IsObject(session) Then
    Set session = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If

    Range("C1").Select

    While ActiveCell <> ""

        ActiveCell.Offset(1, 0).Select
        
        If ActiveCell <> "" Then
        
            Set TargetRange = ActiveCell
            
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").Text = "MM01"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").Key = "M"
            
'
            If TargetRange.Offset(0, -1).Value = "TIPO1" Then
                session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key = "ROH"
                
                    '\/abrir caixa de selecao de visoes\/
                session.findById("wnd[0]").sendVKey 0
                
                    '\/Dados básicos 1 (0)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True
                    
                    '\/'Dados básicos 2 (1)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).Selected = True
                    
                    '\/Compras (9)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).Selected = True
                    
                    '\/Comércio exterior: importação (10)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).Selected = True
                
                    '\/rolar para selecionar os campos que nao aparecem na primeira tela\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.Position = 10
                
                    '\/MRP 1 (12)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True
                    
                    '\/MRP 2 (13)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True
                    
                    '\/MRP 3 (14)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).Selected = True
                    
                    '\/MRP 4 (15)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).Selected = True
                    
                    '\/Dds.gerais centro/armazen.1 (19)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).Selected = True
                    
                    '\/Dds.gerais centro/armazmto.2 (20)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(20).Selected = True
                    
                    '\/Administração de depósitos 1 (21)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).Selected = True
                    
                    '\/Administração de depósitos 2 (22)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).Selected = True
                    
                    '\/Administração de qualidade (23)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(23).Selected = True
                    
                    '\/Contabilidade 1 (24)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(24).Selected = True
                    
                    '\/Contabilidade 2 (25)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(25).Selected = True
                
                    '\/Calculo de preço 1 (26)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(26).Selected = True
                
                    '\/Calculo de preço 2 (27)\/
                session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(27).Selected = True
                    
                    '\/Abrir caixa de definicao de niveis organizacionais\/
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                    
                    '\/definir centro 1001\/
                session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "xxx"
                    
                    '\/definir deposito 0010 - almoxarifado
                session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = "xxx"
                
                    '\/Abrir visoes para preenchimento do cadastro\/
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                
                    '\/exportar codigo SAP do cadastro para o excel\/
                TargetRange.Offset(0, -2).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB1:SAPLMGD1:1002/ctxtRMMG1-MATNR").Text
                
                    'Inserir nome do cadastrante
                TargetRange.Offset(0, 34).Value = Environ("username")
                    
                    'Inserir data de cadastro
                TargetRange.Offset(0, 35).Value = Date & " " & Time
                TargetRange.Offset(0, 35).Value = Format(TargetRange.Offset(0, 35).Value, "dd/mm/yyyy hh:mm:ss")
                             
                    '\/Inserir descricao breve \/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").Text = TargetRange.Value
                    
                    '\/Inserir unidade\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MEINS").Text = TargetRange.Offset(0, 1).Value
                    
                    '\/Verificar quantidade de caracateres para acrescer a quantidade de zeros correta à esquerda.\/
                If TargetRange.Offset(0, 2) = "xx" Or TargetRange.Offset(0, 2) = "xx" Then
                        '\/Inserir grupo de mercadorias\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MATKL").Text = TargetRange.Offset(0, 2).Value
                
                ElseIf Len(TargetRange.Offset(0, 2)) = 1 Then
                        '\/Inserir grupo de mercadorias\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MATKL").Text = "00" & TargetRange.Offset(0, 2).Value
                        
                ElseIf Len(TargetRange.Offset(0, 2)) = 2 Then
                        '\/Inserir grupo de mercadorias\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MATKL").Text = "0" & TargetRange.Offset(0, 2).Value
                
                ElseIf Len(TargetRange.Offset(0, 2)) = 3 Then
                        '\/Inserir grupo de mercadorias\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2001/ctxtMARA-MATKL").Text = TargetRange.Offset(0, 2).Value
                End If
                    
                    '\/Verificar se é importado\/
                If TargetRange.Offset(0, 10).Value = "X" Or TargetRange.Offset(0, 10).Value = "XX" Then
                        
                        '\/Abrir tela de descrições internacionais\/
                    session.findById("wnd[0]").sendVKey 5
                        
                        '\/Setar idioma EN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT/ctxtSKTEXT-SPRAS[0,1]").Text = "EN"
                        
                        '\/Adicionar descricao internacional EN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,1]").Text = TargetRange.Offset(0, 4).Value
                    
                        '\/Abrir tela de descricao de importacao\/
                    session.findById("wnd[0]").sendVKey 9
                    
                        '\/Adicionar descricao de importacao\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell").Text = TargetRange.Offset(0, 6).Value
                    
                Else
                    
                    If TargetRange.Offset(0, 5).Value <> "N/A" Then
                    
                            '\/Abrir tela de descricao complementar\/
                        session.findById("wnd[0]").sendVKey 9
                        
                            '\/Adicionar descricao complementar\/
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU05/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:2031/cntlLONGTEXT_GRUNDD/shellcont/shell").Text = TargetRange.Offset(0, 5).Value

                    End If
                End If
                
                If TargetRange.Offset(0, 1).Value = "XX" Then
                    
                        '\/Abrir tela de conversão de unidades\/
                    session.findById("wnd[0]").sendVKey 6
                        
                        '\/Adicionar quantidade 1 para UN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREN[0,1]").Text = "1"
                    
                        '\/Adicionar unidade UN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/ctxtSMEINH-MEINH[1,1]").Text = "UN"
                    
                        '\/Adicionar quantidade 1 para XX\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREZ[4,1]").Text = "1"
                        
                        '\/Adicionar quantidade 1 para PEÇ\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREN[0,2]").Text = "1"
                    
                        '\/Adicionar unidade PEÇ\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/ctxtSMEINH-MEINH[1,2]").Text = "PEÇ"
                    
                        '\/Adicionar quantidade 1 para XX\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREZ[4,2]").Text = "1"
                
                ElseIf TargetRange.Offset(0, 1).Value = "UN" Then
                    
                        '\/Abrir tela de conversão de unidades\/
                    session.findById("wnd[0]").sendVKey 6
                    
                        '\/Adicionar quantidade 1 para PEÇ\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREN[0,10]").Text = "1"
                    
                        '\/Adicionar unidade PEÇ\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/ctxtSMEINH-MEINH[1,10]").Text = "PEÇ"
                    
                        '\/Adicionar quantidade 1 para UN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREZ[4,10]").Text = "1"
            
                ElseIf TargetRange.Offset(0, 1).Value = "PEÇ" Then
                        
                        '\/Abrir tela de conversão de unidades\/
                    session.findById("wnd[0]").sendVKey 6
                    
                        '\/Adicionar quantidade 1 para UN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREN[0,10]").Text = "1"
                    
                        '\/Adicionar unidade UN\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/ctxtSMEINH-MEINH[1,10]").Text = "UN"
                    
                        '\/Adicionar quantidade 1 para PEÇ\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpZU02/ssubTABFRA1:SAPLMGMM:2110/subSUB2:SAPLMGD1:8020/tblSAPLMGD1TC_ME_8020/txtSMEINH-UMREZ[4,10]").Text = "1"
                    
                End If
                
                        '\/Voltar para dados básicos (F3) - voltar\/
                    session.findById("wnd[0]").sendVKey 3
                
                    '\/Selecionar visao dados basicos 2\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP02").Select
                    
                    '\/Selecionar visao compras\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10").Select
                    
                    If TargetRange.Offset(0, 1).Value = "XX" Then
                    
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARA-BSTME").Text = "UN"
                        
                    End If
                    
                        '\/Adicionar grupo de compradores\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = TargetRange.Offset(0, 11).Value
                    
                    If TargetRange.Offset(0, 13).Value <> "N/A" Then
                            '\/Adicionar numero fabricante\/
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2313/ctxtMARA-MFRNR").Text = TargetRange.Offset(0, 13).Value
                    End If
                    
                    If TargetRange.Offset(0, 14).Value <> "N/A" Then
                            '\/Adicionar PN\/
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2313/txtMARA-MFRPN").Text = TargetRange.Offset(0, 14).Value
                    End If
                    
                    '\/selecionar visao Comercio exterior: importacao\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11").Select
                    
                        '\/adicionar NCM\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2205/ctxtMARC-STEUC").Text = TargetRange.Offset(0, 15).Value
                    
                        '\/adicionar categoria CFOP\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2203/ctxtMARC-INDUS").Text = TargetRange.Offset(0, 16).Value
                
                    '\/verificar se é importado XX
                If TargetRange.Offset(0, 10).Value = "IMPORTADO (XX)" Then
                        
                        '\/selecionar visao texto pedido compras\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").Select
                    
                        '\/Adicionar hana code\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = TargetRange.Offset(0, 17).Value
                
                End If
                    
                    '\/selecionar visao MRP 1\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").Select
                
                        '\/adicionar grupo MRP\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-DISGR").Text = "0040"
                        
                        '\/adicionar tipo de MRP\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "PD"
                    
                        '\/adicionar Planejador MRP\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = "120"
                    
                        '\/adicionar RegraCalcTamLotes\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = "WB"
                
                    '\/selecionar visao MRP 2\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").Select
                    
                        '\/adicionar Deposito de producao\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").Text = "0040"
                    
                        '\/adicionar Baixa por explosao\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-RGEKZ").Text = "1"
                    
                        '\/adicionar Depos.suprimto.ext.\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").Text = "0010"
                    
                        '\/adicionar Prazo entrega prevista\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").Text = TargetRange.Offset(0, 23).Value
                    
                        '\/adicionar Tempo processamento EM\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").Text = TargetRange.Offset(0, 24).Value
                
                    '\/selecionar visao MRP 3\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15").Select
                session.findById("wnd[0]").sendVKey 0
                
                        '\/adicionar verif. disponib.\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = "SR"
                
                    '\/selecionar visao MRP 4\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16").Select
                        
                    '\/selecionar visao dados gerais centro/armazenamento 1\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21").Select
                
                    '\/selecionar visao dados gerais centro/armazenamento 2\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22").Select
                    
                    '\/selecionar visao administracao de depositos 1\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").Select
                    
                    '\/selecionar visao administracao de depositos 2\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP25").Select
                    
                    '\/selecionar visao administracao de qualidade\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").Select
                    
                        'verificar HSMS
                    If TargetRange.Offset(0, 18).Value = "SIM" Then
                        
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").Selected = True
                        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").Text = "XXxx"
                    
                    End If
                    
                    '\/selecionar visao Contabilidade 1\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27").Select
                      
                        '\/adicionar classe de avaliacao\/
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").Text = TargetRange.Offset(0, 19).Value
                    
                    '\/Verificar se preço é zero e adicionar preco caso não seja\/
                If TargetRange.Offset(0, 20).Value = 0 Or TargetRange.Offset(0, 20).Value = "" Then
                    
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    
                    
                ElseIf TargetRange.Offset(0, 20).Value <> 0 Or TargetRange.Offset(0, 20).Value <> "" Then
                    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP27/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/subSUBCURR:SAPLCKMMAT:0200/txtCKMMAT_DISPLAY-STPRS_1").Text = TargetRange.Offset(0, 20).Value
                    
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    
                    
                    'If TargetRange.Offset(0, -1).Value <> "TIPO01" Then
                    session.findById("wnd[0]").sendVKey 0
                    'End If
                    
                End If
                
                    '\/passar para a visão contabilidade 2\/
                session.findById("wnd[0]").sendVKey 0
                    
                    '\/adicionar utilizacao do material\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP28/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2806/ctxtMBEW-MTUSE").Text = TargetRange.Offset(0, 21).Value
                    
                    '\/adicionar origem do material\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP28/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2806/ctxtMBEW-MTORG").Text = TargetRange.Offset(0, 22).Value
                
                    '\/selecionar visao calculo de preço 1\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP29").Select
                
                    '\/selecionar visao Calculo de preço 2\/
                session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP30").Select
               
                    '\/finalizar e salvar cadastro\/
                session.findById("wnd[0]").sendVKey 11
                
                    '\/confirmar fechamento
                session.findById("wnd[0]").sendVKey 0
                
                session.findById("wnd[0]").sendVKey 3
                
                    '\/ir para criacao coletiva de depositos\/
                session.findById("wnd[0]/tbar[0]/okcd").Text = "MMSC"
                
                    '\/confirmar
                session.findById("wnd[0]").sendVKey 0
                
                    '\/adicionar codigo SAP\/
                session.findById("wnd[0]/usr/ctxtRM03M-MATNR").Text = TargetRange.Offset(0, -2).Value
                
                    '\/adicionar centro\/
                session.findById("wnd[0]/usr/ctxtRM03M-WERKS").Text = "1001"
                
                    '\/abrir tela de criacao de depositos\/
                session.findById("wnd[0]").sendVKey 0
                
                    '\/adicionar depositos\/
                session.findById("wnd[0]/usr/sub:SAPMM03M:0195/ctxtRM03M-LGORT[1,0]").Text = "0001"
                session.findById("wnd[0]/usr/sub:SAPMM03M:0195/ctxtRM03M-LGORT[2,0]").Text = "0002"
                session.findById("wnd[0]/usr/sub:SAPMM03M:0195/ctxtRM03M-LGORT[3,0]").Text = "0003"
                session.findById("wnd[0]/usr/sub:SAPMM03M:0195/ctxtRM03M-LGORT[4,0]").Text = "0004"
                    
                    '\/finalizar criacao de depositos\/
                session.findById("wnd[0]").sendVKey 11
                
                    '\/voltar para tela inicial\/
                session.findById("wnd[0]").sendVKey 3
                
                TargetRange.Select
                
'
            ElseIf TargetRange.Offset(0, -1).Value = "TIPO02" Then

'              [ PROCESSO SIMILAR É REALIZADO, PODENDO EXISTIR NOVAS VISÕES E CAMPOS DE ACORDO COM O TIPO DO ITEM E REGRA DE NEGÓCIO]

            ElseIf TargetRange.Offset(0, -1).Value = "TIPO03" Then

'              [ PROCESSO SIMILAR É REALIZADO, PODENDO EXISTIR NOVAS VISÕES E CAMPOS DE ACORDO COM O TIPO DO ITEM E REGRA DE NEGÓCIO]

            ElseIf TargetRange.Offset(0, -1).Value = "TIPO04" Then

'              [ PROCESSO SIMILAR É REALIZADO, PODENDO EXISTIR NOVAS VISÕES E CAMPOS DE ACORDO COM O TIPO DO ITEM E REGRA DE NEGÓCIO]

            End If
        End If
    Wend
               
    
    
'
' O trechos abaixos são responsáveis por realizar a criação do histórico
    If ActiveCell = "" Then
        
        If TargetRange.Offset(0, -1).Value <> "XXX" And TargetRange.Offset(0, -1).Value <> "XX" Then
            
                'selecionar sheet 'Cadastro sem MRP'
            Sheets("Cadastro sem MRP").Activate
                
                'selecionar primeiro codigo
            Range("A2").Select
                
                'zerar contador de itens cadastrados
            CountCopyCells = 0
            
                'rodar contador de itens cadastrados
            While ActiveCell <> ""
            
                    'rodar contador de itens cadastrados
                CountCopyCells = CountCopyCells + 1
                
                    'descer uma linha para verificar se há próximo item
                ActiveCell.Offset(1, 0).Select
                
            Wend
                
                'copiar o range de valores exportados da sharepoint list
            Range("A2" & ":AK" & CountCopyCells + 1).Copy
                
                'selecionar sheet histórico
            Sheets("Historico").Activate
            
                'desfiltrar se estiver filtrado
            If ActiveSheet.FilterMode = True Then
                ActiveSheet.ShowAllData
                Sheets("Cadastro sem MRP").Activate
                Range("A2" & ":AK" & CountCopyCells + 1).Copy
                Sheets("Historico").Activate
                
            End If
            
                'verificar se há itens na base de histórico
            If Range("B2") = "" Then
                    'selecionar a célula B2 se não houver itens na base de histórico
                Range("B2").Select
            Else
                    'selecionar o cabeçalho da coluna código ("B") para utilização do CTRL + para baixo c/ offset de uma colum
                Range("B1").End(xlDown).Offset(1, 0).Select
            End If
            
                'colar os dados copiados na célula que está selecionada.
                'Valores copiados: sheet 'Cadastro sem MRP', de coluna A até coluna AK
            ActiveSheet.Paste
            
                'selecionar sheet 'Cadastro sem MRP'
            Sheets("Cadastro sem MRP").Activate
            
                'Selecionar primeiro valor da coluna CADASTRADOEM
            Range("AL2").Select
                
                'Verificar se há um ou mais de um item
            If Range("AL3") = "" Then
                    'Se a segunda linha não contém informação, então só há uma informação.
                    'Copiar informação única
                Range("AL2").Copy
            Else
                    'Se na segunda linha há informação, então há mais de uma informação.
                    'Copiar informações
                Range(Selection, Selection.End(xlDown)).Copy
            End If
                
                'Selecionar sheet Histórico para colar informação de CADASTRADOEM
            Sheets("Historico").Activate
            
                'Verificar se há alguma informação na coluna DATA da base (caso não tenha, significa que a base está vazia)
            If Range("A2") = "" Then
                    'se não tem, selecionar primeira linha para colar os dados
                Range("A2").Select
            Else
                    'se tem informações selecionar primeira linha vazia para colar os dados
                Range("A2").End(xlDown).Offset(1, 0).Select
            End If
                
                'Colar os dados de CADASTRADOEM na base de histórico
            ActiveSheet.Paste
            
            'formata a data para o padrão Brasil
            Range(Selection, Selection).NumberFormat = "dd/mm/yyyy hh:mm:ss"
            
                'selecionar sheet 'Cadastro sem MRP'
            Sheets("Cadastro sem MRP").Activate
                
                'Deletar linhas da sheet de cadastro
            Rows("2:" & CountCopyCells + 1).Delete Shift:=xlUp
            
                'Selecionar primeiras celulas para reenquadrar a visao da sheet
            Range("D1").Select
            
                'Selecionar sheet histórico
            Sheets("Historico").Activate
            
                'Selecionar ultima linha da base
            Range("B2").End(xlDown).Offset(-CountCopyCells + 1).Select
            
            
'
        Else
        
        
                
                'selecionar sheet 'Cadastro com MRP'
            Sheets("Cadastro com MRP").Activate
                
                'selecionar primeiro código
            Range("A2").Select
                
                'zerar contador de itens cadastrados
            CountCopyCells = 0
                
                'rodar contador de itens cadastrados
            While ActiveCell <> ""
                    
                    'somar +1 para cada item cadastrado
                CountCopyCells = CountCopyCells + 1
                    
                    'descer uma linha para verificar se há próximo item
                ActiveCell.Offset(1, 0).Select
                
            Wend
            
                'copiar o range de valores exportados da sharepoint list
            Range("A2" & ":AJ" & CountCopyCells + 1).Copy
                
                'selecionar sheet histórico
            Sheets("Historico").Activate
            
                'desfiltrar se estiver filtrado
            If ActiveSheet.FilterMode = True Then
                ActiveSheet.ShowAllData
                Sheets("Cadastro com MRP").Activate
                Range("A2" & ":AJ" & CountCopyCells + 1).Copy
                Sheets("Historico").Activate
            End If
            
                'verificar se há itens na base de histórico
            If Range("B2") = "" Then
                    'selecionar a célula B2 se não houver itens na base de histórico
                Range("B2").Select
            Else
                    'selecionar próxima célula vazia para a colagem dos dados
                Range("B1").End(xlDown).Offset(1, 0).Select
            End If
            
                'colar os dados copiados na célula que está selecionada.
                'Valores copiados: sheet 'Cadastro com MRP', de coluna A até coluna AJ
            ActiveSheet.Paste
                
                'selecionar a sheet 'Cadastro com MRP'
            Sheets("Cadastro com MRP").Activate
                
                'Selecionar primeiro valor da coluna CADASTRADOPOR
            Range("BG2").Select
                
                'Verificar se há um ou mais de um item
            If Range("BG3") = "" Then
                    'Se a segunda linha não contém informação, então só há uma informação.
                    'Copiar informação única
                Range("BG2").Copy
            Else
                    'Se na segunda linha há informação, então há mais de uma informação.
                    'Copiar informações
                Range(Selection, Selection.End(xlDown)).Copy
            End If
                
                'Selecionar sheet Histórico para colar informação de CADASTRADOPOR
            Sheets("Historico").Activate
                
                'Verificar se há alguma informação na coluna CADASTRADOPOR da base (caso não tenha, significa que a base está vazia)
            If Range("AL2") = "" Then
                    'se não tem, selecionar primeira linha para colar os dados
                Range("AL2").Select
            Else
                    'se tem informações selecionar primeira linha vazia para colar os dados
                Range("AL1").End(xlDown).Offset(1, 0).Select
            End If
            
                'Colar os dados de CADASTRADOPOR na base de histórico
            ActiveSheet.Paste
                
                'Selecionar a sheet 'Cadastro com MRP' para copiar a info de CADASTRADOEM
            Sheets("Cadastro com MRP").Activate
            
                'Selecionar a primeira informação da coluna CADASTRADOEM
            Range("BH2").Select
                
                'Verificar se há um ou mais itens
            If Range("BH3") = "" Then
                    'Se há um, copiar o valor único
                Range("BH2").Copy
            Else
                    'Se há mais de um, copiar todos os valores
                Range(Selection, Selection.End(xlDown)).Copy
            End If
                
                'Selecionar a sheet histórico
            Sheets("Historico").Activate
                
                'Verificar se há alguma informação na base
            If Range("A2") = "" Then
                    'Se não há, selecionar a primeira linha após o cabeçalho
                Range("A2").Select
            Else
                    'Se há, selecionar primeira linha vazia
                Range("A2").End(xlDown).Offset(1, 0).Select
            End If
                
                'Colar as informações de CADASTRADOEM na primeira coluna ("A") da base de histórico
            ActiveSheet.Paste
            
                'Formatar tipo de data das datas coladas para corrigir o problema do VBA de colar como mm/dd/yyyy
            Range(Selection, Selection).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                            
                'Selecionar Sheet 'Cadastro com MRP'
            Sheets("Cadastro com MRP").Activate
            
                'copiar o range de valores exportados da sharepoint list
            Range("AK2" & ":BF" & CountCopyCells + 1).Copy
            
                'Selecionar a sheet histórico
            Sheets("Historico").Activate
                
                'Selecionar célula do primeiro item cadastrado nesta operação (número de itens que o CountCopyCells armazenou)
            Range("B2").End(xlDown).Offset(-CountCopyCells + 1, 38).Select
            
                'Colar valores de MRP na base histórico
            ActiveSheet.Paste
            
                'Selecionar Sheet 'Cadastro com MRP'
            Sheets("Cadastro com MRP").Activate
            
                'Deletar linhas da sheet de cadastro
            Rows("2:" & CountCopyCells + 1).Delete Shift:=xlUp
            
                'Selecionar primeiras celulas para reenquadrar a visao da sheet
            Range("D1").Select
                
                'Selecionar sheet histórico
            Sheets("Historico").Activate
                
                'Selecionar ultima linha da base
            Range("B2").End(xlDown).Offset(-CountCopyCells + 1).Select
            
        End If
        
    End If
    
End Sub



