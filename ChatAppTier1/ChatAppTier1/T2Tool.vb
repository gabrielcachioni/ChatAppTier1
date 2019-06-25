Public Class T2Tool
    Dim fimLog As Boolean
    Dim foundLog As Boolean
    Dim count As Integer
    Dim r As Random = New Random
    Dim ranNum As Integer
    Dim se As String
    Dim prodinfo As String
    Dim Outl As Object
    Dim user As String

    Function GetUserName() As String
        If TypeOf My.User.CurrentPrincipal Is
    Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            Dim username As String = parts(1)
            Return username
        Else
            ' The application is using custom authentication.
            Return My.User.Name
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        WebBrowser1.Navigate("https://report.avaya.com/siebelreports/casedetails.aspx?case_id=" + TextBox8.Text)
        While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
            Application.DoEvents()
        End While

        Try
            prod.Text = WebBrowser1.Document.GetElementById("lOpSkill").InnerText
            siteName.Text = WebBrowser1.Document.GetElementById("lSiteName").InnerText
            siteName.Text = WebBrowser1.Document.GetElementById("lSiteName").InnerText
            dateCreated.Text = WebBrowser1.Document.GetElementById("lDateCreated").InnerText
            se = WebBrowser1.Document.GetElementById("lSECode").InnerText
            prodinfo = WebBrowser1.Document.GetElementById("lProduct").InnerText
        Catch ex As Exception
            MsgBox("Insira um numero de chamado valido!", MsgBoxStyle.Exclamation)
            assignTo.Text = "Falha ao indentificar"

        End Try


        If prod.Text = "Communication Manager" Then
            prod.Text = "CM"
        End If
        If prod.Text = "CTI" Then
            prod.Text = "AES"
        End If
        If prod.Text = "System Manager" Then
            prod.Text = "SYSTEM MANAGER"
        End If
        If prod.Text = "Session Manager (ASM)" Then
            prod.Text = "SESSION MANAGER"
        End If
        If prod.Text = "Witness Call Recorder" Then
            prod.Text = "ACR"
        End If
        If prod.Text = "Workforce Management" Or prodinfo = "WFO AFTERMARKET ORDER" Then
            prod.Text = "WFO"
        End If

        If prod.Text = "Oceanalytics" Then
            prod.Text = "OCEANALYTICS"
        End If

        If prod.Text = "Oceana Workspaces" Then
            prod.Text = "OCEANA"
        End If

        If prod.Text = "IP Office" Then
            prod.Text = "IPO"
        End If

        If prod.Text = "IPO Contact Center" Then
            prod.Text = "IPO"
        End If

        If prod.Text = "CMS" Then
            prod.Text = "CMS"
        End If

        If prod.Text = "Voice Portal" Then
            prod.Text = "EPM"
        End If
        If prod.Text = "Proactive Outreach Manager" Then
            prod.Text = "POM"
        End If

        If prod.Text = "Proactive Contact" Then
            prod.Text = "PDS"
        End If

        If prod.Text = "Predictive Dialer" Then
            prod.Text = "PDS"
        End If

        If prod.Text = "AWFOS" Then
            prod.Text = "WFO"
        End If

        ahardmanborgSRs.Text = "-"
        fgsantosSRs.Text = "-"
        fkhaguiwaraSRs.Text = "-"
        gcachioniSRs.Text = "-"
        neves11SRs.Text = "-"
        passoseSRs.Text = "-"

        assignTo.Text = "Carregando!"



        fimLog = False
        foundLog = False
        'logicaSystemManager
        If prod.Text = "SYSTEM MANAGER" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fgsantos")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fgsantosSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fgsantosSRs.Text = fgsantosSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=passose")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            passoseSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            passoseSRs.Text = passoseSRs.Text.Remove(0, 6)

            If fgsantosSRs.Text < fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FGSANTOS - FERNANDO GABRIEL (ROBS)")
            End If

            If fgsantosSRs.Text > fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FKHAGUIWARA - FELIPE HAGUIWARA (HAGUIS)")
            End If

            If fgsantosSRs.Text = fkhaguiwaraSRs.Text Then
                If passoseSRs.Text < fgsantosSRs.Text Then
                    assignTo.Text = ("PASSOSE - EDSON PASSOS")
                Else
                    ranNum = (r.Next(1, 3))
                    If ranNum = 1 Then
                        assignTo.Text = ("RECOMENDADO: FGSANTOS - FERNANDO GABRIEL (ROBS)")
                    End If
                    If ranNum = 2 Then
                        assignTo.Text = ("RECOMENDADO: FKHAGUIWARA - FELIPE HAGUIWARA (HAGUIS)")
                    End If
                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica System Manager
        ' Logica Session Manager
        If prod.Text = "SESSION MANAGER" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fgsantos")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fgsantosSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fgsantosSRs.Text = fgsantosSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=passose")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            passoseSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            passoseSRs.Text = passoseSRs.Text.Remove(0, 6)

            If fgsantosSRs.Text < fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FGSANTOS - FERNANDO GABRIEL (ROBS)")
            End If

            If fgsantosSRs.Text > fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FKHAGUIWARA - FELIPE HAGUIWARA (HAGUIS)")
            End If

            If fgsantosSRs.Text = fkhaguiwaraSRs.Text Then
                If passoseSRs.Text < fgsantosSRs.Text Then
                    assignTo.Text = ("PASSOSE - EDSON PASSOS")
                Else
                    ranNum = (r.Next(1, 3))
                    If ranNum = 1 Then
                        assignTo.Text = ("RECOMENDADO: FGSANTOS - FERNANDO GABRIEL (ROBS)")
                    End If
                    If ranNum = 2 Then
                        assignTo.Text = ("RECOMENDADO: FKHAGUIWARA - FELIPE HAGUIWARA (HAGUIS)")
                    End If

                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica SESSION MANAGER
        ' Logica ACR
        If prod.Text = "ACR" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=neves11")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            neves11SRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            neves11SRs.Text = neves11SRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=gcachioni")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            gcachioniSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            gcachioniSRs.Text = gcachioniSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=ahardmanborg")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            ahardmanborgSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            ahardmanborgSRs.Text = ahardmanborgSRs.Text.Remove(0, 6)

            If gcachioniSRs.Text < neves11SRs.Text Then
                assignTo.Text = ("GCACHIONI - GABRIEL CACHIONI")
            End If

            If gcachioniSRs.Text > neves11SRs.Text Then
                assignTo.Text = ("NEVES11 - CARLOS NEVES (SNOW)")
            End If

            If neves11SRs.Text = gcachioniSRs.Text Then
                If ahardmanborgSRs.Text < gcachioniSRs.Text Then
                    assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
                Else
                    ranNum = (r.Next(1, 3))
                    If ranNum = 1 Then
                        assignTo.Text = ("RECOMENDADO: GCACHIONI - GABRIEL CACHIONI")
                    End If
                    If ranNum = 2 Then
                        assignTo.Text = ("RECOMENDADO: NEVES11 - CARLOS NEVES (SNOW)")
                    End If

                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica ACR
        ' Logica WFO
        If prod.Text = "WFO" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=neves11")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            neves11SRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            neves11SRs.Text = neves11SRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=gcachioni")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            gcachioniSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            gcachioniSRs.Text = gcachioniSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=ahardmanborg")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            ahardmanborgSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            ahardmanborgSRs.Text = ahardmanborgSRs.Text.Remove(0, 6)

            If gcachioniSRs.Text < neves11SRs.Text Then
                assignTo.Text = ("GCACHIONI - GABRIEL CACHIONI")
            End If

            If gcachioniSRs.Text > neves11SRs.Text Then
                assignTo.Text = ("NEVES11 - CARLOS NEVES (SNOW)")
            End If

            If neves11SRs.Text = gcachioniSRs.Text Then
                If ahardmanborgSRs.Text < gcachioniSRs.Text Then
                    assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
                Else
                    ranNum = (r.Next(1, 3))
                    If ranNum = 1 Then
                        assignTo.Text = ("RECOMENDADO: GCACHIONI - GABRIEL CACHIONI")
                    End If
                    If ranNum = 2 Then
                        assignTo.Text = ("RECOMENDADO: NEVES11 - CARLOS NEVES (SNOW)")
                    End If
                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica WFO
        ' Logica OCEANA
        If prod.Text = "OCEANA" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=gcachioni")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            gcachioniSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            gcachioniSRs.Text = gcachioniSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)

            If gcachioniSRs.Text < fkhaguiwaraSRs.Text Then
                assignTo.Text = ("GCACHIONI - GABRIEL CACHIONI")
            End If

            If gcachioniSRs.Text > fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FKHAGUIWARA - FELIPE HAGUIWARE (HAGUIS)")
            End If

            If fkhaguiwaraSRs.Text = gcachioniSRs.Text Then
                ranNum = (r.Next(1, 3))
                If ranNum = 1 Then
                    assignTo.Text = ("RECOMENDADO: GCACHIONI - GABRIEL CACHIONI")
                End If
                If ranNum = 2 Then
                    assignTo.Text = ("RECOMENDADO: FKHAGUIWARA - FELIPE HAGUIWARE (HAGUIS)")
                End If

            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica OCEANA
        ' Logica OCEANALYTICS
        If prod.Text = "OCEANALYTICS" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=gcachioni")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            gcachioniSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            gcachioniSRs.Text = gcachioniSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)

            If gcachioniSRs.Text < fkhaguiwaraSRs.Text Then
                assignTo.Text = ("GCACHIONI - GABRIEL CACHIONI")
            End If

            If gcachioniSRs.Text > fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FKHAGUIWARA - FELIPE HAGUIWARE (HAGUIS)")
            End If

            If fkhaguiwaraSRs.Text = gcachioniSRs.Text Then
                ranNum = (r.Next(1, 3))
                If ranNum = 1 Then
                    assignTo.Text = ("RECOMENDADO: GCACHIONI - GABRIEL CACHIONI")
                End If
                If ranNum = 2 Then
                    assignTo.Text = ("RECOMENDADO: FKHAGUIWARA - FELIPE HAGUIWARE (HAGUIS)")
                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica OCEANALYTICS
        ' Logica EPM
        If prod.Text = "EPM" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=neves11")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            neves11SRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            neves11SRs.Text = neves11SRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=AHARDMANBORG")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            ahardmanborgSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            ahardmanborgSRs.Text = ahardmanborgSRs.Text.Remove(0, 6)

            If ahardmanborgSRs.Text < neves11SRs.Text Then
                assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
            End If

            If ahardmanborgSRs.Text > neves11SRs.Text Then
                assignTo.Text = ("NEVES11 - CARLOS NEVES (SNOW)")
            End If

            If ahardmanborgSRs.Text = neves11SRs.Text Then
                ranNum = (r.Next(1, 3))
                If ranNum = 1 Then
                    assignTo.Text = ("RECOMENDADO: AHARDMANBORG - Arthur Hardman (THUK)")
                End If
                If ranNum = 2 Then
                    assignTo.Text = ("RECOMENDADO: NEVES11 - CARLOS NEVES (SNOW)")
                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica EPM
        ' Logica EPM
        If prod.Text = "EPM" Then
            assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
            fimLog = True
            foundLog = True
        End If
        'fim Logica EPM
        ' Logica POM
        If prod.Text = "POM" Then
            assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
            fimLog = True
            foundLog = True
        End If
        'fim Logica POM
        ' Logica PDS
        If prod.Text = "PDS" Then
            assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
            fimLog = True
            foundLog = True
        End If
        'fim Logica PDS
        ' Logica APC
        If prod.Text = "APC" Then
            assignTo.Text = ("AHARDMANBORG - Arthur Hardman (THUK)")
            fimLog = True
            foundLog = True
        End If
        'fim Logica APC
        ' Logica CMS
        If prod.Text = "CMS" Then
            assignTo.Text = ("PASSOSE - EDSON PASSOS (PASSOSE)")
            fimLog = True
            foundLog = True
        End If
        'fim Logica CMS
        ' Logica IPO
        If prod.Text = "IPO" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fgsantos")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fgsantosSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fgsantosSRs.Text = fgsantosSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)


            If fgsantosSRs.Text > fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FKHAGUIWARA - Felipe Haguiwara (Haguis)")
            End If

            If fgsantosSRs.Text < fkhaguiwaraSRs.Text Then
                assignTo.Text = ("FGSANTOS - Fernando Gabriel (Robs)")
            End If

            If fkhaguiwaraSRs.Text = fgsantosSRs.Text Then
                ranNum = (r.Next(1, 3))
                If ranNum = 1 Then
                    assignTo.Text = ("RECOMENDADO: FKHAGUIWARA - Felipe Haguiwara (Haguis)")
                End If
                If ranNum = 2 Then
                    assignTo.Text = ("RECOMENDADO: FGSANTOS - Fernando Gabriel (Robs)")
                End If
            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica IPO
        ' Logica AES
        If prod.Text = "AES" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=gcachioni")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            gcachioniSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            gcachioniSRs.Text = gcachioniSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=AHARDMANBORG")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            ahardmanborgSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            ahardmanborgSRs.Text = ahardmanborgSRs.Text.Remove(0, 6)

            If ahardmanborgSRs.Text > gcachioniSRs.Text Then
                assignTo.Text = ("GCACHIONI - GABRIEL CACHIONI (CACHS)")
            End If

            If ahardmanborgSRs.Text < gcachioniSRs.Text Then
                assignTo.Text = ("AHARDMANBORG - ARTHUR HARDMAN (THUK)")
            End If

            If ahardmanborgSRs.Text = gcachioniSRs.Text Then
                ranNum = (r.Next(1, 3))
                If ranNum = 1 Then
                    assignTo.Text = ("RECOMENDADO: GCACHIONI - GABRIEL CACHIONI (CACHS)")
                End If
                If ranNum = 2 Then
                    assignTo.Text = ("RECOMENDADO: AHARDMANBORG - ARTHUR HARDMAN (THUK)")
                End If

            End If

            fimLog = True
            foundLog = True
        End If
        'fim Logica AES
        ' Logica CM
        If prod.Text = "CM" Then

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fgsantos")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fgsantosSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fgsantosSRs.Text = fgsantosSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=fkhaguiwara")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            fkhaguiwaraSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            fkhaguiwaraSRs.Text = fkhaguiwaraSRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=neves11")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            neves11SRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            neves11SRs.Text = neves11SRs.Text.Remove(0, 6)

            WebBrowser1.Navigate("https://report.avaya.com/siebelreports/openreport.aspx?dpGroup=All&dpEmpRegion=All&dpDispPCARs=0&dpSeverityNew=0&dpBillable=0&dpProdHouse=All&cbInactive=True&cbTestFls=True&dpType=All&dpOpenTea=0&dpAlarms=0&lbCapsule=All&lbSource=All&lbSalesTheatre=All&dpBacklogState=All&dpRegion=All&dpSubRegion=All&dpCountry=All&dpStatus=All&dpBp=0&cbIncludeCompletedCancelled=False&rblMsl=0&dpBacklogState2=8&lbActivities=&tbCoachId=passose")
            While WebBrowser1.ReadyState <> WebBrowser1.ReadyState.Complete
                Application.DoEvents()
            End While
            passoseSRs.Text = WebBrowser1.Document.GetElementById("cphMain_lRows").InnerText
            passoseSRs.Text = passoseSRs.Text.Remove(0, 6)



            Dim phrases As String() = {"FGSANTOS - SRs: " + fgsantosSRs.Text, "FKHAGUIWARA - SRs: " + fkhaguiwaraSRs.Text, "NEVES 11 - SRs: " + neves11SRs.Text, "PASSOSE - SRs: " + passoseSRs.Text}

            phrases = phrases.OrderByDescending(Function(q) Int32.Parse(q.Split(" ").Last)).ToArray()

            For Each s In phrases
                assignTo.Text = s
            Next
            fimLog = True
            foundLog = True
        End If
        'fim Logica CM

        If foundLog = False Then
            MessageBox.Show("Produto não roteavel para T2")
        End If
        Button4.Enabled = True

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ranNum = (r.Next(1, 4))
        user = My.User.Name
        user = user.Remove(0, 7)

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles siteName.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub dateCreated_Click(sender As Object, e As EventArgs) Handles dateCreated.Click

    End Sub

    Private Sub TextBox8_Leave(sender As Object, e As EventArgs) Handles TextBox8.Leave
        If TextBox8.Text = "" Then
            TextBox8.Text = "SR Number"
        End If
    End Sub
    Private Sub Textbox8_Enter(sender As Object, e As EventArgs) Handles TextBox8.Enter
        TextBox8.Text = ""
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Button3(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button3_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim time As String = InputBox("Digite quantos minutos deseja adicionar ao chamado.")

        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Try
                Dim omsg As Object
                omsg = Outl.CreateItem(0)
                omsg.To = "ssblinbox@avaya.com"
                omsg.subject = ("Update " + TextBox8.Text + " Time " + time)
                omsg.body = (user + " is routing this case to: " + assignTo.Text + ".")
                omsg.Send()
            Catch ex As Exception
                MessageBox.Show("Envio Cancelado.")
            End Try

        End If



    End Sub
End Class