Imports System.IO

Module Principal

    Sub Main(ByVal comandos() As String)
        If comandos.Count > 0 Then
            Dim DDFeAPI As New DDFeAPI
            Dim resposta As String
            Dim argumentos As String = ""
            Dim token As String = comandos(1)

            DDFeAPI.gravaLinhaLog("[ARGUMENTOS_CAPTURADOS]")
            For i As Integer = 0 To comandos.Count - 1
                If (i = 1) Then
                    argumentos &= "[token] "
                Else
                    argumentos &= comandos(i) & " "
                End If
            Next
            DDFeAPI.gravaLinhaLog(argumentos)
            DDFeAPI.Ptoken = token

            If (comandos(0).ToUpper() = "ENVIACONTEUDOPARAAPI") Then
                Dim caminho As String = comandos(2)
                Dim url As String = comandos(3)
                Dim tpConteudo As String = comandos(4)
                Dim id As String = comandos(5)

                Dim conteudo = LerArquivoTexto(caminho)

                resposta = DDFeAPI.enviaConteudoParaAPI(conteudo, url, tpConteudo)

                EscreverArquivoTexto("enviaConteudoParaAPI_" & id & ".json", resposta)

            ElseIf (comandos(0).ToUpper() = "MANIFESTACAO") Then
                Dim caminho As String = comandos(2)
                Dim CNPJInteressado As String = comandos(3)
                Dim tpEvento As String = comandos(4)
                Dim tpAmb As String = comandos(5)
                Dim nsu As String = comandos(6)
                Dim chave As String = comandos(7)
                Dim xJust As String = comandos(8)
                Dim id As String = comandos(9)

                DDFeAPI.gravaLinhaLog("ret")
                resposta = DDFeAPI.manifestacao(caminho, CNPJInteressado, tpEvento, tpAmb, nsu, chave, xJust)
                DDFeAPI.gravaLinhaLog(resposta)
                EscreverArquivoTexto("manifestacao_" & id & ".json", resposta)

            ElseIf (comandos(0).ToUpper() = "DOWNLOADUNICO") Then
                DDFeAPI.gravaLinhaLog("json")
                Dim caminho As String = comandos(2)
                Dim CNPJInteressado As String = comandos(3)
                Dim tpAmb As String = comandos(4)
                Dim modelo As String = comandos(5)
                Dim nsu As String = comandos(6)
                Dim chave As String = comandos(7)
                Dim incluirPDF As Boolean = comandos(8)
                Dim apenasComXml As Boolean = comandos(9)
                Dim comEventos As Boolean = comandos(10)
                Dim id As String = comandos(11)

                resposta = DDFeAPI.downloadUnico(caminho, CNPJInteressado, tpAmb, modelo, nsu, chave, incluirPDF, apenasComXml, comEventos)

                EscreverArquivoTexto("downloadUnico_" & id & ".json", resposta)

            ElseIf (comandos(0).ToUpper() = "DOWNLOADLOTE") Then
                Dim caminho As String = comandos(2)
                Dim CNPJInteressado As String = comandos(3)
                Dim tpAmb As String = comandos(4)
                Dim modelo As String = comandos(5)
                Dim ultNSU As Integer = comandos(6)
                Dim incluirPdf As Boolean = comandos(7)
                Dim apenasComXml As Boolean = comandos(8)
                Dim comEventos As Boolean = comandos(9)
                Dim apenasPendManif As Boolean = comandos(10)
                Dim retornoSimples As Boolean = comandos(11)
                Dim id As String = comandos(12)

                resposta = DDFeAPI.downloadLote(caminho, CNPJInteressado, tpAmb, modelo, ultNSU, incluirPdf, apenasComXml, comEventos, apenasPendManif, retornoSimples)

                EscreverArquivoTexto("downloadLote_" & id & ".json", resposta)

            Else
                FinalizarApp()

            End If
        End If

        FinalizarApp()
    End Sub

    Private Function LerArquivoTexto(caminho As String) As String
        Dim conteudo = ""
        If File.Exists(caminho) Then
            conteudo = My.Computer.FileSystem.ReadAllText(caminho)
        Else
            MessageBox.Show("Arquivo " & caminho & " não existe ou não foi encontrado.")
        End If

        LerArquivoTexto = conteudo
    End Function

    Private Sub EscreverArquivoTexto(nome As String, conteudo As String)
        Dim caminho = Application.StartupPath & "\respostas\"
        Dim nomeArquivo = caminho & nome

        Try
            If Not Directory.Exists(caminho) Then
                Directory.CreateDirectory(caminho)
            End If

        Catch ex As IOException
            DDFeAPI.gravaLinhaLog("Não foi possível criar o diretório 'respostas'")
            MessageBox.Show("Não foi possível criar o diretório 'respostas'")
        End Try

        System.IO.File.WriteAllText(nomeArquivo, conteudo)
    End Sub

    Private Sub FinalizarApp()
        Application.ExitThread()
    End Sub
End Module
