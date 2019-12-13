Imports System.IO
Imports System.Net
Imports Newtonsoft.Json

Public Class DDFeAPI
    Public Property resp As String
    Private Shared token As String

    Public Property Ptoken() As String
        Get
            Return token
        End Get
        Set(ByVal valor As String)
            token = valor
        End Set
    End Property

    Public Shared Function enviaConteudoParaAPI(ByVal conteudo As String, ByVal url As String, ByVal tpConteudo As String) As String
        Dim retorno As String = ""
        Dim httpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        httpWebRequest.Method = "POST"
        httpWebRequest.Headers("X-AUTH-TOKEN") = token

        If tpConteudo = "txt" Then
            httpWebRequest.ContentType = "text/plain;charset=utf-8"
        ElseIf tpConteudo = "xml" Then
            httpWebRequest.ContentType = "application/xml;charset=utf-8"
        Else
            httpWebRequest.ContentType = "application/json;charset=utf-8"
        End If

        Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
            streamWriter.Write(conteudo)
            streamWriter.Flush()
            streamWriter.Close()
        End Using

        Try
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)

            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                retorno = streamReader.ReadToEnd()
            End Using

        Catch ex As WebException

            If ex.Status = WebExceptionStatus.ProtocolError Then
                Dim response As HttpWebResponse = CType(ex.Response, HttpWebResponse)

                Using streamReader = New StreamReader(response.GetResponseStream())
                    retorno = streamReader.ReadToEnd()
                End Using

                Select Case CInt(response.StatusCode)
                    Case 401
                        MessageBox.Show("Token não enviado ou inválido")
                    Case 403
                        MessageBox.Show("Token sem permissão")
                    Case 404
                        MessageBox.Show("Não encontrado, verifique o retorno para mais informações")
                    Case Else
                End Select
            End If
        End Try

        Return retorno
    End Function

    Public Shared Function manifestacao(ByVal caminho As String, ByVal CNPJInteressado As String, tpEvento As String, tpAmb As String, nsu As String, Optional chave As String = "",
                                        Optional xJust As String = "") As String

        Dim manifestacaoTipo As New ManifestacaoTipo
        If (tpEvento = "210240") Then
            manifestacaoTipo.tpEvento = tpEvento
            manifestacaoTipo.xJust = xJust
        Else
            manifestacaoTipo.tpEvento = tpEvento
        End If

        Dim manifestacaoJson As New ManifestacaoJson With {
            .CNPJInteressado = CNPJInteressado,
            .tpAmb = tpAmb,
            .manifestacao = manifestacaoTipo
        }

        If ((nsu = "") Or (nsu = "null")) Then
            manifestacaoJson.chave = chave
        Else
            manifestacaoJson.nsu = nsu
        End If

        Dim json As String = JsonConvert.SerializeObject(manifestacaoJson, New JsonSerializerSettings With {.NullValueHandling = NullValueHandling.Ignore})

        Dim url As String = "https://ddfe.ns.eti.br/events/manif"

        gravaLinhaLog("[MANIFESTAÇÃO_DADOS]")
        gravaLinhaLog(json)

        Dim resposta As String = enviaConteudoParaAPI(json, url, "json")
        gravaLinhaLog("[MANIFESTAÇÃO_RESPOSTA]")
        gravaLinhaLog(resposta)

        Dim jsonRetorno = JsonConvert.DeserializeObject(Of JsonRetornoManifestacao)(resposta)

        tratamentoManifestacao(jsonRetorno, tpEvento, chave, caminho)

        Return resposta

    End Function

    Public Shared Sub tratamentoManifestacao(ByVal jsonRetorno As Object, ByVal tpEvento As String, ByVal chave As String, ByVal caminho As String)


        Dim status As String = jsonRetorno.status
        Dim xMotivo As String
        If (status = "200") Then
            salvarDocManisfestacao(jsonRetorno, tpEvento, chave, caminho)
            xMotivo = jsonRetorno.retEvento.xMotivo
            MessageBox.Show(xMotivo)
        ElseIf (status = "-3") Then
            xMotivo = jsonRetorno.erro.xMotivo
            MessageBox.Show(xMotivo)
        Else
            MessageBox.Show(jsonRetorno.motivo)
        End If
    End Sub

    Public Shared Sub salvarDocManisfestacao(ByVal jsonRetorno As Object, ByVal tpEvento As String, ByVal chave As String, ByVal caminho As String)
        Dim xml As String = jsonRetorno.retEvento.xml
        salvarXML(xml, caminho, chave, "", tpEvento)
    End Sub


    Public Shared Function downloadUnico(ByVal caminho As String, ByVal CNPJInteressado As String, ByVal tpAmb As String,
                                         ByVal modelo As String, nsu As String, ByVal Optional chave As String = "", ByVal Optional incluirPDF As Boolean = False, Optional apenasComXml As Boolean = False,
                                         ByVal Optional comEventos As Boolean = False)

        Dim downloadUnicoJson As New DownloadUnicoJson With {
                 .CNPJInteressado = CNPJInteressado,
                 .tpAmb = tpAmb,
                 .incluirPDF = incluirPDF
        }


        If ((nsu <> "") Or (nsu <> "null")) Then
            downloadUnicoJson.nsu = nsu
            downloadUnicoJson.modelo = modelo
        Else
            downloadUnicoJson.chave = chave
            downloadUnicoJson.apenasComXml = apenasComXml
            downloadUnicoJson.comEventos = comEventos
        End If

        Dim json As String = JsonConvert.SerializeObject(downloadUnicoJson, New JsonSerializerSettings With {.NullValueHandling = NullValueHandling.Ignore})

        Dim url As String = "https://ddfe.ns.eti.br/dfe/unique"

        gravaLinhaLog("[DOWNLOAD_UNICO_DADOS]")
        gravaLinhaLog(json)

        Dim resposta As String = enviaConteudoParaAPI(json, url, "json")
        gravaLinhaLog("[DOWNLOAD_UNICO_RESPOSTA]")
        gravaLinhaLog(resposta)

        Dim jsonRetorno = JsonConvert.DeserializeObject(Of JsonRetornoDownloadUnico)(resposta)

        tratamentoDownloadUnico(caminho, incluirPDF, jsonRetorno)

        Return resposta

    End Function

    Public Shared Sub tratamentoDownloadUnico(ByVal caminho As String, ByVal incluirPDF As Boolean, ByVal jsonRetorno As Object)
        Dim status As String = jsonRetorno.status
        If (status = "200") Then
            salvarDocUnico(caminho, incluirPDF, jsonRetorno)
            MessageBox.Show("Download Unico feito com sucesso")
        Else
            MessageBox.Show(jsonRetorno.motivo)
        End If
    End Sub

    Public Shared Sub salvarDocUnico(ByVal caminho As String, ByVal incluirPDF As Boolean, ByVal jsonRetorno As Object)
        Dim listaDocs As String = jsonRetorno.listaDocs
        If (listaDocs = "false") Then
            Dim xml = jsonRetorno.xml
            Dim chave = jsonRetorno.chave
            Dim modelo = jsonRetorno.modelo
            salvarXML(xml, caminho, chave, modelo)

            If (incluirPDF = True) Then
                Dim pdf = jsonRetorno.pdf
                salvarPDF(pdf, caminho, chave, modelo)
            End If
        Else
            Dim tpEvento As String = ""
            Dim xmls = jsonRetorno.xmls
            For i As Integer = 0 To xmls.Count - 1
                If (i >= 1) Then
                    tpEvento = xmls(i).tpEvento
                End If
                Dim xml = xmls(i).xml
                Dim chave = xmls(i).chave
                Dim modelo = xmls(i).modelo
                If ((xml <> "") OrElse (xml <> vbNullString) OrElse (Len(xml) > 0)) Then
                    If ((incluirPDF = True) And (tpEvento = "")) Then
                        Dim pdf = xmls(i).pdf
                        salvarPDF(pdf, caminho, chave, modelo, tpEvento)
                        tpEvento = ""
                    End If
                    salvarXML(xml, caminho, chave, modelo, tpEvento)
                End If
            Next
        End If
    End Sub


    Public Shared Function downloadLote(ByVal caminho As String, ByVal CNPJInteressado As String, ByVal tpAmb As String,
                                        ByVal modelo As String, ByVal ultNSU As Integer, ByVal Optional incluirPdf As Boolean = False,
                                        ByVal Optional apenasComXml As Boolean = False,
                                        ByVal Optional comEventos As Boolean = False,
                                        ByVal Optional apenasPendManif As Boolean = False,
                                        ByVal Optional retornoSimples As Boolean = False) As String
        Dim downloadLoteParam As New DownloadLoteJson With {
            .CNPJInteressado = CNPJInteressado,
            .ultNSU = ultNSU,
            .tpAmb = tpAmb,
            .modelo = modelo,
            .incluirPDF = incluirPdf
        }
        If (apenasPendManif = True) Then
            downloadLoteParam.apenasPendManif = apenasPendManif
        Else
            downloadLoteParam.apenasComXml = apenasComXml
            downloadLoteParam.comEventos = comEventos
        End If

        Dim json As String = JsonConvert.SerializeObject(downloadLoteParam, New JsonSerializerSettings With {.NullValueHandling = NullValueHandling.Ignore})

        Dim url As String = "https://ddfe.ns.eti.br/dfe/bunch"

        gravaLinhaLog("[DOWNLOAD_LOTE_DADOS]")
        gravaLinhaLog(json)

        Dim resposta As String = enviaConteudoParaAPI(json, url, "json")

        gravaLinhaLog("[DOWNLOAD_LOTE_RESPOSTA]")
        gravaLinhaLog(resposta)

        Dim jsonRetorno As Object = JsonConvert.DeserializeObject(Of JsonRetornoDownloadLote)(resposta)
        Dim retorno As String = tratamentoDownloadLote(caminho, modelo, incluirPdf, jsonRetorno)

        If (retornoSimples = True) Then
            If (retorno <> vbNullString) Then
                resposta = retorno
            End If
        End If
        Return resposta
    End Function

    Public Shared Function tratamentoDownloadLote(ByVal caminho As String, ByVal modelo As String, ByVal incluirPdf As Boolean, ByVal jsonRetorno As Object) As String
        Dim status = jsonRetorno.status
        If (status = "200") Then
            Dim chRet() As String = salvarDocsLote(caminho, modelo, incluirPdf, jsonRetorno)
            Dim chaves() As String = chRet.Where(Function(palavra) Not String.IsNullOrEmpty(palavra)).ToArray()
            Dim downloadLoteRet As New DownloadLoteRet With {
                .status = status,
                .ultNSU = jsonRetorno.ultNSU,
                .chaves = chaves
            }
            Dim json As String = JsonConvert.SerializeObject(downloadLoteRet, Formatting.Indented, New JsonSerializerSettings With {.NullValueHandling = NullValueHandling.Ignore})
            MessageBox.Show("Download em Lote feito com sucesso")
            Return json
        Else
            MessageBox.Show(jsonRetorno.motivo.ToString)
            Return vbNullString
        End If
    End Function

    Public Shared Function salvarDocsLote(ByVal caminho As String, ByVal modelo As String, ByVal incluirPdf As Boolean, ByVal jsonRetorno As Object)
        Dim xmls = jsonRetorno.xmls
        Dim chaves(xmls.Count - 1) As String
        For i As Integer = 0 To xmls.Count - 1
            Dim xml = xmls(i).xml
            If (xml = "") OrElse (xml = vbNullString) OrElse (Len(xml) = 0) Then
                Continue For
            End If
            Dim tpEvento = ""
            chaves(i) = xmls(i).chave
            If (InStr("tpEvento", xmls(i).ToString)) Then
                tpEvento = xmls(i).tpEvento
            Else
                If ((incluirPdf = True) And (tpEvento = "")) Then
                    Dim pdf = xmls(i).pdf
                    salvarPDF(pdf, caminho, chaves(i), modelo, tpEvento)
                    tpEvento = ""
                End If
            End If
            salvarXML(xml, caminho, chaves(i), modelo, tpEvento)
        Next
        Return chaves
    End Function



    Public Shared Sub salvarXML(ByVal xml As String, ByVal caminho As String, ByVal chave As String, ByVal modelo As String, ByVal Optional tpEvento As String = "")
        Try
            If Not Directory.Exists(caminho) Then Directory.CreateDirectory(caminho)
            If Not caminho.EndsWith("\") Then caminho += "\"
        Catch ex As IOException
            gravaLinhaLog("[CRIAR_DIRETORIO]" & caminho)
            gravaLinhaLog(ex.Message)
            Throw New Exception("Erro: " & ex.Message)
        End Try

        Dim extencao As String
        If (modelo = "55") Then
            extencao = "-procNFe.xml"
        ElseIf (modelo = "57") Then
            extencao = "-procCTe.xml"
        ElseIf (modelo = "98") Then
            extencao = "-procNFSe.xml"
        Else
            extencao = "-procEven.xml"
        End If

        Dim localParaSalvar As String = caminho & tpEvento & chave & extencao
        Dim ConteudoSalvar As String = ""
        ConteudoSalvar = xml.Replace("\""", "")
        File.WriteAllText(localParaSalvar, ConteudoSalvar)
    End Sub

    Public Shared Sub salvarPDF(ByVal pdf As String, ByVal caminho As String, ByVal chave As String, ByVal modelo As String, ByVal Optional tpEvento As String = "")
        Try
            If Not Directory.Exists(caminho) Then Directory.CreateDirectory(caminho)
            If Not caminho.EndsWith("\") Then caminho += "\"
        Catch ex As IOException
            gravaLinhaLog("[CRIAR_DIRETORIO]" & caminho)
            gravaLinhaLog(ex.Message)
            Throw New Exception("Erro: " & ex.Message)
        End Try

        Dim extencao As String
        If (modelo = "55") Then
            extencao = "-procNFe.pdf"
        ElseIf (modelo = "57") Then
            extencao = "-procCTe.pdf"
        Else
            extencao = "-procNFSe.pdf"
        End If

        Dim localParaSalvar As String = caminho & tpEvento & chave & extencao
        Dim bytes As Byte() = Convert.FromBase64String(pdf)
        If File.Exists(localParaSalvar) Then File.Delete(localParaSalvar)
        Dim stream As FileStream = New FileStream(localParaSalvar, FileMode.CreateNew)
        Dim writer As BinaryWriter = New BinaryWriter(stream)
        writer.Write(bytes, 0, bytes.Length)
        writer.Close()
    End Sub

    Public Shared Sub gravaLinhaLog(ByVal conteudo As String)
        Dim caminho As String = Application.StartupPath & "\log\"
        Console.Write(caminho)

        If Not Directory.Exists(caminho) Then
            Directory.CreateDirectory(caminho)
        End If

        Dim data As String = DateTime.Now.ToShortDateString()
        Dim hora As String = DateTime.Now.ToShortTimeString()
        Dim nomeArq As String = data.Replace("/", "")

        Using outputFile As StreamWriter = New StreamWriter(caminho & nomeArq & ".txt", True)
            outputFile.WriteLine(data & " " & hora & " - " & conteudo)
        End Using
    End Sub




    Public Class JsonRetornoDownloadLote
        Public status As String
        Public ultNSU As String
        Public xmls As IList(Of JsonRetornoXmls)
        Public motivo As String
    End Class

    Public Class DownloadLoteJson
        Public CNPJInteressado As String
        Public tpAmb As String
        Public ultNSU As Integer
        Public modelo As String
        Public incluirPDF As Boolean
        Public apenasComXml As Boolean
        Public comEventos As Boolean
        Public apenasPendManif As Boolean
        Public removerEventosCodigos As Boolean
    End Class

    Public Class DownloadLoteRet
        Public status As String
        Public ultNSU As String
        Public chaves() As String
    End Class

    Public Class JsonRetornoDownloadUnico
        Public status As String
        Public listaDocs As String
        Public nsu As String
        Public chave As String
        Public modelo As String
        Public vNF As String
        Public xml As String
        Public pdf As String
        Public motivo As String
        Public xmls As IList(Of JsonRetornoXmls)
    End Class

    Public Class JsonRetornoXmls
        Public nsu As String
        Public chave As String
        Public modelo As String
        Public tpEvento As String
        Public vNF As String
        Public xml As String
        Public pdf As String
    End Class

    Public Class DownloadUnicoJson
        Public CNPJInteressado As String
        Public tpAmb As String
        Public nsu As String = vbNullString
        Public chave As String = vbNullString
        Public modelo As String = vbNullString
        Public incluirPDF As Boolean
        Public apenasComXml As Boolean
        Public comEventos As Boolean
    End Class

    Public Class JsonRetornoManifestacao
        Public status As String
        Public motivo As String
        Public retEvento As New JsonRetEventoManifesto
        Public erro As New JsonErroManifesto
    End Class

    Public Class JsonRetEventoManifesto
        Public cStat As String
        Public xMotivo As String
        Public dhRegEvento As String
        Public nProt As String
        Public xml As String
    End Class

    Public Class JsonErroManifesto
        Public cStat As String
        Public xMotivo As String
    End Class

    Public Class ManifestacaoJson
        Public CNPJInteressado As String
        Public nsu As String = vbNullString
        Public chave As String = vbNullString
        Public tpAmb As String
        Public manifestacao As New ManifestacaoTipo
    End Class

    Public Class ManifestacaoTipo
        Public tpEvento As String
        Public xJust As String = vbNullString
    End Class
End Class

