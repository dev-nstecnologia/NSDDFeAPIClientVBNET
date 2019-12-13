Public Class Form1
    Private Sub btnManifestacao_Click(sender As Object, e As EventArgs) Handles btnManifestacao.Click
        'DDFeAPI.manifestacao("68584077000351", "210200", "", "2", "33181068584077000513550010000020891857216623")
        'DDFeAPI.downloadUnico("07364617000135", "C:\Users\matheus.mazzoni\Desktop\mazzoni", "2", "55", "472")
        DDFeAPI.downloadLote("07364617000135", "C:\Users\matheus.mazzoni\Desktop\mazzoni", "2", 0, "55")
    End Sub
End Class
