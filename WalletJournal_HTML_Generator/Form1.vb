Imports System.Net
Imports System.Xml
Imports System.IO

Public Class Form1
    Dim web As New WebClient()
    Dim walletTransactions As New List(Of XElement)
    Dim timestamps As New List(Of String)
    Dim refID As New List(Of Integer)
    Dim transactionSender As New List(Of String)


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GetDocument()
        Timer1.Enabled = True
    End Sub

    Public Sub GetDocument()
        Dim url As String = "https://api.eveonline.com/char/WalletJournal.xml.aspx?keyID=3431927&vCode=A6L29SRsXNWljWTMBbECa9BQ8NXYHY4pNcCAkiQeyBxQxUGzXGSbqYrX9gnyRg7f&characterID=94201985"
        AddHandler web.DownloadStringCompleted, AddressOf DownloadComplete

        Try
            web.DownloadStringAsync(New Uri(url))
        Catch ex As Exception

        End Try
    End Sub

    Public Sub DownloadComplete(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        RemoveHandler web.DownloadStringCompleted, AddressOf DownloadComplete
        Dim xmlDoc As XElement
        Dim writer As New StreamWriter("index.html")
        Dim strReason As String = Nothing
        Dim dblAmount As Double = 0.0
        Dim currentTime = Date.Now(), cachedUntil As Date
        Dim dateDiff As TimeSpan
        Dim msResult As Integer = 0

        Try
            xmlDoc = XElement.Parse(e.Result)
            walletTransactions = xmlDoc...<row>.ToList()
            Dim quickwrite As New StreamWriter("xml.xml")
            quickwrite.WriteLine(xmlDoc.ToString())
            quickwrite.Close()

            ' this portion controls the interval at which the API endpoint is polled
            ' 10 seconds are added to make sure the CCP server has the XML doc ready
            currentTime = currentTime.ToUniversalTime()
            cachedUntil = xmlDoc...<cachedUntil>.Value
            dateDiff = cachedUntil - currentTime
            msResult = dateDiff.TotalMilliseconds
            lblPull.Text = DateAdd(DateInterval.Second, 10, cachedUntil).ToString() & " UTC"
            Timer1.Enabled = False
            Timer1.Interval = msResult + 10000
            Timer1.Enabled = True

            writer.WriteLine("<!DOCTYPE HTML><html><head><link href='style.css' rel='stylesheet'><meta charset='UTF-8'></head><body>")
            writer.WriteLine("<table><thead><tr><td>Timestamp</td><td>Sender</td><td>Receiver</td><td>Amount</td><td>Reason</td></tr></thead><tbody>")

            For counter As Integer = 0 To walletTransactions.Count - 1 Step 1

                If walletTransactions.Item(counter).@ownerName2.ToString().Equals("Thirtyone Organism") Or
                    walletTransactions.Item(counter).@ownerName1.ToString().Equals("Thirtyone Organism") Then
                    Continue For
                End If

                dblAmount = CDbl(walletTransactions.Item(counter).@amount)

                If walletTransactions.Item(counter).@reason.ToString().Contains("IDS") And dblAmount >= 1000000000 Then
                    writer.WriteLine("<tr><td class='t'>" & AddTime(walletTransactions.Item(counter).@date) & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                                     "<td>-" & (dblAmount * 4).ToString("N2") & "</td>" &
                                     "<td class='w'>4x winner!</td></tr>")

                ElseIf walletTransactions.Item(counter).@reason.ToString().Contains("IDS") And dblAmount >= 500000000 Then
                    writer.WriteLine("<tr><td class='t'>" & AddTime(walletTransactions.Item(counter).@date) & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                                     "<td>-" & (dblAmount * 3).ToString("N2") & "</td>" &
                                     "<td class='w'>3x winner!</td></tr>")

                ElseIf walletTransactions.Item(counter).@reason.ToString().Contains("IDS") And dblAmount >= 200000000 Then
                    writer.WriteLine("<tr><td class='t'>" & AddTime(walletTransactions.Item(counter).@date) & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                                     "<td>-" & (dblAmount * 2.5).ToString("N2") & "</td>" &
                                     "<td class='w'>2.5x winner!</td></tr>")

                ElseIf walletTransactions.Item(counter).@reason.ToString().Contains("IDS") And dblAmount >= 50000000 Then
                    writer.WriteLine("<tr><td class='t'>" & AddTime(walletTransactions.Item(counter).@date) & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                                     "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                                     "<td>-" & (dblAmount * 2).ToString("N2") & "</td>" &
                                     "<td class='w'>2x winner!</td></tr>")

                    ' ElseIf walletTransactions.Item(counter).@reason.ToString().Contains("IDS") And dblAmount >= 10000000 Then
                    '     writer.WriteLine("<tr><td class='t'>" & AddTime(walletTransactions.Item(counter).@date) & "</td>" &
                    '                      "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                    '                      "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                    '                      "<td>-" & (dblAmount * 1.5).ToString("N2") & "</td>" &
                    '                      "<td class='w'>1.5x winner!</td></tr>")
                End If

                If walletTransactions.Item(counter).@reason.ToString().Contains("DESC: ") Then
                    strReason = walletTransactions.Item(counter).@reason.ToString()
                    strReason = strReason.Replace("DESC: ", "")
                    walletTransactions.Item(counter).@reason = strReason
                End If

                writer.WriteLine("<tr><td class='t'>" & walletTransactions.Item(counter).@date.ToString() & "</td>" &
                                 "<td>" & walletTransactions.Item(counter).@ownerName1.ToString() & "</td>" &
                                 "<td>" & walletTransactions.Item(counter).@ownerName2.ToString() & "</td>" &
                                 "<td>" & dblAmount.ToString("N2") & "</td>" &
                                 "<td>" & walletTransactions.Item(counter).@reason.ToString() & "</td></tr>")

            Next

            writer.WriteLine("</tbody></table>*Information on this page can take up to 30 minutes to update.</body></html>")
            writer.Close()

        Catch ex As Exception

        End Try
    End Sub

    Public Function AddTime(ByVal iTime As Date) As String
        Dim dateString As String = Nothing
        Dim addedSeconds As Integer = 0

        Randomize()
        addedSeconds = 14
        iTime = iTime.AddSeconds(addedSeconds)

        dateString = Format(iTime, "yyyy-MM-dd HH:mm:ss")

        Return dateString
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        GetDocument()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        web.Proxy() = Nothing
        Timer1.Interval = 60000
    End Sub
End Class
