Option Explicit On
Imports System.Net.Mime.MediaTypeNames

Const sSiteName = "https://www.mcmaster.com/2835T17/"

Private Sub getHTMLContents()
    ' Create Internet Explorer object.
    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False          ' Keep this hidden.

    IE.navigate sSiteName

    ' Wait till IE is fully loaded.
    While IE.readyState <> 4
        DoEvents
    Wend
   
    Dim oHDoc As HTMLDocument     ' Create document object.
    Set oHDoc = IE.Document
   
    Dim oHEle As HTMLTextElement
    Set oHEle = oHDoc.getElementById("Prce")

    ' Print the extracted price
    Application.Wait(Now + TimeValue("00:00:00001"))

    Dim priceElement As HTMLTextElement
    Set priceElement = oHDoc.getElementsByClassName("PrceTxt")(0)

    Dim priceText As String
    priceText = priceElement.innerText

    Debug.Print priceText

    Cells(2, 4) = priceText

    ' Clean up.
    IE.Quit
    Set IE = Nothing
    Set oHEle = Nothing
    Set oHDoc = Nothing
End Sub