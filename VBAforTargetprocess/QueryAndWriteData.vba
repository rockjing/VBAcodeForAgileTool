Sub getReportData()

Call clearRawData

Sheet1.Range("a1").Value = "Id"
Sheet1.Range("b1").Value = "name"

Sheet1.Range("c1").Value = "Entity State"
Sheet1.Range("d1").Value = "state"

Sheet1.Range("e1").Value = "sprint"
Sheet1.Range("f1").Value = "srpint2"

Sheet1.Range("g1").Value = "feature"

Sheet1.Range("h1").Value = "Entity Type"

Sheet1.Range("i1").Value = "Effort"


Dim appendUrl As String

appendUrl = "UserStories?include=[Id,Name,EntityState,iteration,feature,effort]&take=1000&where=Project.Name%20eq%20%27%E4%BC%98%E7%89%A7%E5%B9%B3%E5%8F%B0%27"

Set result = queryData(appendUrl)

Call writeStory(result)


End Sub



Sub writeStory(result As Variant)

Dim i As Integer
i = 1
For Each Node In result.SelectNodes("//UserStory")

If IsNull(Node) Or Node Is Nothing Then
Exit For
End If




i = i + 1
Sheet1.Range("a" & i) = Node.Attributes.getNamedItem("Id").Value
Sheet1.Range("b" & i) = Node.Attributes.getNamedItem("Name").Value




'Set tmpNode = Node.ChildNodes.Item(0) ' state
'Sheet1.Range("c" & i) = tmpNode.Attributes.getNamedItem("Name").Value
'Sheet1.Range("d" & i) = tmpNode.Attributes.getNamedItem("Name").Value



Sheet1.Range("c" & i).Value = getValue(True, "EntityState", Node)
Sheet1.Range("d" & i) = Sheet1.Range("c" & i)


Sheet1.Range("e" & i) = getValue(True, "Iteration", Node)
Sheet1.Range("f" & i) = Sheet1.Range("e" & i)
 

Sheet1.Range("g" & i) = getValue(True, "Feature", Node)

 

Sheet1.Range("h" & i) = "User Story"

Sheet1.Range("i" & i) = getValue(False, "Effort", Node)

Next


End Sub



Sub clearRawData()

Sheet1.Cells.Clear


End Sub

Function getValue(isAttribute As Boolean, elementName As String, Node As Variant) As String

Dim cNode As Variant
Dim tmpNode As Variant
Dim hasFound As Boolean

For Each cNode In Node.ChildNodes
  
  If elementName = cNode.BaseName Then
  hasFound = True
  Set tmpNode = cNode
  Exit For
  End If
Next

If hasFound <> True Then
 getValue = "NOT Found!"
 
End If



If isAttribute = True Then

If IsEmpty(tmpNode) Or IsNull(tmpNode) Or tmpNode Is Nothing Or IsNull(tmpNode.Attributes.getNamedItem("Name")) Or tmpNode.Attributes.getNamedItem("Name") Is Nothing Then
Else

getValue = tmpNode.Attributes.getNamedItem("Name").Value

End If
Else
getValue = tmpNode.Text
End If


End Function


Function queryData(appendUrl) As Variant


Dim url As String
Dim user As String
Dim password As String

url = Sheet3.Range("b2").Value
user = Sheet3.Range("b3").Value
password = Sheet3.Range("b4").Value


Set objHTTP = CreateObject("MSXML2.XMLHTTP")
objHTTP.Open "GET", url + "/api/v1/" + appendUrl, False, user, password
objHTTP.send
Set result = objHTTP.responseXML

Set queryData = result



End Function
