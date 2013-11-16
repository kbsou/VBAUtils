####  xlJson

  Used to read and write Json files through Excel, eliminating extra steps like converting Json to xml or vice versa

#### Example of usage


```vb

Const jsonString As String = "{" & _
        """firstName"": ""John"", " & _
        """lastName"": ""Smith"", " & _
        """age"": 32, " & _
        """address"": { " & _
            """streetAddress"": ""21 2nd Street"", " & _
            """city"": ""New York"", " & _
            """state"": ""NY"", " & _
            """postalCode"": 10021 " & _
        "}, " & _
        """phones"": [ " & _
            "{ " & _
            "   ""type"": ""home"", " & _
            "   ""phoneNumber"": ""212 555-1234"" " & _
            "}, " & _
            "{ " & _
            "   ""type"": ""fax"", " & _
            "   ""phoneNumber"": ""646 555-4567"" " & _
            "} " & _
        "] " & _
    "} "


Private Sub testJsonObject()
    
    Dim Jobj As New xlJson, _
        phones As Variant, _
        phone As Variant

    'Decodes the Json string to create the object
    Jobj.Decode jsonString
    
    ' § is a shortcut to properties
    Debug.Assert Jobj.§("firstName") = Jobj.Properties.firstName
    
    With Jobj
        
        Debug.Print "Name: " & .§.lastName & ", " & .§.firstName
        'Alternative way to call:
        Debug.Print "City " & .§("address").city
        
        'Adds a new property to the object
        .AddProperty "email", "john@smith.com"
        
        Debug.Print "Contact:"
        'Assigns the phones array to a variant
        phones = Jobj.ToArray(.§.phones)
        For Each phone In phones
            ' Since type is a reserved keyword, the VBA IDE will change
            ' "type" to "Type", causing an error since JScript is case sensitive,
            ' and so we'll have to use the CallByName method instead of writing phone.type
            Debug.Print vbTab & CallByName(phone, "type", VbGet) & ": " & phone.phoneNumber
         Next phone
         
        'gets the property we just added
        Debug.Print vbTab & "email: "; .§("email")
        
        'Changes the age property
        .§.age = 33
        
        Stop
        'Reencodes the object with the changes
        Debug.Print .Encode
        
        Stop
        'Dumps the object to a file:
        .Dump ThisWorkbook.path & "\john_smith.json"
    End With
    
End Sub

```

#### Getting content from an Url

You can use xlJson to get content from an Url that returns a Json string. The following example prints some info from GitHub API

```vb
Sub urltest()

    Dim jObj As New xlJson
    
    jObj.DecodeFromUrl "https://api.github.com/users/kbsou/repos"
    Debug.Print j.§(0).full_name
    
End Sub
```