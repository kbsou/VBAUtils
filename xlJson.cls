VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************** Class xlJson ***********************************************+
'                                                                                                   '
'   Version 0.1                                                                                     '
'   Contains methods and functions for using JSon data objects                                      '
'                                                                                                   '
'======================================== License ==================================================+
'                                                                                                   '
'   This piece of code is licensed MIT License                                                      '
'                                                                                                   '
' Contributors:                                                                                     '
'   kbsou - https://github.com/kbsou - xlJson v0.1                                                  '
'                                                                                                   '
'======================================== IMPORTANT ================================================+
' This class uses refs to: (They will be added automatically)                                       '
'       Microsoft Scripting Runtime                                                                 '
'       Microsoft Script Control                                                                    '
'===================================================================================================+

Option Explicit


Private engine As ScriptControl
Public JObject As Object


Private Sub Class_Initialize()
    
    Set engine = New ScriptControl
    engine.Language = "JScript"
        
    engine.AddCode _
        "function getKeys(jsonObj) {" & _
            " var keys = new Array(); " & _
            " for (var i in jsonObj) { keys.push(i); } " & _
            " return keys; " & _
        "} "
        
    engine.AddCode _
        "function getType(jsonObj, property) { " & _
            " return typeof(jsonObj[property]); " & _
         "} "
         
    engine.AddCode _
        "function encode(jsonObj){" & _
            "function menc(obj, array) {" & _
                "var encoded = '';" & _
                "array = array || false;" & _
                "for (var key in obj) {" & _
                    "var val = obj[key];" & _
                    "key = (!array)? " & quote("key") & " + ': ': '';" & _
                    "if ( typeof val == 'object'){" & _
                        "if ( val instanceof Array ) {encoded += key + '[' + menc(val, true) + '],';}" & _
                        "else { encoded += key + '{' + menc(val) + '}, ';}" & _
                    "}" & _
                    "else{" & _
                        "if(typeof val == 'string'){ encoded += key + " & quote("val") & " + "","" ;}" & _
                        "else{encoded += key + val + ',';}" & _
                    "}" & _
                "}" & _
                "return encoded.slice(0, ((array)?(-2):(-1)) );" & _
            "}" & _
            "return '{' + menc(jsonObj) + '}';" & _
        "}"
    
    engine.AddCode _
        "function addKey(jsonObj, key, value) { " & _
            " jsonObj[key] = value; " & _
            " return jsonObj;" & _
        "}"
    
    engine.AddCode _
        "function removeKey(jsonObj, key) { " & _
            " var json = jsonObj; " & _
            " delete json[key]; " & _
            " return json " & _
        "}"
        
End Sub

Private Function quote(word As String)
    quote = "'""' + " & word & " + '""'"
End Function

Public Sub Decode(jsonString As String)
    Set JObject = engine.Eval("(" + jsonString + ")")
End Sub

Public Function Encode() As String
    Encode = engine.Run("encode", JObject)
End Function

Public Function GetType(Property As String)
    GetType = engine.Run("getType", Property)
End Function

Property Get Properties(Property As String) As Variant

    If IsObject(CallByName(JObject, Property, VbGet)) Then
        Set Properties = CallByName(JObject, Property, VbGet)
    Else
        Properties = CallByName(JObject, Property, VbGet)
    End If
    
End Property

Property Let Properties(Property As String, val As Variant)

    If IsObject(value) Then
        CallByName JObject, Property, VbSet, value
    Else
        CallByName JObject, Property, VbLet, value
    End If
    
End Property

'§ is a shortcut to properties
Property Get §(Property As String) As Variant

    If IsObject(CallByName(JObject, Property, VbGet)) Then
        Set § = CallByName(JObject, Property, VbGet)
    Else
        § = CallByName(JObject, Property, VbGet)
    End If
    
End Property

Property Let §(Property As String, value As Variant)

    If IsObject(value) Then
        CallByName JObject, Property, VbSet, value
    Else
        CallByName JObject, Property, VbLet, value
    End If
    
End Property

Public Sub RemoveProperty(propertyName As String)
    Set JObject = engine.Run("removeKey", JObject, propertyName)
End Sub

Public Sub AddProperty(propertyName As String, value As Variant)
    Set JObject = engine.Run("addKey", JObject, propertyName, value)
End Sub

Public Function ToArray(Optional arr As Object = Nothing) As Variant()
    
    Dim i As Integer, _
        size As Integer, _
        resp() As Variant
    
    If arr Is Nothing Then Set arr = JObject
    
    size = CallByName(arr, "length", VbGet) - 1
    
    ReDim resp(size)
    ReDim ToArray(size)
    
    For i = 0 To size
        If IsObject(CallByName(arr, i, VbGet)) Then
            Set resp(i) = CallByName(arr, i, VbGet)
        Else
            resp(i) = CallByName(arr, i, VbGet)
        End If
    Next i
    
    ToArray = resp
    
End Function

Public Function GetKeys() As String()
    
    Dim Keys As String, _
        key As Variant

    For Each key In engine.Run("getKeys", JObject)
        Keys = Keys & key & ";"
    Next
   
    GetKeys = Split(Left(Keys, Len(Keys) - 1), ";")
    
End Function

Public Sub Dump(path As String)
' Writes the encoded object to a file given by path
                          
    Dim fso As New Scripting.FileSystemObject, _
        ts As Scripting.TextStream
    
    Set ts = fso.OpenTextFile(path, ForWriting, True)
    ts.Write Me.Encode
    ts.Close
 
End Sub

