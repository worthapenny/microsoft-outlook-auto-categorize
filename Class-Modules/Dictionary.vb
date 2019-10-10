Option Explicit

Public KeyValuePairs As Collection ' open access but allows iteration
Public Tag As Variant            ' read/write unrestricted

Private Sub Class_Initialize()
   Set KeyValuePairs = New Collection
End Sub

Private Sub Class_Terminate()
   Set KeyValuePairs = Nothing
End Sub

' in Scripting.Dictionary this is writeable, here we have only vbtextCompare because we are using a Collection
Public Property Get CompareMode() As VbCompareMethod
   CompareMode = vbTextCompare   '=1; vbBinaryCompare=0
End Property

Public Property Let Item(key As String, Item As Variant)    ' dic.Item(Key) = value ' update a scalar value for an existing key
   Let KeyValuePairs.Item(key).value = Item
End Property

Public Property Set Item(key As String, Item As Variant)    ' Set dic.Item(Key) = value ' update an object value for an existing key
   Set KeyValuePairs.Item(key).value = Item
End Property

Public Property Get Item(key As String) As Variant
   AssignVariable Item, KeyValuePairs.Item(key).value
End Property

' Collection parameter order is Add(Item,Key); Dictionary is Add(Key,Item) so always used named arguments
Public Sub Add(key As String, Item As Variant)
   Dim oKVP As KeyValuePair
   Set oKVP = New KeyValuePair
   oKVP.key = key
   If IsObject(Item) Then
      Set oKVP.value = Item
   Else
      Let oKVP.value = Item
   End If
   KeyValuePairs.Add Item:=oKVP, key:=key
End Sub

Public Property Get Exists(key As String) As Boolean
   On Error Resume Next
   Exists = TypeName(KeyValuePairs.Item(key)) > ""  ' we can have blank key, empty item
End Property

Public Sub Remove(key As String)
   'show error if not there rather than On Error Resume Next
   KeyValuePairs.Remove key
End Sub

Public Sub RemoveAll()
   Set KeyValuePairs = Nothing
   Set KeyValuePairs = New Collection
End Sub

Public Property Get count() As Long
   count = KeyValuePairs.count
End Property

Public Property Get Items() As Variant     ' for compatibility with Scripting.Dictionary
Dim vlist As Variant, i As Long
If Me.count > 0 Then
   ReDim vlist(0 To Me.count - 1) ' to get a 0-based array same as scripting.dictionary
   For i = LBound(vlist) To UBound(vlist)
      AssignVariable vlist(i), KeyValuePairs.Item(1 + i).value ' could be scalar or array or object
   Next i
   Items = vlist
End If
End Property

Public Property Get keys() As String()
Dim vlist() As String, i As Long
If Me.count > 0 Then
   ReDim vlist(0 To Me.count - 1)
   For i = LBound(vlist) To UBound(vlist)
      vlist(i) = KeyValuePairs.Item(1 + i).key   '
   Next i
   keys = vlist
End If
End Property

Public Property Get KeyValuePair(index As Long) As Variant  ' returns KeyValuePair object
    Set KeyValuePair = KeyValuePairs.Item(1 + index)            ' collections are 1-based
End Property

Private Sub AssignVariable(variable As Variant, value As Variant)
   If IsObject(value) Then
      Set variable = value
   Else
      Let variable = value
   End If
End Sub

Public Function AllValuesString() As String
    Dim stringBuilder As String: stringBuilder = ""
    Dim aValue As Variant
    For Each aValue In Items()
        stringBuilder = CStr(aValue) & "," & stringBuilder
    Next
    AllValuesString = Left(stringBuilder, Len(stringBuilder) - 1)
End Function

Public Sub DebugPrint()
   Dim lItem As Long, lIndex As Long, vItem As Variant, oKVP As KeyValuePair
   lItem = 0
   For Each oKVP In KeyValuePairs
      lItem = lItem + 1
      Debug.Print lItem; oKVP.key; " "; TypeName(oKVP.value);
      If InStr(1, TypeName(oKVP.value), "()") > 0 Then
         vItem = oKVP.value
         Debug.Print "("; CStr(LBound(vItem)); " to "; CStr(UBound(vItem)); ")";
         For lIndex = LBound(vItem) To UBound(vItem)
            Debug.Print " (" & CStr(lIndex) & ")"; TypeName(vItem(lIndex)); "="; vItem(lIndex);
         Next
         Debug.Print
      Else
         Debug.Print "="; oKVP.value
      End If
   Next
End Sub
