# rjcToolsResource
A simple CSV and add-in to store and manage resources used by RJC's Excel and Access Tools

## Release Location
`R:\Technical Resources\RJC Sandbox\STG Testing\rjcToolsResource\release`

## How To Use
### Excel
1. Add a `Resource` in your Excel VBA with the following code:

   ```VBA
   Function GetResourcePath(toolID As String, version As String, resourceType As String)
      Application.DisplayAlerts = False
      
      Dim resourceCodePath As String
      Dim func As String
      resourceCodePath = "'R:\Technical Resources\RJC Sandbox\STG Testing\rjcToolsResource\release\RjcToolsResource.xlam'!"
      func = "GetResourcePath"
      
      GetResourcePath = Application.Run(resourceCodePath & func, toolID, version, resourceType)
      
      On Error Resume Next
      Pa.Close False 'Network only
      On Error GoTo 0
   Exit Function
 
   Err:
      MsgBox "Unable to retrieve resource path. Please contact RJC Structural Technical Group.", vbCritical
      Application.DisplayAlerts = True
   End Function
   ```
   
2. Set Public variable(s) for the resource paths in your `Resource` module. For example:
   ```VBA
   Public CodeFile As String
   ```

3. Add a sub to the `Resource` module to set the resources
   ```VBA
   Public Sub SetResourcePaths()
     Dim toolID As String
     Dim version As String
     
     toolID = "WDRIFT"
     version = "2.2.0"
     
     CodeFile = "'" & GetResourcePath(toolID, version, "Backend") & "'!"
   End Sub
   ```

 4. Call `SetResourcePaths` in `Workbook_Open` event
