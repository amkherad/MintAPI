Attribute VB_Name = "modMain"
Option Explicit

Public tApplication As IApplication

Public Sub Main()
    Set tApplication = Application.InitializeNewApplication( _
        CurrentProcess, "com.mintapi.materialloader", App.Title, _
        App.ProductName, App.Path, Directory.ConcatPath(App.Path, App.EXEName), _
        App.EXEName, "", "", "", "", "", "", "", App.CompanyName, _
        App.Major, App.Minor, App.Revision, "", App.LegalCopyright, _
        "", "", "", 0, App, _
        "", "", "", "", "")
        
End Sub
