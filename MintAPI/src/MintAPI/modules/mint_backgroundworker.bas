Attribute VB_Name = "mint_backgroundworker"
Option Explicit

Private p_Thread As Thread


Public Sub Construct()
    p_Thread = Thread.Create(Method.FromReference("WorkerMethod", AddressOf mint_worker_thread, Prototype.VoidMethod))
    p_Thread.Name = "MintWorkerThread"
    p_Thread.IsBackground = True
    
    'Call p_Thread.Start
End Sub


Private Sub mint_worker_thread()
    
End Sub
