Attribute VB_Name = "Logging"
Sub addLog(log As String)
  '' Convenience method to avoid having to CallByName everywhere.
  CallByName frmLog, "addLog", VbMethod, log
End Sub
