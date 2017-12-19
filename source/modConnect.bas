Attribute VB_Name = "modConnect"
Option Explicit

Function GetCString() As String
    Dim db As New Database
        GetCString = db.CStringBuilder("\\Njes1s3029\nh-dmo\8245-QRC-CDR\OMU\Repository\Super_Portal\Portal_DEV.mdb", "portal2018")
End Function


