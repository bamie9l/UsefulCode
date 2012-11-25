Option Compare Database
Option Explicit

''' <summary>
''' Used to update table schemas in Access database en masse
''' </summary>
Private Sub AlterTableSchema()

    Dim dbs As Database, tdf As TableDef
    Set dbs = CurrentDb
    
    For Each tdf In dbs.TableDefs
        For Each fld In tdf.Fields
            If fld.Name = "LastAmendedWhen" Then
                'DoCmd.RunSQL ("ALTER TABLE " & tdf.Name & " DROP Column " & fld.Name & ";")
                Debug.Print "Table Name: " & tdf.Name & ", Fieldname was: " & fld.Name
                fld.Name = "LastAmendedDate"
            End If
        Next
    Next
    
    dbs.Close
    
End Sub