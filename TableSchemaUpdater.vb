Option Compare Database
Option Explicit

''' <summary>
''' Used to update table schemas in Access database en masse
''' </summary>
Private Sub AlterTableSchema()

    Dim fld As Field
    Dim dbs As Database, tdf As TableDef
    Set dbs = CurrentDb
    
    For Each tdf In dbs.TableDefs
        If Left(tdf.Name, 4) <> "MSys" Then
            'For Each fld In tdf.Fields
                'If fld.Name = "Archived" Or fld.Name = "Archive" Then
                    'DoCmd.RunSQL ("ALTER TABLE " & tdf.Name & " DROP Column " & fld.Name & ";")
                    DoCmd.RunSQL ("ALTER TABLE " & tdf.Name & " ADD COLUMN LastAmendedByUserName TEXT(255);")
                    'Debug.Print "Table Name: " & tdf.Name & ", Fieldname was: " & fld.Name
'                    fld.Name = "Id"
                'End If
            'Next
        End If
    Next
    
    dbs.Close
    
End Sub
