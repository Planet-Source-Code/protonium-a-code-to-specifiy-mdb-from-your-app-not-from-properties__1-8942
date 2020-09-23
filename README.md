<div align="center">

## a code to specifiy \.MDB from your app not from properties


</div>

### Description

this is a code to specifiy .MDB from your app not from properties it is just for a beginner...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[protoNium](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/protonium.md)
**Level**          |Beginner
**User Rating**    |3.1 (50 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/protonium-a-code-to-specifiy-mdb-from-your-app-not-from-properties__1-8942/archive/master.zip)





### Source Code

```
'Firstly put a Data control in your form and
'add this code to the form load or anywhere that
'is suitable
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\alamat.mdb"
End Sub
```

