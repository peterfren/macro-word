[Get started with vba office](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office)

[Object Model list](https://docs.microsoft.com/en-us/office/vba/api/overview/word/object-model)

# Dasar
* open file: ```visualbasic  Documents.Open FileName:="C:\MyFiles\MyDoc.doc", ReadOnly:=True ```
* Save file: ```visualbasic Application.ActiveDocument.Save ```

* seleksi / text blocking : 
```visualbasic     
    Application.Selection.Value = "Hello World"
```
* get information of what you select:
```visualbasic  
Selection.Information()
```

* Dialog file ms word  [Application.FileDialog property](https://docs.microsoft.com/en-us/office/vba/api/word.application.filedialog)


