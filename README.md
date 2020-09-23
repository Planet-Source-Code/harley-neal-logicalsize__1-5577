<div align="center">

## LogicalSize


</div>

### Description

Resize and Center an image control(maintaining image proportion)(remember to load an image)

inside a picturebox control. This code rescales and centers the image to a size small enough

to fit inside any give picture box. Good for thumbnails. I don't know if

this code is bug proof... Let me know what you think ,Thanks
 
### More Info
 
none??


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Harley Neal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/harley-neal.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/harley-neal-logicalsize__1-5577/archive/master.zip)

### API Declarations

```
'Put this code in a Regular module: declare this..
Public SmartSize= new class1 ' SmartSize can be any name
```


### Source Code

```
'Start a new project
'Add a new module and class module to your project
'Add a picture box (with an Image control inside of it)to your form.
'load an image into the image control
'Put this code in the standard module: declare this..
Public SmartSize= new class1 ' SmartSize can be any name class1 the name of the module
'paste the code below to the class module
'the cushion variable will space the image away from the picture edge.
Public Sub LogicalSize(ContainerObj As Object, ImgObj As Object, ByVal Cushion As Integer)
Dim VertChg, HorzChg As Integer
Dim iRatio As Double
Dim ActualH, ActualW As Integer
Dim ContH, ContW As Integer
On Error GoTo LogicErr
With ImgObj 'hide picture while changing size
 .Visible = False
 .Stretch = False 'set actual size
End With
VertChg = 0: HorzChg = 0
ActualH = ImgObj.Height 'actual picture height
ActualW = ImgObj.Width 'actual picture width
ContH = ContainerObj.Height - Cushion 'set max. picture height
ContW = ContainerObj.Width - Cushion 'set max. picture width
CenterCTL ContainerObj, ImgObj 'center picture
If ImgObj.Top < Cushion Or ImgObj.Left < Cushion Then 'is picture larger than container
 If ActualH <> ActualW Then 'picture is not square
  If ActualH > ActualW Then 'height is greater
   iRatio = (ActualH / ActualW) 'get ratio between height and width
   HorzChg = 10 'scale down by 10 units per loop
   VertChg = CInt(Format(iRatio * 10, "####"))
  Else 'width is greater
   iRatio = (ActualW / ActualH) 'get ratio between height and width
   VertChg = 10 'scale down by 10 units per loop
   HorzChg = CInt(Format(iRatio * 10, "####")) 'round number
  End If
 Else 'picture is square
  VertChg = 10 'scale both height and width equally
  HorzChg = 10
 End If
 Do Until ActualH <= ContH And ActualW <= ContW
  ActualH = ActualH - VertChg 'scale height down
  ActualW = ActualW - HorzChg 'scale width down
  If ActualH < 100 Then
   ActualH = 100 'set min. picture height=100
   Exit Do
  ElseIf ActualW < 100 Then
   ActualW = 100 'set min. picture width=100
   Exit Do
  End If
 Loop
 With ImgObj 'set new height and width
  .Stretch = True
  .Height = ActualH
  .Width = ActualW
 End With
End If
CenterCTL ContainerObj, ImgObj 'center picture in container
ImgObj.Visible = True 'show picture
Exit Sub
LogicErr:
MsgBox "An Error occured while rescaling this image. Image size maybe invalid.", vbSystemModal + vbExclamation, "Resize Error!"
End Sub
Public Sub CenterCTL(FRMObj As Object, OBJ As Control)
With OBJ
 .Top = (FRMObj.Height / 2) - (OBJ.Height / 2)
 .Left = (FRMObj.Width / 2) - (OBJ.Width / 2)
 .ZOrder
End With
End Sub
'Call the Logical Size method like this
'put this code anywhere, in button click, image click whereever you want
SmartSize.LogicalSize Picture1, Image1, 100
```

