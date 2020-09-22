<div align="center">

## TestEmailWithManyManyAttachments


</div>

### Description

You want to send 10 or more files using your VB program but you always get an error, well not any more, here is the solution. Enjoy!
 
### More Info
 
MAPISession control,   MAPIMessages control

I dont think so!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demetrios Pyrenis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demetrios-pyrenis.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demetrios-pyrenis-testemailwithmanymanyattachments__1-4512/archive/master.zip)





### Source Code

```
'
'
'
' mapSess = MAPISession Control
' mapMess = MAPIMessages Control
'
'
private sub TestEmailWithManyManyAttachments()
dim Attachments() as string
dim TotAttachments as long
dim i as long
dim attPos as integer
  TotAttachments=2 ' or more
  Redim Attachments(TotAttachments)
  Attachments(1)="c:\config.sys"
  Attachments(2)="c:\autoexec.bat"
  mapSess.LogonUI = True
  mapSess.SignOn
  mapMess.SessionID = mapSess.SessionID
  mapMess.Compose
  mapMess.MsgSubject = "Some Subject"
  mapMess.MsgNoteText = "  bla bla bla bla bla"
  attPos = 1
  For i = 1 To TotAttachments
    If Dir( Attachments(i) ) <> "" Then ' Chek that file exists
      mapMess.AttachmentIndex = i - 1
      mapMess.AttachmentPosition = attPos
      mapMess.AttachmentPathName = Attachments(i)
      attPos = attPos + 1
    End If
  Next i
  DoEvents
  mapMess.Send True
  DoEvents
  mapSess.SignOff
end sub
```

