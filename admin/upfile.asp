<!--#include FILE="upload.inc"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''建立上传对象
formPath="../../UpLoadPic"
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>35000 then
 	response.write "<font size=2>图片大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 file1="../../UpLoadPic/"
 file2="../UpLoadPic/"
 fn=file.FileName
 'filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 filename0=file1&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 filename2=file2&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt 
 'filename1=replace(replace(filename0, CHR(32), ""), CHR(9), "")
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(filename0)   ''保存文件
  response.write "<script>parent.document.forms[0].upfile.value='"&FileName2&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象

session("upface")="done"

Htmend iCount&" 个文件上传结束!"

sub HtmEnd(Msg)
 set upload=nothing
 response.write "<font size=2>图片上传成功</font>"
 response.end
end sub


%>
</body>
</html>