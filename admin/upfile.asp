<!--#include FILE="upload.inc"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����
formPath="../../UpLoadPic"
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>35000 then
 	response.write "<font size=2>ͼƬ��С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 file1="../../UpLoadPic/"
 file2="../UpLoadPic/"
 fn=file.FileName
 'filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 filename0=file1&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 filename2=file2&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt 
 'filename1=replace(replace(filename0, CHR(32), ""), CHR(9), "")
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(filename0)   ''�����ļ�
  response.write "<script>parent.document.forms[0].upfile.value='"&FileName2&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���

session("upface")="done"

Htmend iCount&" ���ļ��ϴ�����!"

sub HtmEnd(Msg)
 set upload=nothing
 response.write "<font size=2>ͼƬ�ϴ��ɹ�</font>"
 response.end
end sub


%>
</body>
</html>