<%@ page contentType="text/html;charset=utf-8" %>
<%@ page language="java" import="java.io.*,java.sql.*,java.util.*" %>
<%@ page language="java" import="org.apache.commons.fileupload.*,org.apache.commons.fileupload.disk.*,org.apache.commons.fileupload.servlet.*"%>
<%@ include file="connectionInfo.jsp"%>
<%@ include file="FilePathInfo.jsp"%>
<%!
/*------------------------------------------------------------
officeFileItem:文档FileItem类
attachFileItem:附件的FileItem类
officefileNameDisk:文档保存到磁盘上的路径
attachFileNameDisk:附件保存到磁盘上的路径
--------------------------------------------------------------*/
public Connection conn = null;
public Statement stmt = null;
public PreparedStatement pstmt = null;
public ResultSet rs = null;
public FileItem officeFileItem =null ;
public FileItem attachFileItem =null ;
public String officefileNameDisk = "";
public String attachFileNameDisk ="";

/*------------------------------------------------------------
在新建文档的时候调用，返回当前文档在数据库中应该对应的id号。
--------------------------------------------------------------*/
public int GetMaxID()
{
	int mResult=0;
	String mSql=new String();
	mSql = "select max(id)+1 as MaxID from "+officeFileInfoTableName;
	try
	{
		rs = stmt.executeQuery(mSql); 
		if(rs.next())
		{
			mResult=rs.getInt("MaxID");
		}
		rs.close();
		if(mResult==0) mResult=1;
		mSql = "insert into "+officeFileInfoTableName+" (id) values ("+mResult+")" ;
		stmt.execute(mSql);
	}
	catch(SQLException e)
	{
		System.out.println("error:"+e.getMessage());
		mResult=0;
	}
	return (mResult);
}
/*------------------------------------------------------------
保存文档到服务器磁盘，返回值true，保存成功，返回值为false时，保存失败。
--------------------------------------------------------------*/
public boolean saveFileToDisk()
{
	File officeFileUpload = null;
	File attachFileUpload = null;
	String officeFileUploadPath ="";
	String attachFileUploadPath ="";
	boolean result=true ;

	try
	{
		System.out.println("officefilepath:"+absoluteOfficeFileDir+officefileNameDisk);
		System.out.println("attachfilepath:"+absoluteAttachFileDir+attachFileNameDisk);
		if(!officefileNameDisk.equalsIgnoreCase("")&&officeFileItem!=null)
		{
			officeFileUpload =  new File(absoluteOfficeFileDir+officefileNameDisk);
			officeFileItem.write(officeFileUpload);
		}
		if(!attachFileNameDisk.equalsIgnoreCase("")&&attachFileItem!=null)
		{
			attachFileUpload =  new File(absoluteAttachFileDir+attachFileNameDisk);
			attachFileItem.write(attachFileUpload);	
		}
	}
	catch(FileNotFoundException e){}
	catch(Exception e)
	{
		System.out.println("error saveFileToDisk:"+e.getMessage());
		e.printStackTrace();
		result=false;
		}
	return result;	
}
/*------------------------------------------------------------
删除服务器上原有的文件，返回值true，删除成功，false，删除失败。
--------------------------------------------------------------*/
public boolean deleteDiskFile(int deleteFileId)
{
	String sqlStr = "select * from "+officeFileInfoTableName+" where id="+deleteFileId ;
	String officeFileNameDel= "" ;
	String attachFileNameDel= "" ;
	File officeFileDel =null;
	File attachFileDel = null;
	boolean result = true ;
	try
	{
		rs = stmt.executeQuery(sqlStr); 
		if(rs.next())
		{
			officeFileNameDel=rs.getString("FILENAMEDISK")==null?"":rs.getString("FILENAMEDISK").trim();
			attachFileNameDel=rs.getString("ATTACHFILENAMEDISK")==null?"":rs.getString("ATTACHFILENAMEDISK").trim();
			if(!officeFileNameDel.equalsIgnoreCase(""))
			{
				officeFileDel = new File(absoluteOfficeFileDir+officeFileNameDel);
				if(officeFileDel.exists())
				{
					System.out.println("officefilefound!officefilefound!");
					if(!officeFileDel.delete())
					{result=false;}
				}
			}
			if(!attachFileNameDel.equalsIgnoreCase("")&&attachFileItem!=null)
			{
				attachFileDel = new File(absoluteAttachFileDir+attachFileNameDel);
				if(attachFileDel.exists())
				{
					System.out.println("attachfilefound!attachfilefound!");
					if(!attachFileDel.delete())
					{result=false;}
				}	
			}
		}
		rs.close();
}
	catch(Exception e)
	{
		System.out.println("zl error:"+e.getMessage());
		result=false;
	}
	return result;
}
%>
<%
	int fileId = 0 ;	
	long fileSize = 0;
	String fileName = "" ;
	String otherData ="";
	String attachFileDescribe ="" ;
	String attachFileName = "" ;
	String fileType ="";
	String mySqlStr = "";
	String result = "";
	attachFileNameDisk = "";
	boolean isNewRecode = true ;
	officeFileItem =null ;
	attachFileItem =null ;
	DiskFileItemFactory factory = new DiskFileItemFactory();
	// 设置最多只允许在内存中存储的数据,单位:字节
	factory.setSizeThreshold(4096);
	// 设置一旦文件大小超过setSizeThreshold()的值时数据存放在硬盘的目录
	factory.setRepository(new File(tempFileDir));
	ServletFileUpload upload = new ServletFileUpload(factory);
	//设置允许用户上传文件大小,单位:字节
	upload.setSizeMax(1024*1024*4);
	List fileItems = null;
	try{fileItems=upload.parseRequest(request);}
	catch(FileUploadException e)
	{
		out.println("the max upload size is 4m,cheeck upload file size!");
		out.println(e.getMessage());
		e.printStackTrace();
		return;
	}
	Iterator iter = fileItems.iterator();
	attachFileItem=null;
	while (iter.hasNext()) 
	{
		FileItem item = (FileItem) iter.next();
		//打印提交的文本域和文件域名称
		//out.println(item.getFieldName());
		if(item.isFormField())
		{
			if(item.getFieldName().equalsIgnoreCase("fileId"))
			{   
				//if fileId=0,the file is new recode
				try
				{fileId=Integer.parseInt(item.getString().trim());}
				catch(NumberFormatException e)
				{
					System.out.println("NumberFormatException:"+e.getMessage());
					fileId=0;
				}
			}
			if(item.getFieldName().equalsIgnoreCase("fileName"))
			{
				fileName=item.getString("utf-8").trim();
			}
			if(item.getFieldName().equalsIgnoreCase("otherData"))
			{
				otherData=item.getString("utf-8").trim();
				otherData=otherData.equalsIgnoreCase("")?"请输入附加数据":otherData;
			}
			if(item.getFieldName().equalsIgnoreCase("attachFileDescribe"))
			{
				attachFileDescribe=item.getString("utf-8").trim();
				attachFileDescribe=attachFileDescribe.equalsIgnoreCase("")?"请输入附件描述":attachFileDescribe;
			}	
			if(item.getFieldName().equalsIgnoreCase("fileType"))
			{
				fileType=item.getString("utf-8").trim();
			}	
		}else
		{
			if(item.getFieldName().equalsIgnoreCase("attachFile"))
			{
				attachFileItem = item ;
			}
			if(item.getFieldName().equalsIgnoreCase("upLoadFile"))
			{officeFileItem=item;}
		}
	}
	try
	{
		Class.forName(DBDriver);
	}catch(ClassNotFoundException e)
	{out.println("error"+e.getMessage());return;}
	try
	{
		conn = DriverManager.getConnection(ConnStr,userName,userPasswd);    
		stmt=conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
	}catch(SQLException e)
	{out.println("error:"+e.getMessage());return;}
	if(fileId==0) 
	{fileId = GetMaxID();}
	if(fileId!=0&&!fileName.equalsIgnoreCase("")&&officeFileItem!=null)
	{
		if(deleteDiskFile(fileId))//删除数据库中fileid对应的服务器文件
		{
			if(attachFileItem!=null)//如果上传附件为空,应该使用数据库中原有的附件数据
			{
					attachFileNameDisk = attachFileItem.getName();
					attachFileNameDisk = attachFileNameDisk.substring(attachFileNameDisk.lastIndexOf("\\")+1);
					//保存到磁盘中的附件为了防止重复,名字取为 fileId+".attacfile."+attachFileName
					attachFileNameDisk = fileId+".attacfile."+attachFileNameDisk;
			}
			else
			{
					rs = stmt.executeQuery("select * from "+officeFileInfoTableName + " where id="+fileId); 
					rs.next();	
					attachFileNameDisk = rs.getString("ATTACHFILENAMEDISK")==null?"":rs.getString("ATTACHFILENAMEDISK").toString();
					rs.close();
			}
			mySqlStr = "update "+officeFileInfoTableName+" set filename=?,filesize=?,otherdata=?,filetype=?,filenamedisk=?,attachfilenamedisk=?,attachfiledescribe=? where id="+fileId;
			fileSize = officeFileItem.getSize();
			//保存到磁盘中的文档为了防止重复,名字取为 fileId+".officefile."+attachFileName
			officefileNameDisk = fileId+".officefile."+fileName;
			//out.println(officefileUrl+"<br>"+attachFileUrl+"<br>");
			if(saveFileToDisk())
			{
				try
				{
					pstmt = conn.prepareStatement(mySqlStr);
					pstmt.setString(1,fileName);
					pstmt.setInt(2,(int)fileSize);
					pstmt.setString(3,otherData);
					pstmt.setString(4,fileType);
					pstmt.setString(5,officefileNameDisk);
					pstmt.setString(6,attachFileNameDisk);
					pstmt.setString(7,attachFileDescribe);
					pstmt.execute();
					result="文档保存成功。";
				}catch(SQLException e)
				{result="SQLException,updatedatabase:"+e.getMessage();}
			}
			else
			{result="save file failed,please check upload file size,the max size is 4M";}
		}
		else
		{result="delete file failed";}
	}
	else{result="wrong information";}
	out.println(result);
	try
	{
		if(pstmt!=null)pstmt.close(); 
		if(stmt!=null)stmt.close(); 
		if(conn!=null)conn.close(); 
	}
	catch(Exception Error)
	{
		out.println(result+"<br>")	;
		out.println("数据库资源清理失败.");
	}
/*	conn = DriverManager.getConnection(ConnStr,userName,userPasswd);    
	stmt=conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
	rs = stmt.executeQuery("select * from "+officeFileInfoTableName+" fileid="+fileId); 
*/

%>