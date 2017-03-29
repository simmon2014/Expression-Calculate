// Input.cpp : 实现文件
//

#include "stdafx.h"
#include "Cal.h"
#include "Input.h"
#include "MyWord10.h"


// CInput

CInput::CInput()
{
}

CInput::~CInput()
{
}


// CInput 成员函数
int CInput::ReadFile(CString path)
{
	CFile file;
	long nFileLength;
	long realLength;
	char * buffer;
	CString strText;
	if(file.Open(path,CFile::modeRead))
	{
		nFileLength=(long)file.GetLength();
		buffer=new char[nFileLength+1];
		realLength=file.Read(buffer,nFileLength);
		strText=CString(buffer,realLength);
		file.Close();
		delete []buffer;
	}
	else
	{
		return 0;
	}
 return	this->ReadStr(strText);
}



//int CInput::ReadStr(CString text)
//{
//    int status;
//	CString str;
//	int start,end;
//	CString startvar,endvar,startValue,endValue,startRel,endRel;
//	startvar="[Var]";endvar="[EndVar]";startValue="[Value]";endValue="[EndValue]";startRel="[Rel]";endRel="[EndRel]";
//	
//	start=text.Find(startvar);if(start==-1) return 0;
//	end=text.Find(endvar);if(end==-1) return 0;
//    str=text.Mid(start+startvar.GetLength(),end-start-startvar.GetLength());
//	status=this->ParseVar(str);
//    if(status==0) return 0;
//
//
//	start=text.Find(startValue);if(start==-1) return 0;
//	end=text.Find(endValue);if(end==-1) return 0;
//	str=text.Mid(start+startValue.GetLength(),end-start-startValue.GetLength());
//	status=this->ParseValue(str);
//	if(status==0) return 0;
//
//	start=text.Find(startRel);if(start==-1) return 0;
//	end=text.Find(endRel);if(end==-1) return 0;
//	str=text.Mid(start+startRel.GetLength(),end-start-startRel.GetLength());
//	status=this->ParseRel(str);
//	if(status==0) return 0;
//
//	return 1;
//}
//
//
//int CInput::ParseVar(CString text)
//{
//	/////////////读每一行，加入数组
//	CString str=text;
//	CStringArray strArray;
//	int curPos= 0;
//	CString blank="\n";
//	CString resToken;
//	resToken=str.Tokenize(blank,curPos);
//	while (resToken != "")
//	{
//		strArray.Add(resToken);
//		resToken= str.Tokenize(blank,curPos);
//	};
//
//	int index;
//	CString name,note;
//	Variant var;
//	var.Value=0;
//	
//	var.state=0;
//	var.Name="";
//	var.Note="";
//	m_VarArray.RemoveAll();
//	for(int i=0;i<strArray.GetSize();i++)
//	{
//		str=strArray[i];
//        index=str.Find(";");
//		if(index!=-1)
//		{
//			name=str.Mid(0,index);
//			note=str.Mid(index+1);
//			var.Name=name.Trim();
//			var.Note=note.Trim();
//		}
//		else
//		{
//			if(str.Trim()=="") continue;
//			var.Name=str.Trim();
//			var.Note="";
//		}
//		m_VarArray.Add(var);
//	}
//	return 1;
//}
//
//
//int CInput::ParseValue(CString text)
//{
//	/////////////读每一行，加入数组
//	CString str=text;
//	CStringArray strArray;
//	int curPos= 0;
//	CString blank=";";
//	CString resToken;
//	resToken=str.Tokenize(blank,curPos);
//	while (resToken != "")
//	{
//		strArray.Add(resToken);
//		resToken= str.Tokenize(blank,curPos);
//	};
//
//	int index;
//	CString name,strValue;
//
//	for(int i=0;i<strArray.GetSize();i++)
//	{
//		str=strArray[i];
//		index=str.Find("=");
//		if(index!=-1)
//		{
//			name=str.Mid(0,index);
//			strValue=str.Mid(index+1);
//			name=name.Trim();
//			strValue=strValue.Trim();
//			for(int j=0;j<m_VarArray.GetSize();j++)
//			{
//				if(m_VarArray[j].Name==name)
//				{
//					m_VarArray[j].strValue=strValue;
//					m_VarArray[j].Value=atof(strValue);
//					m_VarArray[j].state=1;
//				}
//			}
//
//			
//		}
//	
//	}
//	return 1;
//}
//
//
//int CInput::ParseRel(CString text)
//{
//	/////////////读每一行，加入数组
//	CString str=text;
//	CStringArray strArray;
//	int curPos= 0;
//	CString blank=";";
//	CString resToken;
//	resToken=str.Tokenize(blank,curPos);
//	while (resToken != "")
//	{
//		strArray.Add(resToken);
//		resToken= str.Tokenize(blank,curPos);
//	};
//
//	m_Rel.RemoveAll();
//	m_SourceRel.RemoveAll();
//    for(int i=0;i<strArray.GetSize();i++)
//	{
//		str=strArray[i];
//		str=str.Trim();
//		if(str=="") continue;
//		m_SourceRel.Add(str);
//		m_Rel.Add(str);
//	}
//	return 1;
//}






int CInput::Check()
{
	int status;
	CStringArray import,export;
	CString str;
	for(int i=0;i<m_VarArray.GetSize();i++)
	{
		//if(m_VarArray[i].state==1)
			import.Add(m_VarArray[i].Name);
		//if(m_VarArray[i].state==0)
		//	export.Add(m_VarArray[i].Name);
	}

	for(int i=0;i<m_SourceRel.GetSize();i++)
	{
		str=m_SourceRel[i];
		status=util.CheckEquation(str,&import,&export);
		if(status==0) return 0;
	}
	return 1;
}


int CInput::ReadStr(CString text)
{
    int status;
	CString str;
	int start,end;
	CString startvar,endvar,startRel,endRel;
	startvar="[参数开始]";endvar="[参数结束]";startRel="[公式开始]";endRel="[公式结束]";
	
	start=text.Find(startvar);if(start==-1) return 0;
	end=text.Find(endvar);if(end==-1) return 0;
    str=text.Mid(start+startvar.GetLength(),end-start-startvar.GetLength());
	status=this->ParseVar(str);
    if(status==0) return 0;

	start=text.Find(startRel);if(start==-1) return 0;
	end=text.Find(endRel);if(end==-1) return 0;
	str=text.Mid(start+startRel.GetLength(),end-start-startRel.GetLength());
	status=this->ParseRel(str);
	if(status==0) return 0;

	return 1;
}


int CInput::ParseVar(CString text)
{
	int status;
	/////////////读每一行，加入数组
	CString str=text;
	CStringArray strArray;
	int curPos= 0;
	CString blank="\n";
	CString resToken;
	resToken=str.Tokenize(blank,curPos);
	while (resToken != "")
	{
		strArray.Add(resToken);
		resToken= str.Tokenize(blank,curPos);
	};

	int index;
	CString fullname,name,note,strValue;
	Variant var;
	var.Value=0;
	
	var.state=1;
	var.Name="";
	var.Note="";
	m_VarArray.RemoveAll();
	for(int i=0;i<strArray.GetSize();i++)
	{
		str=strArray[i];
        index=str.Find(";");
		if(index!=-1)
		{
			if(index==0) continue;
			fullname=str.Mid(0,index);
			note=str.Mid(index+1);
			status=this->GetPre(fullname,name,strValue,"=");
			if(status!=1) continue;
            if(name.Trim()=="") continue;
			strValue=strValue.Trim();
			var.Name=name.Trim();
			var.Note=note.Trim();
			var.strValue=strValue;
			var.Value=atof(strValue);
			
		}
		else
		{
			
			fullname=str;
			note="";
			status=this->GetPre(fullname,name,strValue,"=");
			if(status!=1) continue;
			if(name.Trim()=="") continue;
			strValue=strValue.Trim();
			var.Name=name.Trim();
			var.Note=note.Trim();
			var.strValue=strValue;
			var.Value=atof(strValue);
		}
		m_VarArray.Add(var);
	}
	return 1;
}


int CInput::GetPre(CString text,CString &pre,CString &next,CString blank)
{
  int index;
  index=text.Find(blank);
  if(index==-1) return -1;
  if(index==0) 
  {
	  pre="";
	  next=text.Mid(index+1);
	  return 0;
  }
  else
  {
	  pre=text.Mid(0,index);
	  next=text.Mid(index+1);
  }
   return 1;
}



int CInput::ParseValue(CString text)
{
	/////////////读每一行，加入数组
	CString str=text;
	CStringArray strArray;
	int curPos= 0;
	CString blank=";";
	CString resToken;
	resToken=str.Tokenize(blank,curPos);
	while (resToken != "")
	{
		strArray.Add(resToken);
		resToken= str.Tokenize(blank,curPos);
	};

	int index;
	CString name,strValue;

	for(int i=0;i<strArray.GetSize();i++)
	{
		str=strArray[i];
		index=str.Find("=");
		if(index!=-1)
		{
			name=str.Mid(0,index);
			strValue=str.Mid(index+1);
			name=name.Trim();
			strValue=strValue.Trim();
			for(int j=0;j<m_VarArray.GetSize();j++)
			{
				if(m_VarArray[j].Name==name)
				{
					m_VarArray[j].strValue=strValue;
					m_VarArray[j].Value=atof(strValue);
					m_VarArray[j].state=1;
				}
			}

			
		}
	
	}
	return 1;
}


int CInput::ParseRel(CString text)
{
	
	
	
	int status;
	/////////////读每一行，加入数组
	CString str=text;
	CStringArray strArray;
	int curPos= 0;
	CString blank="\n";
	CString resToken;
	resToken=str.Tokenize(blank,curPos);
	while (resToken != "")
	{
		strArray.Add(resToken);
		resToken= str.Tokenize(blank,curPos);
	};

	int index;
	CString fullname,name,note,strValue;
	Variant var;
	var.Value=0;

	var.state=0;
	var.Name="";
	var.Note="";

	m_Rel.RemoveAll();
	m_SourceRel.RemoveAll();
	for(int i=0;i<strArray.GetSize();i++)
	{
		str=strArray[i];
		index=str.Find(";");
		if(index!=-1)
		{
			if(index==0) continue;
			fullname=str.Mid(0,index);
			note=str.Mid(index+1);
			status=this->GetPre(fullname,name,strValue,"=");
			if(status!=1) continue;
			if(name.Trim()=="") continue;
			
			var.Name=name.Trim();
			var.Note=note.Trim();
			

		}
		else
		{

			fullname=str;
				note="";
			status=this->GetPre(fullname,name,strValue,"=");
			if(status!=1) continue;
			if(name.Trim()=="") continue;
			strValue=strValue.Trim();
			var.Name=name.Trim();
			var.Note=note.Trim();
			var.strValue=strValue;
			var.Value=atof(strValue);
		}

		m_SourceRel.Add(fullname);
		m_Rel.Add(fullname);
		m_VarArray.Add(var);
	}
	return 1;

}





int CInput::Compute()
{
   int num;
   num=m_Rel.GetSize();
	for(int i=0;i<num;i++)
	{
       CalOnce();
	}
	return 1;
}


int CInput::CalOnce()
{
	CStringArray export;
	for(int i=0;i<m_VarArray.GetSize();i++)
	{
		if(m_VarArray[i].state==0)
			export.Add(m_VarArray[i].Name);
	}

	if(export.GetSize()==0) return 0;

	int num;
	CString X;

    int index=-1;
	for(int i=0;i<m_Rel.GetSize();i++)
	{
		num=util.FindVar(m_Rel[i],export,X);
		if(num==1)
		{
			index=i;
			break;
		}
	}
	if(index==-1) return 0;
	CString equal=m_Rel[index];
	for(int i=0;i<m_VarArray.GetSize();i++)
	{
		if(m_VarArray[i].state==0) continue;
		if(m_VarArray[i].state==1||m_VarArray[i].state==2)
		{
			equal=util.ReplaceXtoDouble(equal,m_VarArray[i].Name,m_VarArray[i].Value);
		}
	}

	double result;
	result=util.ComputeEquation(equal,X);
	m_Rel.RemoveAt(index);

	for(int i=0;i<m_VarArray.GetSize();i++)
	{
		if(m_VarArray[i].Name==X)
		{
			m_VarArray[i].state=2;
			m_VarArray[i].Value=result;
		}
	}

     return 1;
}





int CInput::WriteInput(CString path)
{
  CString str;
  CString var="",value="",rel="";
  CString temp,temp2;
for(int i=0;i<m_VarArray.GetSize();i++)
{
  temp.Format("%s;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Note);
  var=var+temp;
}
 temp.Format("[Var]\r\n");
 temp2.Format("[EndVar]\r\n");
 var=temp+var+temp2;

 for(int i=0;i<m_VarArray.GetSize();i++)
 {
    if(m_VarArray[i].state==1)
	{
	 temp.Format("%s=%f;\r\n",m_VarArray[i].Name,m_VarArray[i].Value);
		 value=value+temp;
	}
 }
 temp.Format("[Value]\r\n");
 temp2.Format("[EndValue]\r\n");
value=temp+value+temp2;


 for(int i=0;i<m_SourceRel.GetSize();i++)
 {
  temp.Format("%s;\r\n",m_SourceRel[i]);
			 rel=rel+temp;
 }
 temp.Format("[Rel]\r\n");
 temp2.Format("[EndRel]\r\n");
 rel=temp+rel+temp2;

 str=var+value+rel;

	CFile file;

	CString strText;
	if(file.Open(path,CFile::modeWrite|CFile::modeCreate|CFile::shareDenyNone))
	{
		file.Write(str.GetBuffer(),str.GetLength());
		file.Close();
	
	}
	else
	{
		return 0;
	}
	return 1;
}



//int CInput::WriteOutput(CString path)
//{
//	CString str;
//	CString value="";
//	CString temp,temp2;
//	
//
//	for(int i=0;i<m_VarArray.GetSize();i++)
//	{
//		if(m_VarArray[i].state==2)
//		{
//			temp.Format("%s=%f;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Value,m_VarArray[i].Note);
//				value=value+temp;
//		}
//	}
//	temp.Format("[Result]\r\n");
//	temp2.Format("[EndResult]\r\n");
//		value=temp+value+temp2;
//
//
//
//	str=value;
//
//	CFile file;
//
//	CString strText;
//	if(file.Open(path,CFile::modeWrite|CFile::modeCreate|CFile::shareDenyNone))
//	{
//		file.Write(str.GetBuffer(),str.GetLength());
//		file.Close();
//
//	}
//	else
//	{
//		return 0;
//	}
//	return 1; 
//
//}





 int CInput::WriteDoc(CString path1,CString path2)
 {
	 CApplication app;
	 COleVariant vTrue((short)TRUE),	vFalse((short)FALSE);
	 app.CreateDispatch(_T("Word.Application"));
	 app.put_Visible(FALSE);

	 CDocuments docs=app.get_Documents();
	 //下面是定义VARIANT变量；
	 COleVariant varFilePath(path1.GetBuffer());
	 COleVariant varstrNull("");
	 COleVariant varZero((short)0);
	 COleVariant varTrue(short(1),VT_BOOL);
	 COleVariant varFalse(short(0),VT_BOOL);
	 docs.Open2000(varFilePath,varFalse,varFalse,varFalse,
		 varstrNull,varstrNull,varFalse,varstrNull,
		 varstrNull,varTrue,varTrue,varTrue);
	 CDocument0 mydoc=app.get_ActiveDocument();
	 CSelection sel=app.get_Selection();
	
	 
	 CFind object=sel.get_Find();

	 COleVariant Warp(long(1));
	 COleVariant Replace(long(2));
	 CString findText,replaceText;

	 COleVariant t1;
	 COleVariant t2;

	 for(int i=0;i<m_VarArray.GetSize();i++)
	 {
		 findText.Format("$%s$",m_VarArray[i].Name);
		 replaceText.Format("%0.4f",m_VarArray[i].Value);
	 t1=COleVariant(findText.GetBuffer());
     t2= COleVariant(replaceText.GetBuffer());
		 object.Execute(t1,vFalse,vFalse,vFalse,vFalse,vFalse,vTrue,Warp,vFalse,t2,Replace,vFalse,vFalse,vFalse,vFalse);
	 }

	 

	 //  app.put_Visible(TRUE);
     app.put_Visible(TRUE);

	 mydoc.SaveAs (COleVariant(path2.GetBuffer()),vFalse,vFalse, COleVariant(""),vFalse,
		 COleVariant(""),vFalse,vFalse,vFalse,vFalse,vFalse,vFalse,vFalse,vFalse,vFalse,vFalse);

	 object.ReleaseDispatch();
	 sel.ReleaseDispatch();
	 mydoc.ReleaseDispatch();
	 docs.ReleaseDispatch();
	 app.ReleaseDispatch();
	 return 1;
 }



 /************************************************************************/
 /*                                                                      */
 /************************************************************************/




 int CInput::WriteOutput(CString path)
 {
	 CString str;
	 CString var="",value="",rel="";
	 CString temp,temp2;
	

	 for(int i=0;i<m_VarArray.GetSize();i++)
	 {
		 if(m_VarArray[i].state==1)
		 {
			 temp.Format("%s=%f;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Value,m_VarArray[i].Note);
			 value=value+temp;
		 }
	 }
	 temp.Format("[已知参数开始]\r\n");
	 temp2.Format("[已知参数结束]\r\n");
	 value=temp+value+temp2;

	 for(int i=0;i<m_VarArray.GetSize();i++)
	 {
		 if(m_VarArray[i].state==2)
		 {
			 temp.Format("%s=%f;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Value,m_VarArray[i].Note);
			 var=var+temp;
		 }
	 }
	 temp.Format("[结果]\r\n");
	 temp2.Format("[结果]\r\n");
	 var=temp+var+temp2;
	

	 str=value+var;

	 CFile file;

	 CString strText;
	 if(file.Open(path,CFile::modeWrite|CFile::modeCreate|CFile::shareDenyNone))
	 {
		 file.Write(str.GetBuffer(),str.GetLength());
		 file.Close();

	 }
	 else
	 {
		 return 0;
	 }
	 return 1;
 }


 int CInput::CalExEqual(CString input,CString & output)
 {
	 CString strText;
	 int status;
	 strText=input;
	 status=this->ReadStr(strText);if(status==0) return 0;
	 status=this->Check();
	 if(status==0) 
	 {
		 AfxMessageBox("输入格式错误，计算失败!");
		 return 0;
	 }

	 int num;
	 num=m_Rel.GetSize();
	 for(int i=0;i<num;i++)
	 {
		 CalOnce();
	 }
	
	 status=this->WriteOutputStr(output);
	 return 1;
 }

 int CInput::WriteOutputStr(CString & FullText)
 {
	 CString str;
	 CString var="",value="",rel="";
	 CString temp,temp2;


	 for(int i=0;i<m_VarArray.GetSize();i++)
	 {
		 if(m_VarArray[i].state==1)
		 {
			 temp.Format("%s=%f;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Value,m_VarArray[i].Note);
			 value=value+temp;
		 }
	 }
	 temp.Format("[已知参数开始]\r\n");
	 temp2.Format("[已知参数结束]\r\n");
	 value=temp+value+temp2;

	 for(int i=0;i<m_VarArray.GetSize();i++)
	 {
		 if(m_VarArray[i].state==2)
		 {
			 temp.Format("%s=%f;%s\r\n",m_VarArray[i].Name,m_VarArray[i].Value,m_VarArray[i].Note);
			 var=var+temp;
		 }
	 }
	 temp.Format("[结果]\r\n");
	 temp2.Format("[结果]\r\n");
	 var=temp+var+temp2;


	 str=value+var;

	 FullText=str;
	 return 1;
 }
