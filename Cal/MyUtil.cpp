// MyUtil.cpp : ʵ���ļ�
//
#pragma  once
#include "stdafx.h"
#include "MyUtil.h"
#define _USE_MATH_DEFINES
#include <math.h>


// CMyUtil

CMyUtil::CMyUtil()
{
	
}

CMyUtil::~CMyUtil()
{
	
}


// CMyUtil ��Ա����
//����·����solidָ���solid�ļ�


CString CMyUtil::ReplaceString(CString str,CString str1,CString str2)
{
	CString ss;
	ss=str1+":::";
	str.Replace(str1,ss);
	str.Replace(str2,str1);
	str.Replace(ss,str2);
	return str;
}

  BOOL CMyUtil::IsDigit(int c)
  {
	  if(c>='0'&&c<='9')
		  return TRUE;
	  else
		  return FALSE;
  }

BOOL CMyUtil::IsDouble(CString str)
{
	CString left,right;
	int c;
	int i;
	int nIndex=str.Find(".");
   if(nIndex!=-1)
   {
	   right=str.Mid(nIndex+1);
	   left=str.Left(nIndex);
   }
   else
   {
		right="12";
		left=str;
   }
   c=left.GetAt(0);
   if(left.GetLength()==1)
   {
	   if(!IsDigit(c)) return FALSE;
   }
   else
   {
	   if(!IsDigit(c)||c=='0') return FALSE;
	   for( i=1;i<left.GetLength();i++)
	   {
		   c=left.GetAt(i);
		   if(!IsDigit(c)) return FALSE;
	   }
   }

   for(i=0;i<right.GetLength();i++)
   {
		 c=right.GetAt(i);	  
	   if(!IsDigit(c)) return FALSE;
   }
	return TRUE;
}

BOOL CMyUtil::IsFunction(CString str)
{
	CStringArray funArray;
	funArray.Add("sin");
	funArray.Add("cos");
	for(int i=0;i<funArray.GetSize();i++)
	{
		if(str==funArray[i])
			return TRUE;
	}
	return FALSE;
}

/************************************************************************/
/* ������ʽ 
	�����ҵ�����,�����Լ�����������ı��ʽ,�ж��Ƿ��Ǻ���
	�������滻��ĳһʵ��.���´�����ʽ.
	��ͷ����,���ȿ��Ƿ�ʵ��.
	����ʵ��,������Array��,�ҵ�ƥ����,��������������Array��;
	���򷵻س���.
	��������ͬ�������ȥ��,���´�����ʽ.

*/
/************************************************************************/
BOOL CMyUtil::CheckExpression(CString Expression,CStringArray * importDim,CStringArray * exportDim)
{
	int rIndex,min;
	CString str,sub;
	int i;
	str=Expression;
	min=str.Find("(");
	int n=0;
	if(min!=-1)
{
	
	for(i=min;i<str.GetLength();i++)
	{
			char ch=str.GetAt(i);
				if(ch=='(')
				{
					n++;
				}
				else if(ch==')')
				{
					n--;
					if(n==0)
					{
						rIndex=i;
						break;
					}
				}
		}
	if(i==str.GetLength())
	{
		str="���Ų�ƥ��";
		AfxMessageBox(str);
		return FALSE;
	}
	sub=str.Mid(min+1,rIndex-min+1-2);///����������ַ���
	if(CheckExpression(sub,importDim,exportDim)==FALSE) return FALSE;
    sub=str.Mid(min,rIndex-min+1);///�������ŵ��ַ���
	///��������,����
	CString str1,str3;
	str1=str.Mid(0,min);///���ǰ��α��ʽ,����������.
	int a1,a2;
	a1=str1.ReverseFind('+');
	a2=str1.ReverseFind('-');
	if(a2>a1) a1=a2;
	a2=str1.ReverseFind('*');
	if(a2>a1) a1=a2;
	a2=str1.ReverseFind('/');
	if(a2>a1) a1=a2;
	a2=str1.ReverseFind('(');
	if(a2>a1) a1=a2;
	BOOL fff=FALSE;
	if(a1+1<str1.GetLength())
	{
		str3=str1.Mid(a1+1);
		CStringArray  label;
		label.Add("acos");
		label.Add("asin");
		label.Add("atan");
		label.Add("cos");
		//label.Add("cosh");
		label.Add("exp");
		label.Add("fabs");
		label.Add("log");
		label.Add("log10");
		label.Add("sin");
	//	label.Add("sinh");
		label.Add("tan");
	//	label.Add("tanh");
		label.Add("sqrt");


		label.Add("pow");
		label.Add("max");
		label.Add("min");
	
	for(i=0;i<label.GetSize();i++)
	{
		if(str3==label.GetAt(i))
		{
			sub=str3+sub;
			fff=TRUE;
		}
	}
	if(fff==FALSE)
	{
		str="����ǰ��Ҫ���������";
		AfxMessageBox(str);
		return FALSE;
	}
}
	////�����������
	str.Replace(sub,"1.23456789");
	return CheckExpression(str,importDim,exportDim);
}
min=0;
//str=Expression.Tokenize("*+/-",min);
str=Expression.Tokenize("*+/-,",min);
if(str.GetLength()!=min-1)
{
	str=""+str+"  ǰ���ص��������";
	str.Replace("1.23456789","����");
	AfxMessageBox(str);
	return FALSE;
}
///�Ƿ�ʵ��
if(IsDouble(str))
{
	if(str==Expression)
	{
		return TRUE;
	}
	else
	{
		if(min==Expression.GetLength()) 
		{
			str="��  "+str+"   ������ִ���,������ź��������������";
			str.Replace("1.23456789","����");
			AfxMessageBox(str);
			return FALSE;
		}
		str=Expression.Mid(min);
		return CheckExpression(str,importDim,exportDim);
	}
}

/////////////�Ƿ�ߴ�������ʶ
BOOL nFlag=FALSE;
BOOL pFlag=TRUE;
for(i=0;i<importDim->GetSize();i++)
{
	if(str==importDim->GetAt(i))
	{
		nFlag=TRUE;
		for(int m=0;m<exportDim->GetSize();m++)
		{
			if(str==exportDim->GetAt(m))
				pFlag=FALSE;
		}
		if(pFlag==TRUE)
		exportDim->Add(str);
	}
}
if(str=="PI"||str=="E")/////�жϳ���
	nFlag=TRUE;
if(nFlag==FALSE) 
{
	min=str.Find("1.23456789");
	if(min!=-1)
		str="�����������,��ƥ��";
	else
		str="�Ҳ���"+str+"��һ��,����";
	AfxMessageBox(str);
	return FALSE;
}

if(str==Expression)
{
	return TRUE;
}
else
{
	if(min==Expression.GetLength()) 
	{
		str="��  "+str+"   ������ִ���,������ź��������������";
		str.Replace("1.23456789","����");
		AfxMessageBox(str);
		return FALSE;
	}
	str=Expression.Mid(min);
	return CheckExpression(str,importDim,exportDim);
}

	return TRUE;
}

/************************************************************************/
/* ��鷽�� 
�����жϵȺ�.
�ڷֱ�����߱��ʽ���ұ߱��ʽ.
*/
/************************************************************************/
BOOL CMyUtil::CheckEquation(CString Expression,CStringArray * importDim,CStringArray * exportDim)
{
	int min;
	CString str,left,right;
	str=Expression;
	min=str.Find("=");
	if(min==-1)
	{
		str="����û�еȺ�";
		AfxMessageBox(str);
		return FALSE;
	}
   
	if(min==0)
	{
		str="���̵Ⱥ�ǰ����������";
		AfxMessageBox(str);
		return FALSE;
	}

	if(min==str.GetLength()-1)
	{
		str="���̵Ⱥź�����������";
		AfxMessageBox(str);
		return FALSE;
	}
	left=str.Left(min);
	right=str.Mid(min+1);
	min=right.Find("=");
	if(min!=-1)
	{
		str="���̵Ⱥ�����������";
		AfxMessageBox(str);
		return FALSE;
	}
	if(CheckExpression(left,importDim,exportDim))
	{
		if(CheckExpression(right,importDim,exportDim))
			return TRUE;
	}
	return FALSE;
	
}

/************************************************************************/
/*/////�������ɰ���sin(),cos(),log(),pow();                                                                      */
/************************************************************************/

double CMyUtil::CalString(CString expression)
{
	int start,end;
	int t;
	CString str;
	CString str1,str2,str3;
	double temp=-1;
	double d1,d2;
	char c;
	try
{/////////////tryǰ����
		///////////////////��������
		start=expression.Find('(');
		if(start!=-1)
		{
			end=expression.Find(')',start);
			if(end!=-1)
		 {
			 str=expression.Mid(0,end);
			 start=str.ReverseFind('(');

			 str=expression.Mid(start+1,end-start+1-2);


			 //////////////////////////////////////////////////////////////////////////
			 int a_index;
			 CString a_a,a_b;
			 a_index=str.Find(",");
			 if(a_index!=-1)
			 {
				 a_a=str.Mid(0,a_index);
				 a_b=str.Mid(a_index+1);
				 a_a=a_a.Trim();
				 a_b=a_b.Trim();

			 }
			 else
			 {
				 temp=CalString(str);
				 str.Format("%f",temp);
				 str.Replace('-','@');
			 }
			
			 str2=expression.Mid(start,end-start+1);

			 str1=expression.Mid(0,start);///���ǰ��μ��㺯��
			 int a1,a2;
			 a1=str1.ReverseFind('+');
			 a2=str1.ReverseFind('-');
			 if(a2>a1) a1=a2;
			 a2=str1.ReverseFind('*');
			 if(a2>a1) a1=a2;
			 a2=str1.ReverseFind('/');
			 if(a2>a1) a1=a2;
			 a2=str1.ReverseFind('(');
			 if(a2>a1) a1=a2;
			 if(a1==str1.GetLength()-1)//����
			 {
				 str3="";
			 }
			 else
			 {
				  str3=str1.Mid(a1+1);
			 }
			 if(str3=="")
				{
				}

				if(str3=="acos")
				{
					temp=acos(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="asin")
				{
					temp=asin(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="atan")
				{
					temp=atan(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="exp")
				{
					temp=exp(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}

				if(str3=="sin")
				{
					temp=sin(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="cos")
				{
					temp=cos(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}

				if(str3=="log")
				{
					temp=log(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="tan")
				{
					temp=tan(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="sqrt")
				{
					temp=sqrt(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="fabs")
				{
					temp=fabs(temp);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}

				if(str3=="pow")
				{
					double d_a=atof(a_a);
					double d_b=atof(a_b);
					temp=pow(d_a,d_b);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}

				if(str3=="max")
				{
					double d_a=atof(a_a);
					double d_b=atof(a_b);
					temp=max(d_a,d_b);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}
				if(str3=="min")
				{
					double d_a=atof(a_a);
					double d_b=atof(a_b);
					temp=min(d_a,d_b);
					str.Format("%f",temp);
					str.Replace('-','@');
					str2=str3+str2;
				}



				expression.Replace(str2,str);
				return CalString(expression);
		 }
			else
		 {
			 AfxMessageBox("���Ų�ƥ��!");
			 return -1;
		 }
		}
		//////////////////
		///////////////////����˳�
		start=expression.FindOneOf("*/");
		if(start!=-1)
		{
			t=start;
			str1=expression.Tokenize("+-*/",t);
			str2=expression.Mid(0,start);/////
			t=str2.ReverseFind('+');
			end=str2.ReverseFind('-');
			if(t>end) end=t;
			str3=str2.Mid(end+1);
			c=expression.GetAt(start);
			switch(c) {
	case '*':
		str2=str3+"*"+str1;
		//str3.Replace('@','-');
		//d1=atof(str3);
		d1=CalString(str3);///��������
		d2=CalString(str1);
		temp=d1*d2;
		str.Format("%f",temp);
		str.Replace('-','@');
		expression.Replace(str2,str);
		return CalString(expression);
		break;
	case '/':
		str2=str3+"/"+str1;
		d1=CalString(str3);
		d2=CalString(str1);
		temp=d1/d2;
		str.Format("%f",temp);
		str.Replace('-','@');
		expression.Replace(str2,str);
		return CalString(expression);
		break;
	default:break;
			}
		}

		///////////////////////
		/////////////////////////����Ӽ�
		start=expression.FindOneOf("+-");
		if(start!=-1)
		{
			t=start;
			str1=expression.Tokenize("+-",t);
			str2=expression.Mid(0,start);/////
			t=str2.ReverseFind('+');
			end=str2.ReverseFind('-');
			if(t>end) end=t;
			str3=str2.Mid(end+1);
			c=expression.GetAt(start);
			switch(c) {
	case '+':
		str2=str3+"+"+str1;

		d1=CalString(str3);
		d2=CalString(str1);
		temp=d1+d2;
		str.Format("%f",temp);
		str.Replace('-','@');

		expression.Replace(str2,str);
		return CalString(expression);
		break;
	case '-':
		str2=str3+"-"+str1;

		d1=CalString(str3);
		d2=CalString(str1);
		temp=d1-d2;
		str.Format("%f",temp);
		str.Replace('-','@');

		expression.Replace(str2,str);
		return CalString(expression);
		break;
	default:
		break;
			}
		}
		//////////////////////////////////////
		///////////
		if(expression=="PI")//////////���س���
		{
			return M_PI;
		}
		if(expression=="E")
		{
			return M_E; 
		}
		expression.Replace('@','-');
		temp=atof(expression);
}//////////////////try������

	catch (...) {
		AfxMessageBox("����");
		return -1;
	}
	return temp;
}


/////////////��ⷽ��
double CMyUtil::ComputeEquation(CString expression,CString x)
{
try
{
	////////////////////
	int start,end;
	CString left,right;
	CString str;
	str="("+x+")";
	expression.Replace(x,str);
	double dleft,dright;
	double c1,c2;
	double result;
	start=expression.Find('=');
	if(start==-1)
	{
		AfxMessageBox("���ǵ�ʽ");
		return -1;
	}
	end=expression.Find(x);
	if(end==-1)
	{
		AfxMessageBox("�Ҳ���δ֪��"+x);
		return -1;
	}
////���δ֪���ڵ�ʽ�ұߣ��������
	if(end>start)
	{
		str=expression;
		left=str.Left(start);
		right=str.Right(str.GetLength()-start-1);
		str=left+"-("+right+")=0";
		return ComputeEquation(str,x);
	}



	/////////////��ʼ����
	str=expression;
	left=str.Left(start);
	right=str.Right(str.GetLength()-start-1);
	///////////����������
	int min=1000;
	int nIndex;
	CStringArray  label;
	label.Add("sin");
	label.Add("cos");
	label.Add("log");
	label.Add("/");

	label.Add("acos");
	label.Add("asin");
	label.Add("atan");
;
	label.Add("exp");
	label.Add("fabs");
	label.Add("log");
	label.Add("log10");
	
	label.Add("tan");

	label.Add("sqrt");
	label.Add("pow");
	label.Add("max");
	label.Add("min");






	str=expression.Left(end);
	for(int i=0;i<label.GetSize();i++)
	{
		nIndex=str.Find(label.GetAt(i));
		if((min>nIndex)&&nIndex!=-1) 
			min=nIndex;
	}
	if(min!=1000)
	{
		str=expression; 
		int n=0,rIndex;
		nIndex=min;
		rIndex=min; 
		int j;
		for(j=min;j<str.GetLength();j++)
		{
			char ch=str.GetAt(j);
			if(ch=='(')
			{
				n++;
			}
			else if(ch==')')
			{
				n--;
				if(n==0)
					break;
			}
		}
		rIndex=j;
		CString fun,sub;
		if(rIndex>end)
		{

			str=expression;
			nIndex=min;
			fun=str.Tokenize("(",nIndex);

			///////////�õ�c1
			sub=str.Mid(min,rIndex-min+1);
			if(fun=="/")
			{
				str.Replace(sub,"*"+sub);
			}
			c1=ComputeEquation(str,sub);
			///////////
			str=expression;

			nIndex=min+fun.GetLength()+1;
			rIndex=rIndex-min-1-fun.GetLength();
			sub=str.Mid(nIndex,rIndex);
			//////�жϺ���
			if(fun=="/")
			{
				c1=1/c1;
				sub.Format("%s=%f",sub,c1);
				return ComputeEquation(sub,x);
			}

			if(fun=="sin")
			{
				c1=asin(c1);
				sub.Format("%s=%f",sub,c1);
				return ComputeEquation(sub,x);
			}

			if(fun=="cos")
			{
				c1=acos(c1);
				sub.Format("%s=%f",sub,c1);
				return ComputeEquation(sub,x);
			}

			if(fun=="log")
			{
				c1=exp(c1);
				sub.Format("%s=%f",sub,c1);
				return ComputeEquation(sub,x);
			}
			// /////////////////////////////
		}
		else
		{
			CString strTemp=str;
			str=expression;
			nIndex=min;
			fun=str.Tokenize("(",nIndex);
			///////////�õ�c1
			str=expression;
			str=str.Mid(min,rIndex-min+1);
			if(fun=='/')
			{
				str="1"+str;/////����1
			}
			c1=CalString(str);
			str.Format("%f",c1);
			str.Replace('-','@');
			if(fun=='/')
			{
				str="*"+str;////���ϳ˺�
			}
			expression.Replace(strTemp,str);
			return ComputeEquation(expression,x);
		}
	}
	///////////////////�����������

	///////////////
	//////////////////////////
	/////���������
	if(end>start)
	{
		dleft=CalString(left);
		str=right;
		str.Replace(x,"1");
		c1=CalString(str);
		str=right;
		str.Replace(x,"2");
		c2=CalString(str);
		result=(dleft-(2*c1-c2))/(c2-c1);
		return result;
	}
	else if(end<start)
	{
		dright=CalString(right);
		str=left;
		str.Replace(x,"1");
		c1=CalString(str);
		str=left;
		str.Replace(x,"2");
		c2=CalString(str);
		result=(dright-(2*c1-c2))/(c2-c1);
		return result;
	}
	else
	{
		AfxMessageBox("����");
		return -1;
	}
////////////////////////////
}///try������
		catch (...) {
			AfxMessageBox("���̼������");
			return -1;
		}
}

//////////////////////////��ⷽ�̽���
/************************************************************************/
/* �ҳ�δ֪���ڷ����е�λ��                                                                     */
/************************************************************************/

 int CMyUtil::FindX(CString expression,CString x)
 {
	 int nIndex;
	 char left,right;
	 CString str;
	 BOOL front=FALSE,end=FALSE;
	 nIndex=expression.Find(x);
	 while(nIndex!=-1)
	 {
		 if(nIndex==0)
			{
				front=TRUE;
			}
		 else
			{
				left=expression.GetAt(nIndex-1);
				if(left=='('||left=='='||left=='+'||left=='-'||left=='*'||left=='/'||left=='<'||left=='>'||left==',')
					front=TRUE;
			}

			if(nIndex+x.GetLength()==expression.GetLength())
			{
				end=TRUE;
			}
			else
			{
				right=expression.GetAt(nIndex+x.GetLength());
				if(right==')'||right=='='||right=='+'||right=='-'||right=='*'||right=='/'||right=='<'||right=='>'||right==',')
					end=TRUE;
			}
			if(front==TRUE&&end==TRUE)
			{
				return nIndex;
			}
			front=FALSE;
			end=FALSE;
			if(nIndex+x.GetLength()==expression.GetLength())
				return -1;
				nIndex=expression.Find(x,nIndex+1);
	 }
	 return -1;
 }

/************************************************************************/
/* �滻δ֪��Ϊ��ֵ                                                                     */
/************************************************************************/

CString CMyUtil::ReplaceXtoDouble(CString expression,CString x,double result)
{
	    int nIndex;
		CString str;
		str=expression;
		nIndex=FindX(str,x);
		if(nIndex==-1) return expression;
        
		if(result<0)
		{
			str.Format("@%f",-result);
		}
		else
		{
			str.Format("%f",result);
		}
	    

		if((nIndex+x.GetLength())<expression.GetLength())
			str=expression.Left(nIndex)+str+expression.Mid(nIndex+x.GetLength());
		else
			str=expression.Left(nIndex)+str;
		return ReplaceXtoDouble(str,x,result);
		
}

/************************************************************************/
/* �滻δ֪��Ϊ�ַ���                                                                     */
/************************************************************************/

CString CMyUtil::ReplaceXtoString(CString expression,CString x,CString result)
{
	int nIndex;
	CString str;
	str=expression;
	nIndex=FindX(str,x);
	if(nIndex==-1) return expression;

	str.Format("%s",result);
	if((nIndex+x.GetLength())<expression.GetLength())
		str=expression.Left(nIndex)+str+expression.Mid(nIndex+x.GetLength());
	else
		str=expression.Left(nIndex)+str;
	//return str;
	return ReplaceXtoString(str,x,result);

}

int CMyUtil::AddPostfix(CStringArray *strArray,CString strNum)
{
	CString str;
	for(int i=0;i<strArray->GetSize();i++)
	{
		str=(*strArray)[i];
        (*strArray)[i]=AddPostfix(str,strNum);
	}
	return 1;
}

CString CMyUtil::AddPostfix(CString expression,CString strNum)
{
	CStringArray tempArray;
	CString str,str1;
	int index=0;
	str=expression.Tokenize("+-*/()=",index);
	while(str!="")
	{
		tempArray.Add(str);
		str=expression.Tokenize("+-*/=()",index);
	}
for(int i=0;i<tempArray.GetSize();i++)
{
    str=tempArray[i];

	if(!IsDouble(str)&&!IsFunction(str))
	{
		str1=str+":"+strNum;
        expression=ReplaceXtoString(expression,str,str1);
	}
}

return expression;
}

/************************************************************************/
/* ��鲻��ʽ*/
/************************************************************************/
BOOL  CMyUtil::CheckInEquation(CString Expression,CStringArray * importDim,CStringArray * exportDim)
{
	int min;
	CString str,left,right,str1;
	str=Expression;
	min=str.FindOneOf("<>");
	if(min==-1)
	{
		str="����ʽ�޲��Ⱥ�";
		AfxMessageBox(str);
		return FALSE;
	}

	if(min==0)
	{
		str="���Ⱥ�ǰ����������";
		AfxMessageBox(str);
		return FALSE;
	}

	if(min==str.GetLength()-1)
	{
		str="���Ⱥź�����������";
		AfxMessageBox(str);
		return FALSE;
	}
	left=str.Left(min);
	str1=str.Mid(min+1,1);
	if(str1=="=")
		right=str.Mid(min+2);
	else
	    right=str.Mid(min+1);
	min=right.FindOneOf("<>=");
	if(min!=-1)
	{
		str="���ʽ���Ⱥ��ص�";
		AfxMessageBox(str);
		return FALSE;
	}
     min=left.Find("=");
	  if(min!=-1)
	  {
		  str="���Ⱥ�ǰ�еȺ�";
		  AfxMessageBox(str);
		  return FALSE;
	  }
	if(CheckExpression(left,importDim,exportDim))
	{
		if(CheckExpression(right,importDim,exportDim))
			return TRUE;
	}
	return FALSE;
}

/************************************************************************/
/* ���㲻��ʽ                                                                     */
/************************************************************************/
BOOL  CMyUtil::ComputeInEquation(CString expression)
{
	int index;
	CString expre=expression;
	CString str;
	CString left,right;
	double dleft,dright;
	index=expre.FindOneOf("<>");
	if(index!=-1)//�����С�ں�
	{
		str=expre.Mid(index+1,1);
		if(str!="=")///û�еȺ�
		{
			left=expre.Left(index);
			right=expre.Mid(index+1);
			dleft=CalString(left);
			dright=CalString(right);
			str=expre.Mid(index,1);
			if(str==">")
			{
				if(dleft>dright) return TRUE;
				else return FALSE;
			}
			if(str=="<")
			{
				if(dleft<dright) return TRUE;
				else return FALSE;
			}
		}
		else ////�еȺ�
		{
			left=expre.Left(index);
			right=expre.Mid(index+2);
			dleft=CalString(left);
			dright=CalString(right);
			str=expre.Mid(index,1);
			if(str==">")
			{
				if(dleft<dright) return FALSE;
				else return TRUE;
			}
			if(str=="<")
			{
				if(dleft>dright) return FALSE;
				else return TRUE;
			}
		}
	}
return FALSE;
}


/************************************************************************/
/* ����ϵ��                                                                     */
/************************************************************************/

double * CMyUtil::ComputeCoefficient(CString expression,CStringArray *x,int n)
{
	double *a=new double[n+1];
	return a;
}



/************************************************************************/
/*        ���㷽����                                                              */
/************************************************************************/
CStringArray *  CMyUtil::ComputeEquationGroup(CStringArray *expression,CStringArray *x,int n)
{
	double **a;
	a=new double *[n];
    int i;
	for(i=0;i<n;i++)
	{
		a[i]=new double[n];
	}
	

	for(i=0;i<n;i++)
	{
		delete[]  a[i];
	}
	delete [] a;

	
	CStringArray *result=new CStringArray;
	return result;
}


int CMyUtil::FindVar(CString expression,CStringArray & export,CString & x)
{

	int status;
	CString str;
	int num=0;
	for(int i=0;i<export.GetSize();i++)
	{
		str=export[i];
		status=this->FindX(expression,str);
		if(status!=-1)
		{
			x=str;
			num++;
		}
	}
	
	return num;
}
