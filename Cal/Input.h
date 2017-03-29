#pragma once

#include "MyUtil.h"


typedef struct Variant
{
	CString Name;
	CString strValue;
	double Value;
	int state;//0,未赋值 1,初值 2,计算值
	CString Note;
}Variant;



// CInput 命令目标

class CInput : public CObject
{
public:
	CInput();
	virtual ~CInput();
public:
	CMyUtil util;
	CArray<Variant,Variant>  m_VarArray;
    CStringArray m_SourceRel;
	CStringArray m_Rel;

	
	
public:
   int ReadFile(CString path);
   int ReadStr(CString text);
   int ParseVar(CString text);
   int ParseValue(CString text);
   int ParseRel(CString text);

   int Compute();
   int Check();
   int CalOnce();


   int GetPre(CString text,CString &pre,CString &next,CString blank);

   int WriteInput(CString path);
   int WriteOutput(CString path);

   int WriteDoc(CString path1,CString path2);


   int CalExEqual(CString input,CString & output);
   int WriteOutputStr(CString & FullText);
   




};


