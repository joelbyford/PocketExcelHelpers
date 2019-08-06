// PXLWriter.h: interface for the CPXLWriter class.
//
//////////////////////////////////////////////////////////////////////

#if 
!defined(AFX_PXLWRITER_H__91FCE1DF_E80E_4C23_A10F_F799D754676C__INCLUDED_)
#define AFX_PXLWRITER_H__91FCE1DF_E80E_4C23_A10F_F799D754676C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pxltypes.h"

class CPXLWriter
{
public:
	void CloseFile(void);
	DWORD CreatePXLFile(TCHAR *szFileName);
	BOOL WriteWorksheet(void);
	BOOL WriteWorkbook(void);
	void WriteDouble(short row, BYTE col, double value, short iXF);
	void WriteFormula(short row, BYTE col, WCHAR* value, double fCurVal, short 
iXF);
	void WriteLabel(short row, BYTE col, WCHAR *value, short iXF);
	BOOL SetSelection(short topRow, BYTE leftCol, short botRow, BYTE rightCol, 
short activeRow, BYTE activeCol);
	BOOL SetWindow2(short topRow, BYTE leftCol);
	BOOL SetDefRowHeight(double pixelsHigh);
	BOOL SetDefColWidth(short charsWide);
	BOOL AddSheet(WCHAR *SheetName);
	BOOL SetBoundSheet(WCHAR *SheetName);
	BOOL SetXF(short fontIdx, short formatIdx, short baseAttrib, short 
textAttrib);
	short CalcWindow1(BOOL bHidden, BOOL bMinimized, BOOL bHScroll, BOOL 
bVScroll, BOOL bShowTabs);
	BYTE CalcFrmlOpts(BOOL bAlwaysCalc, BOOL bCalcOnLoad);
	BYTE CalcColAttrib(BOOL bHidden, BOOL bUserSet);
	short CalcTextAttrib(void);
	short CalcBaseAttrib(BOOL bFixed, BOOL bLocked, BOOL bHidden, BOOL 
bNumPrefix);
	BOOL SetFont(WCHAR *FontName, short FontHeight, FontFormatting FontFormat);
	CPXLWriter();
	virtual ~CPXLWriter();

private:
	BOOL FileExists(LPCTSTR szFileName);
	BOOL CheckFileHandle(void);
	DWORD m_dwRead;
	DWORD m_dwWritten;
	HANDLE m_hOutFile;
};

#endif // 
!defined(AFX_PXLWRITER_H__91FCE1DF_E80E_4C23_A10F_F799D754676C__INCLUDED_)

