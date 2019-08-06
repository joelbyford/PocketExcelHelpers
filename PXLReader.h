// PXLReader.h: interface for the CPXLReader class.
//
//////////////////////////////////////////////////////////////////////

#if 
!defined(AFX_PXLREADER_H__C0A20116_6BD3_4A98_84E4_96BFA477D44E__INCLUDED_)
#define AFX_PXLREADER_H__C0A20116_6BD3_4A98_84E4_96BFA477D44E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "pxltypes.h"
#include "stdio.h"

class CPXLReader
{
public:
	BOOL _EOF(void);
	DWORD ExportToCSV(TCHAR *szSourceFileName, TCHAR *szDestFileName);
	DWORD ExportToCSV(TCHAR *szFileName);
	void CloseFile(void);
	DWORD OpenPXLFile(TCHAR *szFileName);
	CPXLReader();
	virtual ~CPXLReader();

private:
	BOOL m_bEOF;
	double m_fVal;
	WCHAR m_szExtraData[MAX_PATH];
	WCHAR m_szRecordData[MAX_PATH];
	BYTE m_bCurrOpCode;
	DWORD GetNextRecord(void);
	BOOL FileExists(LPCTSTR szFileName);
	BOOL CheckFileHandle(void);
	DWORD m_dwRead;
	DWORD m_dwWritten;
	HANDLE m_hInFile;
    PXL_BLANK       m_pxlBlank;
    PXL_BOF         m_pxlBOF;
    PXL_BOOLERR     m_pxlBoolErr;
    PXL_BOUNDSHEET  m_pxlBoundSheet;
    PXL_COLINFO     m_pxlColInfo;
    PXL_DEFAULTROWHEIGHT    m_pxlDefaultRowHeight;
    PXL_DEFCOLWIDTH m_pxlDefaultColumnWidth;
    PXL_EOF         m_pxlEOF;
    PXL_FILEPASS    m_pxlFilePass;
    PXL_FONT        m_pxlFont;
    PXL_FORMAT      m_pxlFormat;
    PXL_FORMATCOUNT m_pxlFormatCount;
    PXL_FORMULA     m_pxlFormula;
    PXL_LABEL       m_pxlLabel;
    PXL_NAME        m_pxlName;
    PXL_NUMBER      m_pxlNumber;
    PXL_PANE        m_pxlPane;
    PXL_ROW         m_pxlRow;
    PXL_SELECTION   m_pxlSelection;
    PXL_STRING      m_pxlString;
    PXL_WINDOW2     m_pxlWindow2;
    PXL_XF          m_pxlXF;
    PXL_CODEPAGE    m_pxlCodepage;
    PXL_COUNTRY     m_pxlCountry;
    PXL_WINDOW1     m_pxlWindow;
};

#endif // 
!defined(AFX_PXLREADER_H__C0A20116_6BD3_4A98_84E4_96BFA477D44E__INCLUDED_)

