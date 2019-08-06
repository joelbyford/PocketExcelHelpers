// PXLWriter.cpp: implementation of the CPXLWriter class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PXLWriter.h"
#include "stdio.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPXLWriter::CPXLWriter()
{
    m_hOutFile = INVALID_HANDLE_VALUE;
}

CPXLWriter::~CPXLWriter()
{
    CloseFile();
}

BOOL CPXLWriter::SetFont(WCHAR *FontName, short FontHeight, FontFormatting
FontFormat)
{
	BYTE l;
	PXL_FONT font;

	//Font length
	l = wcslen(FontName);

	font.opCode = 0x31;
	font.dwHeight = FontHeight * 20;  //Not sure why, but store 20 times the point size
	font.grbit = 0;                   //supposed to be 2d if italics.  Haven't tried it yet.
	font.reserved1 = 0x00FF;
	font.bls = 400;                   //Not sure why this is 190h (400d)
	font.reserved2 = 0;
	font.uls = 0;
	font.bFamily = 0;
	font.bCharSet =0;
	font.reserved3 = 0;
	font.cch = l;

    //Then the actual font data
    WriteFile(m_hOutFile, &font, sizeof(font), &m_dwWritten, NULL);
    WriteFile(m_hOutFile, FontName, l * sizeof(WCHAR), &m_dwWritten, NULL);
	//fwrite(&font, sizeof(font), 1, FileNumber);
	//fwrite(FontName, (l*2), 1, FileNumber);

    return(TRUE);
}

short CPXLWriter::CalcBaseAttrib(BOOL bFixed, BOOL bLocked, BOOL bHidden, BOOL bNumPrefix)
{
	//Does not look right.  Check when I get a Excel 97 Developmenters book.

	//false equals zero true equals one, so multiplying by zero yeilds zero and 1
	//yeilds itself.  Add all the options up and you have our base attributes.
	//return((bFixed*0x8000)+(bLocked*0x4000)+(bHidden*0x2000)+(bNumPrefix*0x1000));
	return(2);
}

short CPXLWriter::CalcTextAttrib()
{
	//false equals zero true equals one, so multiplying by zero yeilds zero and 1
	//yeilds itself.  Add all the options up and you have our base attributes.
	return(48);
}

BYTE CPXLWriter::CalcColAttrib(BOOL bHidden, BOOL bUserSet)
{
	//Does not look right.

	//false equals zero true equals one, so multiplying by zero yeilds zero and 1
	//yeilds itself.  Add all the options up and you have our base attributes.
	return((bHidden*0x80)+(bUserSet*0x40));

}

BYTE CPXLWriter::CalcFrmlOpts(BOOL bAlwaysCalc, BOOL bCalcOnLoad)
{
	//false equals zero true equals one, so multiplying by zero yeilds zero and 1
	//yeilds itself.  Add all the options up and you have our base attributes.
	return((bAlwaysCalc*0x80)+(bCalcOnLoad*0x40));
}

short CPXLWriter::CalcWindow1(BOOL bHidden, BOOL bMinimized, BOOL bHScroll, BOOL bVScroll, BOOL bShowTabs)
{
	//false equals zero true equals one, so multiplying by zero yeilds zero and1
	//yeilds itself.  Add all the options up and you have our base attributes.
	return((bHidden*0x8000)+(bMinimized*0x4000)+(1*0x2000)+(bHScroll*0x1000)+(bVScroll*0x0800)+(bShowTabs*0x0400)+(1*0x0200)+(1*0x0100));
}

BOOL CPXLWriter::SetXF(short fontIdx, short formatIdx, short baseAttrib, short textAttrib)
{
	//XF
	PXL_XF xf;
	xf.opCode = 0xE0;
	xf.ixfnt = fontIdx;
	xf.ixnf = formatIdx;
	xf.fattributes = 0xFFFFFFFF;     //Not used

	//Base Attribute flags

	//xf.fBaseAttr = 0x0002;
	xf.fBaseAttr = baseAttrib;
	//xf.fTextAttr = 0x0030;
	xf.fTextAttr = textAttrib;
	xf.icvFore = 0xFF;
	xf.icvFill = 0xFF;
	xf.bRight = 0xFF;
	xf.bLeft = 0xFF;
	xf.bTop = 0xFF;
	xf.bBottom = 0xFF;
	xf.backstyle = 0;
	xf.borderstyle = 0;
	//fwrite(&xf, sizeof(xf), 1, FileNumber);
    WriteFile(m_hOutFile, &xf, sizeof(xf), &m_dwWritten, NULL);
	return(TRUE);
}

BOOL CPXLWriter::SetBoundSheet(WCHAR *SheetName)
{
	PXL_BOUNDSHEET sheet;
	BYTE l;

	//Shet name length
	l=wcslen(SheetName);

	sheet.opCode = 0x85;
	sheet.reserved = 0;
	sheet.charLen = l;
	//fwrite(&sheet, sizeof(sheet), 1, FileNumber);
	//fwrite(SheetName, l*2, 1, FileNumber);
    WriteFile(m_hOutFile, &sheet, sizeof(sheet), &m_dwWritten, NULL);
    WriteFile(m_hOutFile, SheetName, l * sizeof(WCHAR), &m_dwWritten, NULL);
	return(TRUE);
}

BOOL CPXLWriter::AddSheet(WCHAR *SheetName)
{
    return (SetBoundSheet(SheetName));
}

BOOL CPXLWriter::SetDefColWidth(short charsWide)
{
	PXL_DEFCOLWIDTH defCW;
	defCW.opCode = 0x55;
	defCW.grbit = 0;
	defCW.coldx = charsWide*256;      //expressed in 256ths of a character
	defCW.ixfe = 0;
	//fwrite(&defCW, sizeof(defCW), 1, FileNumber);
    WriteFile(m_hOutFile, &defCW, sizeof(defCW), &m_dwWritten, NULL);
	return(TRUE);
}

BOOL CPXLWriter::SetDefRowHeight(double pixelsHigh)
{
	PXL_DEFAULTROWHEIGHT defRH;
	defRH.opCode = 0x25;
	defRH.grbit = 0;              //hidden.  Doesn't make sense!
	defRH.miyRw = (short)(pixelsHigh*20);         //height in twips or 1/20th of a point
	//fwrite(&defRH, sizeof(defRH), 1, FileNumber);
    WriteFile(m_hOutFile, &defRH, sizeof(defRH), &m_dwWritten, NULL);
	return(TRUE);
}

BOOL CPXLWriter::SetWindow2(short topRow, BYTE leftCol)
{
	PXL_WINDOW2 w2;
	w2.opCode = 0x3e;
	w2.colleft = leftCol;
	w2.rwtop = topRow;
	w2.grbit = 0;  //frozen stuff.

	//fwrite(&w2, sizeof(w2), 1, FileNumber);
    WriteFile(m_hOutFile, &w2, sizeof(w2), &m_dwWritten, NULL);
	return(TRUE);
}

BOOL CPXLWriter::SetSelection(short topRow, BYTE leftCol, short botRow, BYTE rightCol, short activeRow, BYTE activeCol)
{
	PXL_SELECTION sel;
	sel.opCode = 0x1D;
	sel.rowTop = topRow;
	sel.rowBot = botRow;
	sel.colLeft = leftCol;
	sel.colRight = rightCol;
	sel.colActive = activeCol;
	sel.rowActive = activeRow;
	//fwrite(&sel, sizeof(sel), 1, FileNumber);
    WriteFile(m_hOutFile, &sel, sizeof(sel), &m_dwWritten, NULL);
	return(TRUE);
}

void CPXLWriter::WriteLabel(short row, BYTE col, WCHAR *value, short iXF)
{
	PXL_LABEL TEXT_RECORD;
	short l;


	l = wcslen((WCHAR*)value);
	TEXT_RECORD.opCode = 4;
	//Length of the text portion of the record
	TEXT_RECORD.charLen = l;

	TEXT_RECORD.row = row;
	TEXT_RECORD.column = col;
	TEXT_RECORD.iXFR = iXF;

	//fwrite(&TEXT_RECORD, sizeof(TEXT_RECORD), 1, FileNumber);
	//fwrite(value, l*2, 1, FileNumber);
    WriteFile(m_hOutFile, &TEXT_RECORD, sizeof(TEXT_RECORD), &m_dwWritten, NULL);
    WriteFile(m_hOutFile, value, l * sizeof(WCHAR), &m_dwWritten, NULL);

	return;
}

void CPXLWriter::WriteFormula(short row, BYTE col, WCHAR *value, double fCurVal, short iXF)
{
	PXL_FORMULA oFrml;
	short l;

	oFrml.opCode = 0x06;
	oFrml.row = row;
	oFrml.column = col;
	oFrml.iXFR = iXF;
	oFrml.curValue = fCurVal;
	oFrml.options = CalcFrmlOpts(TRUE, FALSE);
	l = wcslen((WCHAR*)value);
	oFrml.charLen = l;

	//fwrite(&oFrml, sizeof(oFrml), 1, FileNumber);
	//fwrite(value, l*2, 1, FileNumber);
    WriteFile(m_hOutFile, &oFrml, sizeof(oFrml), &m_dwWritten, NULL);
    WriteFile(m_hOutFile, value, l * sizeof(WCHAR), &m_dwWritten, NULL);

	return;
}

void CPXLWriter::WriteDouble(short row, BYTE col, double value, short iXF)
{
	PXL_NUMBER NUMBER_RECORD;
	NUMBER_RECORD.opCode = 3;
	NUMBER_RECORD.iXFR = iXF;
	NUMBER_RECORD.row = row;
	NUMBER_RECORD.column = col;
	NUMBER_RECORD.value = value;
	//fwrite(&NUMBER_RECORD, sizeof(NUMBER_RECORD), 1, FileNumber);
    WriteFile(m_hOutFile, &NUMBER_RECORD, sizeof(NUMBER_RECORD), &m_dwWritten, NULL);
	return;
}

BOOL CPXLWriter::WriteWorkbook()
{
	BOOL bRetVal=TRUE;

	//BOF
	PXL_BOF BOF_Marker;
	BOF_Marker.opCode = 9;
	BOF_Marker.type = PXL_WORKSHEET;
	BOF_Marker.version = PXL_VERSION;
	//fwrite(&BOF_Marker, sizeof(BOF_Marker), 1, FileNumber);
    WriteFile(m_hOutFile, &BOF_Marker, sizeof(BOF_Marker), &m_dwWritten, NULL);

	//CODEPAGE
	PXL_CODEPAGE oCPage;
	oCPage.opCode = 0x42;
	oCPage.icv = 0x04E4;
	//fwrite(&oCPage, sizeof(oCPage), 1, FileNumber);
    WriteFile(m_hOutFile, &oCPage, sizeof(oCPage), &m_dwWritten, NULL);

	//COUNTRY
	PXL_COUNTRY oCntry;
	oCntry.opCode = 0x8C;
	oCntry.iCountryDef = 1;
	oCntry.iCountryWinIni = 1;
	//fwrite(&oCntry, sizeof(oCntry), 1, FileNumber);
    WriteFile(m_hOutFile, &oCntry, sizeof(oCntry), &m_dwWritten, NULL);

	//WINDOW1
	PXL_WINDOW1 oWnd1;
	oWnd1.opCode = 0x3D;
	oWnd1.grbit = CalcWindow1(FALSE, FALSE, TRUE, TRUE, TRUE);
	oWnd1.itabCur = 0;
	//fwrite(&oWnd1, sizeof(oWnd1), 1, FileNumber);
    WriteFile(m_hOutFile, &oWnd1, sizeof(oWnd1), &m_dwWritten, NULL);

	//FONT
	SetFont(TEXT("Tahoma"),10, pxlNoFormat);

	//XF - Accounting Formatting
	SetXF(0,0x04, CalcBaseAttrib(FALSE, FALSE, FALSE, FALSE),CalcTextAttrib());

	//XF - General Formatting
	SetXF(0,0, CalcBaseAttrib(FALSE, FALSE, FALSE, FALSE),CalcTextAttrib());


	//BOUNDSHEET
	SetBoundSheet(TEXT("Sheet1"));

	//EOF
	PXL_EOF EOF_Marker;
	EOF_Marker.opCode = 0x0A;

	//fwrite(&EOF_Marker, sizeof(EOF_Marker), 1, FileNumber);
    WriteFile(m_hOutFile, &EOF_Marker, sizeof(EOF_Marker), &m_dwWritten, NULL);

	return(bRetVal);

}

BOOL CPXLWriter::WriteWorksheet()
{
	BOOL bRetVal=TRUE;

	//BOF
	PXL_BOF BOF_Marker;
	BOF_Marker.opCode = 0x09;
	BOF_Marker.type = PXL_WORKBOOK;
	BOF_Marker.version = PXL_VERSION;
	//fwrite(&BOF_Marker, sizeof(BOF_Marker), 1, FileNumber);
    WriteFile(m_hOutFile, &BOF_Marker, sizeof(BOF_Marker), &m_dwWritten, NULL);

	//DEFAULTCOLWIDTH
	SetDefColWidth(9);   //9 characters wide

	//COLINFO opcode (7Dh)
	PXL_COLINFO oColInfo;
	oColInfo.opCode = 0x7D;
	oColInfo.firstCol = 0;
	oColInfo.lastCol = 255;
	oColInfo.width = 9*256;  //9 characters wide
	oColInfo.iXFR = 0;
	oColInfo.options = CalcColAttrib(TRUE, FALSE);
	//fwrite(&oColInfo, sizeof(oColInfo), 1, FileNumber);
    WriteFile(m_hOutFile, &oColInfo, sizeof(oColInfo), &m_dwWritten, NULL);

	//DEFAULTROWHEIGHT
	SetDefRowHeight(12.75);

	//Data Stuff goes here.
	//WriteValue(xlsText, 0, 0, TEXT("a1"));
	//WriteValue(xlsText, 1, 0, TEXT("a2"));
	//WriteValue(xlsText, 0, 1, TEXT("b1"));
	//WriteValue(xlsText, 1, 1, TEXT("b2"));
	//WriteValue(xlsNumber, 0, 2, TEXT("10000"));
	//WriteValue(xlsNumber, 1, 2, TEXT("20000"));


	return(bRetVal);

}

DWORD CPXLWriter::CreatePXLFile(TCHAR *szFileName)
{
    BOOL bRetVal = TRUE;

	if (INVALID_HANDLE_VALUE != m_hOutFile)
	{
		return ERROR_ALREADY_ASSIGNED; //85 - file is already open & being used.
	}

    if (FileExists(szFileName))
    {
        DeleteFile(szFileName);
    }

    m_hOutFile = INVALID_HANDLE_VALUE;

    m_hOutFile = CreateFile(szFileName,  GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);


	if (INVALID_HANDLE_VALUE == m_hOutFile)
	{
		return GetLastError();
	}


	//Open a file for writing
	//FileNumber = _wfopen(g_szFile, TEXT("wb"));*/

	//Write the default formats
	bRetVal = WriteWorkbook();
	WriteWorksheet();
    if (bRetVal)
    {
        return 0; //No error
    }
    else
    {
        return GetLastError();
    }
}

void CPXLWriter::CloseFile()
{
	if (!CheckFileHandle())
        return; //if file's already closed, don't close it.

    //WINDOW2
	SetWindow2(0,0);

	//SELECTION
	SetSelection(0,0,0,0,0,0);

	//EOF
	PXL_EOF EOF_Marker;
	EOF_Marker.opCode = 0x0A;

	//fwrite(&EOF_Marker, sizeof(EOF_Marker), 1, FileNumber);
	//fclose(FileNumber);
    WriteFile(m_hOutFile, &EOF_Marker, sizeof(EOF_Marker), &m_dwWritten, NULL);

    CloseHandle(m_hOutFile);
	return;
}


BOOL CPXLWriter::CheckFileHandle()
{
	if (INVALID_HANDLE_VALUE == m_hOutFile)
	{
		return FALSE;
	}
    //implied else
    return TRUE;
}

BOOL CPXLWriter::FileExists(LPCTSTR szFileName)
{
    if (0xFFFFFFFF == GetFileAttributes(szFileName))
        return FALSE;
    else
        return TRUE;
}

