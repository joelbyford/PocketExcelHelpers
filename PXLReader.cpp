// PXLReader.cpp: implementation of the CPXLReader class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PXLReader.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPXLReader::CPXLReader()
{
    m_hInFile = INVALID_HANDLE_VALUE;
}

CPXLReader::~CPXLReader()
{
    CloseFile();
}

DWORD CPXLReader::OpenPXLFile(TCHAR *szFileName)
{
    BOOL bRetVal = TRUE;

	if (INVALID_HANDLE_VALUE != m_hInFile)
	{
		return ERROR_ALREADY_ASSIGNED; //85 - file is already open & being used.
	}

    if (!FileExists(szFileName))
    {
        return ERROR_FILE_NOT_FOUND; //2 - File not found
    }

    m_hInFile = INVALID_HANDLE_VALUE;

    m_hInFile = CreateFile(szFileName,  GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);


	if (INVALID_HANDLE_VALUE == m_hInFile)
	{
		return GetLastError();
	}

    m_bEOF = FALSE;

    return 0;
}

BOOL CPXLReader::CheckFileHandle()
{
	if (INVALID_HANDLE_VALUE == m_hInFile)
	{
		return FALSE;
	}
    //implied else
    return TRUE;
}

BOOL CPXLReader::FileExists(LPCTSTR szFileName)
{
    if (0xFFFFFFFF == GetFileAttributes(szFileName))
        return FALSE;
    else
        return TRUE;
}

void CPXLReader::CloseFile()
{
	/*if (!CheckFileHandle())
        return; //if file's already closed, don't close it.*/

    if (m_hInFile != INVALID_HANDLE_VALUE)
        CloseHandle(m_hInFile);
    m_hInFile = INVALID_HANDLE_VALUE;
	return;
}

DWORD CPXLReader::GetNextRecord()
{

    BOOL bResult;
    //read the opcode, back up the file pointer, then based on opcode type, read x bytes
    m_bCurrOpCode = 0xFF;
    bResult = ReadFile(m_hInFile, &m_bCurrOpCode, sizeof(BYTE), &m_dwRead, NULL);
    if (bResult && m_dwRead == 0)
    {
        // we’re at the end of the file
        m_bEOF = TRUE;
        return 0;
    }
    for (int i = 0; i < sizeof(BYTE); i++)
    {
        SetFilePointer(m_hInFile, -1, NULL, FILE_CURRENT);
    }
    memset(&m_szRecordData, 0, sizeof(WCHAR) * MAX_PATH);
    memset(&m_szExtraData, 0, sizeof(WCHAR) * MAX_PATH);
    switch (m_bCurrOpCode)
    {
        case PXLOP_BLANK:
        {
            bResult = ReadFile(m_hInFile, &m_pxlBlank, sizeof(PXL_BLANK), &m_dwRead, NULL);
            break;
        }
        case PXLOP_BOF:
        {
            bResult = ReadFile(m_hInFile, &m_pxlBOF, sizeof(PXL_BOF), &m_dwRead, NULL);
            break;
        }
        case PXLOP_BOOLERR:
        {
            bResult = ReadFile(m_hInFile, &m_pxlBoolErr, sizeof(PXL_BOOLERR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_BOUNDSHEET:
        {
            bResult = ReadFile(m_hInFile, &m_pxlBoundSheet, sizeof(PXL_BOUNDSHEET), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlBoundSheet.charLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_CODEPAGE:
        {
            bResult = ReadFile(m_hInFile, &m_pxlCodepage, sizeof(PXL_CODEPAGE), &m_dwRead, NULL);
            break;
        }
        case PXLOP_COLINFO:
        {
            bResult = ReadFile(m_hInFile, &m_pxlColInfo, sizeof(PXL_COLINFO), &m_dwRead, NULL);
            break;
        }
        case PXLOP_COUNTRY:
        {
            bResult = ReadFile(m_hInFile, &m_pxlCountry, sizeof(PXL_COUNTRY), &m_dwRead, NULL);
            break;
        }
        case PXLOP_DEFAULTROWHEIGHT:
        {
            bResult = ReadFile(m_hInFile, &m_pxlDefaultRowHeight, sizeof(PXL_DEFAULTROWHEIGHT), &m_dwRead, NULL);
            break;
        }
        case PXLOP_DEFCOLWIDTH:
        {
            bResult = ReadFile(m_hInFile, &m_pxlDefaultColumnWidth, sizeof(PXL_DEFCOLWIDTH), &m_dwRead, NULL);
            break;
        }
        case PXLOP_EOF:
        {
            bResult = ReadFile(m_hInFile, &m_pxlEOF, sizeof(PXL_EOF), &m_dwRead, NULL);
            break;
        }
        case PXLOP_FILEPASS:
        {
            //bResult = ReadFile(m_hInFile, &m_pxlFilePass, sizeof(PXL_FILEPASS), &m_dwRead, NULL);
            //if you encounter this, it will not be possible to read any of the remaining
            //portions of the document, so abort!
            return ERROR_INVALID_FUNCTION;
            break;
        }
        case PXLOP_FONT:
        {
            bResult = ReadFile(m_hInFile, &m_pxlFont, sizeof(PXL_FONT), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlFont.cch * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_FORMAT:
        {
            bResult = ReadFile(m_hInFile, &m_pxlFormat, sizeof(PXL_FORMAT), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlFormat.charLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_FORMULA:
        {
            bResult = ReadFile(m_hInFile, &m_pxlFormula, sizeof(PXL_FORMULA), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlFormula.charLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_LABEL:
        {
            bResult = ReadFile(m_hInFile, &m_pxlLabel, sizeof(PXL_LABEL), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlLabel.charLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_NAME:
        {
            bResult = ReadFile(m_hInFile, &m_pxlName, sizeof(PXL_NAME), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlName.nameLen * sizeof(WCHAR), &m_dwRead, NULL);
            if (m_pxlName.defLen)
                bResult = ReadFile(m_hInFile, &m_szExtraData, m_pxlName.defLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_NUMBER:
        {
            bResult = ReadFile(m_hInFile, &m_pxlNumber, sizeof(PXL_NUMBER), &m_dwRead, NULL);
            break;
        }
        case PXLOP_PANE:
        {
            bResult = ReadFile(m_hInFile, &m_pxlPane, sizeof(PXL_PANE), &m_dwRead, NULL);
            break;
        }
        case PXLOP_ROW:
        {
            bResult = ReadFile(m_hInFile, &m_pxlRow, sizeof(PXL_ROW), &m_dwRead, NULL);
            break;
        }
        case PXLOP_SELECTION:
        {
            bResult = ReadFile(m_hInFile, &m_pxlSelection, sizeof(PXL_SELECTION), &m_dwRead, NULL);
            break;
        }
        case PXLOP_STRING:
        {
            bResult = ReadFile(m_hInFile, &m_pxlString, sizeof(PXL_STRING), &m_dwRead, NULL);
            bResult = ReadFile(m_hInFile, &m_szRecordData, m_pxlString.charLen * sizeof(WCHAR), &m_dwRead, NULL);
            break;
        }
        case PXLOP_WINDOW1:
        {
            bResult = ReadFile(m_hInFile, &m_pxlWindow, sizeof(PXL_WINDOW1),&m_dwRead, NULL);
            break;
        }
        case PXLOP_WINDOW2:
        {
            bResult = ReadFile(m_hInFile, &m_pxlWindow2, sizeof(PXL_WINDOW2), &m_dwRead, NULL);
            break;
        }
        case PXLOP_XF:
        {
            bResult = ReadFile(m_hInFile, &m_pxlXF, sizeof(PXL_XF), &m_dwRead, NULL);
            break;
        }
        default: //a serious error has occurred, because if we don't know the opcode, we're sunk. abort.
        {
            m_bEOF = TRUE;
            return ERROR_INVALID_FUNCTION;
        }
    }
    if (bResult && m_dwRead == 0)
    {
        // we’re at the end of the file
        m_bEOF = TRUE;
    }
    return 0;
}

DWORD CPXLReader::ExportToCSV(TCHAR *szFileName)
{
    HANDLE hOutFile;
    char szTemp[MAX_PATH];
    long lrow = 0;
    long lcol = 0;


	hOutFile = INVALID_HANDLE_VALUE;

    hOutFile = CreateFile(szFileName,  GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);


	if (INVALID_HANDLE_VALUE == hOutFile)
	{
		return GetLastError();
	}

    while (! (_EOF()) )
    {

        GetNextRecord();
        switch (m_bCurrOpCode)
        {
            //
            case PXLOP_NUMBER:
            {
                if (m_pxlNumber.column != lcol)
                {
                    //write a comma
                    lcol = m_pxlNumber.column;
                    memset(&szTemp, 0, MAX_PATH);
                    sprintf(szTemp, ",");
                    WriteFile(hOutFile, &szTemp, strlen(szTemp), &m_dwWritten, NULL);
                }
                if (m_pxlNumber.row != lrow)
                {
                    //write a crlf pair
                    lrow = m_pxlNumber.row;
                    memset(&szTemp, 0, MAX_PATH);
                    sprintf(szTemp, "\r\n");
                    WriteFile(hOutFile, &szTemp, strlen(szTemp), &m_dwWritten, NULL);
                }
                memset(&szTemp, 0, MAX_PATH);
                sprintf(szTemp, "%f", m_pxlNumber.value);
                WriteFile(hOutFile, &szTemp, strlen(szTemp), &m_dwWritten, NULL);
                break;
            }
        }
    }
    CloseHandle(hOutFile);

    return 0;
}

DWORD CPXLReader::ExportToCSV(TCHAR *szSourceFileName, TCHAR *szDestFileName)
{
    //one-stop shopping :)
    OpenPXLFile(szSourceFileName);
    ExportToCSV(szDestFileName);
    CloseFile();
    return 0;
}

BOOL CPXLReader::_EOF()
{
    return (m_bEOF);
}

