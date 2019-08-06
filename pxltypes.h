//pxltypes.h

#ifndef __PXLTYPES_H__
#define __PXLTYPES_H__

//////////////////////////////////////
// Local Definitions
//////////////////////////////////////
#define PXL_VERSION 0x010f
#define PXL_WORKSHEET 0x0005
#define PXL_WORKBOOK 0x0010
#define SZ_FILEFILTER		TEXT("Pocket Excel (*.pxl)\0*.pxl\0All Files (*.*)\0*.*\0")
#define SZ_DEFEXT			TEXT("pxl")

//==================

enum eValueTypes
{
    pxlBool = 0,
    pxlNumber = 1,
    pxlText = 2,
	pxlFormula = 3
};

enum FontFormatting
{
   //add these enums together. For example: pxlBold + pxlUnderline
   pxlNoFormat = 0,
   pxlBold = 1,
   pxlItalic = 2,
   pxlUnderline = 4,
   pxlStrikeout = 8
};

#pragma pack(push,1)

/////////////////////////////////////////////////////////////////
//  opcodes for pocket excel
/////////////////////////////////////////////////////////////////

//PocketExcel Commands

#define PXLOP_BLANK   0x01
#define PXLOP_BOF     0x09
#define PXLOP_BOOLERR 0x05
#define PXLOP_BOUNDSHEET  0x85
#define PXLOP_CODEPAGE    0x42
#define PXLOP_COLINFO     0x7D
#define PXLOP_COUNTRY     0x8C
#define PXLOP_DEFAULTROWHEIGHT 0x25
#define PXLOP_DEFCOLWIDTH 0x55
#define PXLOP_EOF         0x0A
#define PXLOP_FILEPASS    0x2F
#define PXLOP_FONT        0x31
#define PXLOP_FORMAT      0x1E
#define PXLOP_FORMULA     0x06
#define PXLOP_LABEL       0x04
#define PXLOP_NAME        0x18
#define PXLOP_NUMBER      0x03
#define PXLOP_PANE        0x41
#define PXLOP_ROW         0x08
#define PXLOP_SELECTION   0x1D
#define PXLOP_STRING      0x07
#define PXLOP_WINDOW1     0x3D
#define PXLOP_WINDOW2     0x3E
#define PXLOP_XF          0xE0
////////////////////////////////////////////////////////////////
//BLANK (01h):Blank cell
struct PXL_BLANK
{
	BYTE opCode;
	short row;
	BYTE column;
	short iXFR;
};

//BOF (09h): Beginning of worksheet/workbook
struct PXL_BOF
{
	BYTE opCode;
	short version;
	short type;  //0005h=workbook, 0010h=Worksheet
};

//BOOLERR (05h): Boolean or Error value
struct PXL_BOOLERR
{
	BYTE opCode;
	short row;
	BYTE col;
	short iXFR;
	BYTE value;
	BYTE type;  //0=Value is Bool, 1=Value is error
};

//BOUNDSHEET (85h): Sheet Information
struct PXL_BOUNDSHEET
{
	BYTE opCode;
	BYTE reserved;  //reserved for future MS use.  Must equal 0!
	BYTE charLen;
	//sheet name goes here with variable length.
};


//COLINFO(2+) (7Dh): Column Formatting [Pocekt Excel 2.0]
struct PXL_COLINFO
{
	BYTE opCode;
	short firstCol;
	short lastCol;
	short width; //in 1/255ths of a character
	short iXFR;
	BYTE options;
};

//DEFAULTROWHEIGHT(25h): Default row height
struct PXL_DEFAULTROWHEIGHT
{
	BYTE opCode;
	short grbit;
	short miyRw; //in twips
};

//DEFCOLWIDTH(55h): Default column width
struct PXL_DEFCOLWIDTH
{
	BYTE opCode;
	short grbit;
	short coldx;
	short ixfe;
};

//EOF(0Ah): End of workbook or worksheet
struct PXL_EOF
{
	BYTE opCode;
};

//FILEPASS(2Fh): File password
//------------
// Don't currently support encryption.
//------------
struct PXL_FILEPASS
{
	BYTE opCode;
};

//FONT(31h): Font Description
struct PXL_FONT
{
	BYTE opCode;
	short dwHeight;
	short grbit;
	short reserved1;
	short bls;
	short reserved2;
	BYTE uls;
	BYTE bFamily;
	BYTE bCharSet;
	BYTE reserved3;
	BYTE cch;
	//Font name goes here with variable length as defined above.
};



//FORMAT(1Eh): Number Format
struct PXL_FORMAT
{
	BYTE opCode;
	BYTE charLen;
	//Format name goes here with variable length as defined above.
};

//FORMATCOUNT(1Fh): Number of Formats in Document
struct PXL_FORMATCOUNT
{
	BYTE opCode;
	BYTE count;
	BYTE field2;
	BYTE field3;
};

//FORMULA (06h): Formula cell
struct PXL_FORMULA
{
	BYTE opCode;
	short row;
	BYTE column;
	short iXFR;
	double curValue; //Current value of the formula post calculation
	BYTE options;
	short charLen;
	//Formula goes here with variable length as defined above.
};

//LABEL(04h): string constant
struct PXL_LABEL
{
	BYTE opCode;
	short row;
	BYTE column;
	short iXFR;
	short charLen; //between 0 and 255 only
	//String goes here with variable length as defined above;
};

//NAME(18h): A Named cell or range in the workbook
struct PXL_NAME
{
	BYTE opCode;
	short options;
	BYTE nameLen;  //in characters (Max 16)
	short defLen; //in bytes (Max 64)
	short idxSheet;
	//Add Name here with variable length as defined above
	//Add Definition here weith variable length as defined above.
};

//NUMBER(03h): Any number value in a cell
struct PXL_NUMBER
{
	BYTE opCode;
	short row;
	BYTE column;
	short iXFR;
	double value;
};

//PANE(41h): Panes in a window
struct PXL_PANE
{
	BYTE opCode;
	short horoPosition;
	short vertPosition;
	short topRowVisible;
	short leftColVisible;
	BYTE paneNum;  //0=botom right; 1=top right; 2=botom left; 3=top left
};

//ROW(08h): Special row formatting
struct PXL_ROW
{
	BYTE opCode;
	short rowNum;
	short rowHeight;
	short options; //Bit 1=hidden; Bit 2=User Set
	short iXFR;
};

//SELECTION(1Dh): What is currently selected
struct PXL_SELECTION
{
	BYTE opCode;
	short rowTop;
	BYTE colLeft;
	short rowBot;
	BYTE colRight;
	short rowActive;
	BYTE colActive;
};

//STRING(07h): String ref from formula
struct PXL_STRING
{
	BYTE opCode;
	BYTE charLen;
	//String goes here with the variable length as defined above.
};


//WINDOW2(3Eh): Sheet level attributes
struct PXL_WINDOW2
{
	BYTE opCode;
	short rwtop;
	BYTE colleft;
	short grbit;
};

//XF(E0h): Extended Format
struct PXL_XF
{
	BYTE opCode;
	short ixfnt;
	short ixnf;
	long fattributes;
	short fBaseAttr;
	short fTextAttr;
	short icvFore;
	short icvFill;
	BYTE bRight;
	BYTE bTop;
	BYTE bLeft;
	BYTE bBottom;
	BYTE backstyle;
	BYTE borderstyle;
};

//CODEPAGE (42h)
struct PXL_CODEPAGE
{
	BYTE opCode;
	short icv;
};

//COUNTRY (8Ch)
struct PXL_COUNTRY
{
	BYTE opCode;
	short iCountryDef;
	short iCountryWinIni;
};

//WINDOW1(3Dh): Workbook level window attributes
struct PXL_WINDOW1
{
	BYTE opCode;
	short grbit;
	short itabCur;
};

#pragma pack(pop)

#endif
