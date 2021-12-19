#include <iostream>
#include <string>

#include "libxl.h"

using namespace libxl;

int main()
{
    Book* book = xlCreateXMLBook();
    book->setLocale("UTF-8");
    if(book) 
    {   
        Font* boldFont = book->addFont();
        boldFont->setBold();

        Font* titleFont = book->addFont();
        titleFont->setName("Arial Black");
        titleFont->setSize(16);

        Format* titleFormat = book->addFormat();
        titleFormat->setFont(titleFont);

        Format* headerFormat = book->addFormat();
        headerFormat->setAlignH(ALIGNH_CENTER);
        headerFormat->setBorder(BORDERSTYLE_THIN);
        headerFormat->setFont(boldFont);        
        headerFormat->setFillPattern(FILLPATTERN_SOLID);
        headerFormat->setPatternForegroundColor(COLOR_TAN);

        Format* descriptionFormat = book->addFormat();
        descriptionFormat->setBorderLeft(BORDERSTYLE_THIN);

        Format* amountFormat = book->addFormat();
        amountFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
        amountFormat->setBorderLeft(BORDERSTYLE_THIN);
        amountFormat->setBorderRight(BORDERSTYLE_THIN);
                
        Format* totalLabelFormat = book->addFormat();
        totalLabelFormat->setBorderTop(BORDERSTYLE_THIN);
        totalLabelFormat->setAlignH(ALIGNH_RIGHT);
        totalLabelFormat->setFont(boldFont);

        Format* totalFormat = book->addFormat();
        totalFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
        totalFormat->setBorder(BORDERSTYLE_THIN);
        totalFormat->setFont(boldFont);
        totalFormat->setFillPattern(FILLPATTERN_SOLID);
        totalFormat->setPatternForegroundColor(COLOR_YELLOW);

        Format* signatureFormat = book->addFormat();
        signatureFormat->setAlignH(ALIGNH_CENTER);
        signatureFormat->setBorderTop(BORDERSTYLE_THIN);
             
        Sheet* sheet = book->addSheet("Invoice");
        if(sheet)
        {
            sheet->writeStr(2, 1, "Invoice No. 3568", titleFormat);

            sheet->writeStr(4, 1, "Name: John Smith");
            sheet->writeStr(5, 1, "Address: San Ramon, CA 94583 USA");

            sheet->writeStr(7, 1, "Description", headerFormat);
            sheet->writeStr(7, 2, "Amount", headerFormat);

            sheet->writeStr(8, 1, "Ball-Point Pens", descriptionFormat);
            sheet->writeNum(8, 2, 85, amountFormat);
            sheet->writeStr(9, 1, "T-Shirts", descriptionFormat);
            sheet->writeNum(9, 2, 150, amountFormat);
            sheet->writeStr(10, 1, "Tea cups", descriptionFormat);
            sheet->writeNum(10, 2, 45, amountFormat);

            sheet->writeStr(11, 1, "Total:", totalLabelFormat);
            sheet->writeFormula(11, 2, "=SUM(C9:C11)", totalFormat);

            sheet->writeStr(14, 2, "Signature", signatureFormat);

            sheet->setCol(1, 1, 40);
            sheet->setCol(2, 2, 15);
          }
		    Sheet* sheet2 = book->addSheet("Writing a rich string with multiple fonts in one cell");
				if(sheet2)
				{
				    RichString* rs = book->addRichString();

				    Font* font1 = rs->addFont();
				    font1->setColor(COLOR_RED);
				    font1->setSize(24);

				    Font* font2 = rs->addFont();
				    font2->setSize(24);

				    Font* font3 = rs->addFont();
				    font3->setColor(COLOR_BLUE);
				    font3->setSize(24);

				    Font* font4 = rs->addFont();
				    font4->setColor(COLOR_GREEN);
				    font4->setSize(24);

				    Font* font5 = rs->addFont();
				    font5->setScript(SCRIPT_SUPER);
				    font5->setSize(24);

				    rs->addText("E", font1);
				    rs->addText("=", font2);
				    rs->addText("m", font3);
				    rs->addText("c", font4);
				    rs->addText("2", font5);

				    sheet2->writeRichStr(5, 3, rs);
				    sheet2->setProtect(true);
				}
    Format* alFormat = book->addFormat();
    alFormat->setAlignH(ALIGNH_LEFT);

    Format* arFormat = book->addFormat();
    arFormat->setAlignH(ALIGNH_RIGHT);

    Format* alignDateFormat = book->addFormat(alFormat);
    alignDateFormat->setNumFormat(NUMFORMAT_DATE);
    Font* linkFont = book->addFont();
    linkFont->setColor(COLOR_BLUE);
    linkFont->setUnderline(UNDERLINE_SINGLE);
    Format* linkFormat = book->addFormat(alFormat);
    linkFormat->setFont(linkFont);
    Sheet* sheet3 = book->addSheet("Formula");
    if(sheet3)
    {
        sheet3->setCol(0, 0, 27);
        sheet3->setCol(1, 1, 10);

        sheet3->writeNum(2, 1, 40, alFormat);
        sheet3->writeNum(3, 1, 30, alFormat);
        sheet3->writeNum(4, 1, 50, alFormat);

        sheet3->writeStr(6, 0, "SUM(B3:B5) = ", arFormat);
        sheet3->writeFormula(6, 1, "SUM(B3:B5)", alFormat);
        sheet3->writeStr(7, 0, "AVERAGE(B3:B5) = ", arFormat);
        sheet3->writeFormula(7, 1, "AVERAGE(B3:B5)", alFormat);
        sheet3->writeStr(8, 0, "MAX(B3:B5) = ", arFormat);
        sheet3->writeFormula(8, 1, "MAX(B3:B5)", alFormat);
        sheet3->writeStr(9, 0, "MIX(B3:B5) = ", arFormat);
        sheet3->writeFormula(9, 1, "MIN(B3:B5)", alFormat);
        sheet3->writeStr(10, 0, "COUNT(B3:B5) = ", arFormat);
        sheet3->writeFormula(10, 1, "COUNT(B3:B5)", alFormat);

        sheet3->writeStr(12, 0, "IF(B7 > 100;\"large\";\"small\") = ", arFormat);
        sheet3->writeFormula(12, 1, "IF(B7 > 100;\"large\";\"small\")", alFormat);

        sheet3->writeStr(14, 0, "SQRT(25) = ", arFormat);
        sheet3->writeFormula(14, 1, "SQRT(25)", alFormat);
        sheet3->writeStr(15, 0, "RAND() = ", arFormat);
        sheet3->writeFormula(15, 1, "RAND()", alFormat);
        sheet3->writeStr(16, 0, "2*PI() = ", arFormat);
        sheet3->writeFormula(16, 1, "2*PI()", alFormat);

        sheet3->writeStr(18, 0, "UPPER(\"libxl\") = ", arFormat);
        sheet3->writeFormula(18, 1, "UPPER(\"libxl\")", alFormat);
        sheet3->writeStr(19, 0, "LEFT(\"window\";3) = ", arFormat);
        sheet3->writeFormula(19, 1, "LEFT(\"window\";3)", alFormat);
        sheet3->writeStr(20, 0, "LEN(\"string\") = ", arFormat);
        sheet3->writeFormula(20, 1, "LEN(\"string\")", alFormat);

        sheet3->writeStr(22, 0, "DATE(2010;3;11) = ", arFormat);
        sheet3->writeFormula(22, 1, "DATE(2010;3;11)", alignDateFormat);
        sheet3->writeStr(23, 0, "DAY(B23) = ", arFormat);
        sheet3->writeFormula(23, 1, "DAY(B23)", alFormat);
        sheet3->writeStr(24, 0, "MONTH(B23) = ", arFormat);
        sheet3->writeFormula(24, 1, "MONTH(B23)", alFormat);
        sheet3->writeStr(25, 0, "YEAR(B23) = ", arFormat);
        sheet3->writeFormula(25, 1, "YEAR(B23)", alFormat);
        sheet3->writeStr(26, 0, "DAYS360(B23;TODAY()) = ", arFormat);
        sheet3->writeFormula(26, 1, "DAYS360(B23;TODAY())", alFormat);

        sheet3->writeStr(28, 0, "B3+100*(2-COS(0)) = ", arFormat);
        sheet3->writeFormula(28, 1, "B3+100*(2-COS(0))", alFormat);
        sheet3->writeStr(29, 0, "ISNUMBER(B29) = ", arFormat);
        sheet3->writeFormula(29, 1, "ISNUMBER(B29)", alFormat);
        sheet3->writeStr(30, 0, "AND(1;0) = ", arFormat);
        sheet3->writeFormula(30, 1, "AND(1;0)", alFormat);

        sheet3->writeStr(32, 0, "HYPERLINK() = ", arFormat);
        sheet3->writeFormula(32, 1, "HYPERLINK(\"http://www.libxl.com\")", linkFormat);
    }
Sheet* sheet4 = book->addSheet("aligning-colors-borders");
if(sheet4)
{
    sheet4->setDisplayGridlines(false);

    sheet4->setCol(1, 1, 30);
    sheet4->setCol(3, 3, 11.4);
    sheet4->setCol(4, 4, 2);
    sheet4->setCol(5, 5, 15);
    sheet4->setCol(6, 6, 2);
    sheet4->setCol(7, 7, 15.4);

    const char* nameAlignH[] = {"ALIGNH_LEFT", "ALIGNH_CENTER", "ALIGNH_RIGHT"};
    AlignH alignH[] = {ALIGNH_LEFT, ALIGNH_CENTER, ALIGNH_RIGHT};

    for(int i = 0; i < sizeof(nameAlignH) / sizeof(const char*); ++i)
    {
        Format* format = book->addFormat();
        format->setAlignH(alignH[i]);
        format->setBorder();
        sheet4->writeStr(i * 2 + 2, 1, nameAlignH[i], format);
    }

    const char* nameAlignV[] = {"ALIGNV_TOP", "ALIGNV_CENTER", "ALIGNV_BOTTOM"};
    AlignV alignV[] = {ALIGNV_TOP, ALIGNV_CENTER, ALIGNV_BOTTOM};

    for(int i = 0; i < sizeof(nameAlignV) / sizeof(const char*); ++i)
    {
        Format* format = book->addFormat();
        format->setAlignV(alignV[i]);
        format->setBorder();
        sheet4->writeStr(2, i * 2 + 3, nameAlignV[i], format);
        sheet4->setMerge(2, 6, i * 2 + 3, i * 2 + 3);
    }

    const char* nameBorderStyle[] = {"BORDERSTYLE_MEDIUM", "BORDERSTYLE_DASHED",
                                        "BORDERSTYLE_DOTTED", "BORDERSTYLE_THICK",
                                        "BORDERSTYLE_DOUBLE", "BORDERSTYLE_DASHDOT"};
    BorderStyle borderStyle[] = {BORDERSTYLE_MEDIUM, BORDERSTYLE_DASHED, BORDERSTYLE_DOTTED,
                                 BORDERSTYLE_THICK, BORDERSTYLE_DOUBLE, BORDERSTYLE_DASHDOT};

    for(int i = 0; i < sizeof(nameBorderStyle) / sizeof(const char*); ++i)
    {
        Format* format = book->addFormat();
        format->setBorder(borderStyle[i]);
        sheet4->writeStr(i * 2 + 12, 1, nameBorderStyle[i], format);
    }

    const char* nameColors[] = {"COLOR_RED", "COLOR_BLUE", "COLOR_YELLOW",
                                   "COLOR_PINK", "COLOR_GREEN", "COLOR_GRAY25"};
    Color colors[] = {COLOR_RED, COLOR_BLUE, COLOR_YELLOW, COLOR_PINK, COLOR_GREEN,
                      COLOR_GRAY25};
    FillPattern fillPatterns[] = {FILLPATTERN_GRAY50, FILLPATTERN_HORSTRIPE,
                                  FILLPATTERN_VERSTRIPE, FILLPATTERN_REVDIAGSTRIPE,
                                  FILLPATTERN_THINVERSTRIPE, FILLPATTERN_THINHORCROSSHATCH};

    for(int i = 0; i < sizeof(nameColors) / sizeof(const char*); ++i)
    {
        Format* format1 = book->addFormat();
        format1->setFillPattern(FILLPATTERN_SOLID);
        format1->setPatternForegroundColor(colors[i]);
        sheet4->writeBlank(i * 2 + 12, 3, format1);

        Format* format2 = book->addFormat();
        format2->setFillPattern(fillPatterns[i]);
        format2->setPatternForegroundColor(colors[i]);
        sheet4->writeBlank(i * 2 + 12, 5, format2);

        Font* font = book->addFont();
        font->setColor(colors[i]);
        Format* format3 = book->addFormat();
        format3->setBorder();
        format3->setBorderColor(colors[i]);
         	book->setRgbMode(true); //rgb
 					format3->setBorderColor(book->colorPack(255,0,0));
        format3->setFont(font);
        sheet4->writeStr(i * 2 + 12, 7, nameColors[i], format3);
    }
}
    Sheet* sheet5 = book->addSheet("Customizing fonts");
if(sheet5)
{
    const char* fonts[] = {"Aria", "Arial Black", "Comic Sans MS", "Courier New",
                              "Impact", "Times New Roman", "Verdana"};

    for(int i = 0; i < sizeof(fonts) / sizeof(const char*); ++i)
    {
        Font* font = book->addFont();
        font->setSize(16);
        font->setName(fonts[i]);
        Format* format = book->addFormat();
        format->setFont(font);
        sheet5->writeStr(i + 2, 3, fonts[i], format);
    }

    int fontSize[] = {8, 10, 12, 14, 16, 20, 25};

    for(int i = 0; i < sizeof(fontSize) / sizeof(int); ++i)
    {
        Font* font = book->addFont();
        font->setSize(fontSize[i]);
        Format* format = book->addFormat();
        format->setFont(font);
        sheet5->writeStr(i + 2, 7, "Text", format);
    }

    Font* font = book->addFont();
    font->setSize(16);
    Format* format = book->addFormat();
    format->setRotation(255);
    format->setFont(font);
    sheet5->writeStr(2, 9, "Vertica", format);
    sheet5->setMerge(2, 8, 9, 9);

    Font* boldFont = book->addFont();
    boldFont->setBold();
    Format* boldFormat = book->addFormat();
    boldFormat->setFont(boldFont);

    Font* italicFont = book->addFont();
    italicFont->setItalic();
    Format* italicFormat = book->addFormat();
    italicFormat->setFont(italicFont);

    Font* underlineFont = book->addFont();
    underlineFont->setUnderline(UNDERLINE_SINGLE);
    Format* underlineFormat = book->addFormat();
    underlineFormat->setFont(underlineFont);

    Font* strikeoutFont = book->addFont();
    strikeoutFont->setStrikeOut();
    Format* strikeoutFormat = book->addFormat();
    strikeoutFormat->setFont(strikeoutFont);

    sheet5->writeStr(2, 1, "Norma");
    sheet5->writeStr(3, 1, "Bold", boldFormat);
    sheet5->writeStr(4, 1, "Italic", italicFormat);
    sheet5->writeStr(5, 1, "Underline", underlineFormat);
    sheet5->writeStr(6, 1, "Strikeout", strikeoutFormat);
}

Sheet* sheet6 = book->addSheet("sorting");
if(sheet6)
{
    sheet6->writeStr(2, 1, "Country");
    sheet6->writeStr(2, 2, "Road injures");
    sheet6->writeStr(2, 3, "Smoking");
    sheet6->writeStr(2, 4, "Suicide");

    sheet6->writeStr(3, 1, "USA");     sheet6->writeStr(4, 1, "UK");
    sheet6->writeNum(3, 2, 64);         sheet6->writeNum(4, 2, 94);
    sheet6->writeNum(3, 3, 69);         sheet6->writeNum(4, 3, 55);
    sheet6->writeNum(3, 4, 49);         sheet6->writeNum(4, 4, 64);

    sheet6->writeStr(5, 1, "Germany"); sheet6->writeStr(6, 1, "Switzerland");
    sheet6->writeNum(5, 2, 88);         sheet6->writeNum(6, 2, 93);
    sheet6->writeNum(5, 3, 46);         sheet6->writeNum(6, 3, 54);
    sheet6->writeNum(5, 4, 55);         sheet6->writeNum(6, 4, 50);

    sheet6->writeStr(7, 1, "Spain");   sheet6->writeStr(8, 1, "Italy");
    sheet6->writeNum(7, 2, 86);         sheet6->writeNum(8, 2, 75);
    sheet6->writeNum(7, 3, 47);         sheet6->writeNum(8, 3, 52);
    sheet6->writeNum(7, 4, 69);         sheet6->writeNum(8, 4, 71);
       Format* formatA = book->addFormat();
        formatA->setBorder();
         	book->setRgbMode(true); //rgb
         	formatA->setBorderColor(book->colorPack(255,0,0));
        formatA->setBorderLeft(BORDERSTYLE_THIN);
        formatA->setBorderRight(BORDERSTYLE_THIN);
    sheet6->writeStr(9, 1, "Greece");  sheet6->writeStr(10, 1, "+-Japan",formatA);
    sheet6->writeNum(9, 2, 67);         sheet6->writeNum(10, 2, 91);
    sheet6->writeNum(9, 3, 23);         sheet6->writeNum(10, 3, 57);
    sheet6->writeNum(9, 4, 87);         sheet6->writeNum(10, 4, 36);

       int id = book->addPicture("img/heart.jpg");
        if(id == -1)
        {
            std::cout << "picture not found" << std::endl;
            return -1;
        }
    sheet6->setPicture(20, 1, id);


    AutoFilter* autoFilter = sheet6->autoFilter();
    autoFilter->setRef(2, 10, 1, 4);

    // sorting by road injures

    autoFilter->setSort(1, true);

    sheet6->applyFilter();

    // shows all countries with the first letter G

    /*FilterColumn* column = autoFilter->column(0);
    column->setCustomFilter(OPERATOR_EQUAL, "G*");
    sheet6->applyFilter();*/
}

//replace

    for(int i = 5; i < book->sheetCount(); ++i)
    {
        Sheet* sheetA = book->getSheet(i);

        for(int row = sheetA->firstRow(); row < sheetA->lastRow(); ++row)
        {
            for(int col = sheetA->firstCol(); col < sheetA->lastCol(); ++col)
            {
                if(sheetA->cellType(row, col) == CELLTYPE_STRING)
                {
                    const char* s = sheetA->readStr(row, col);
                    if(s)
                    {
                        std::string str(s);
                        std::string::size_type pos = str.find("+-");
                        if(pos != std::string::npos)
                        {
                            str.replace(pos, 2, "±");
                        }
                        sheetA->writeStr(row, col, str.c_str());
                    }
                }
            }
        }
    }
          if(book->save("invoice.xlsx")) {
              std::cout << "File invoice.xlsx has been created." << std::endl;
          }
          book->release();
    }

    return 0;
}
