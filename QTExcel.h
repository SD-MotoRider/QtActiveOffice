#ifndef QTEXCEL_H
#define QTEXCEL_H

//The MIT License
//
//Copyright (c) 2008-2018 Michael Simpson
//
//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in
//all copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
//THE SOFTWARE.

#include "qtactiveoffice_global.h"

#include <QList>
#include <QString>
#include <QVariant>

// Excel.h and Excel.cpp are created with this command
// dumpcpp {00020813-0000-0000-C000-000000000046} -o excel
// run dumpcpp from a QT environment command line in the generated folders directory
#include "Excel.h"

// Excel's borders can't be used as a mask
// Creating our own for efficiency
const quint32 kInsideHorizontal(1);
const quint32 kInsideVertical(2);
const quint32 kDiagonalDown(4);
const quint32 kDiagonalUp(8);
const quint32 kEdgeBottom(16);
const quint32 kEdgeLeft(32);
const quint32 kEdgeRight(64);
const quint32 kEdgeTop(128);
const quint32 kLastBorderMask(128);

class QTACTIVEOFFICE_EXPORT QTExcel
{
public:
	QTExcel();
	~QTExcel();

	bool New(bool visible = true);
	bool Open(const QString& excelXLSFile, bool visible = true);
	void ScreenUpdating(bool updateScreen);
	void DisplayAlerts(bool displayAlerts);

	void Quit(void);

	QStringList WorkBookNames(void);
	quint32 WorkBookCount(void);
	bool SelectWorkbook(quint32 index);
	bool SelectWorkbook(const QString& sheetName);
	bool ActivateWorkbook(void);

	QStringList WorkSheetNames(void);
	quint32 WorkSheetCount(void);
	bool SelectWorksheet(quint32 index);
	bool SelectWorksheet(const QString& sheetName);
	bool ActivateWorksheet(void);
	bool RenameSheet(const QString& newName);
	void SetWorksheetGridLines(bool showGridlines);

	QVariant GetCellValue(const QString& cellReference);
	quint32 GetCellValues(const QString& cellReference, QList<QVariant>& results);

	void SetCellValue(const QString& cellReference, const QString& cellValue);
	void SetCellValue(const QString& cellReference, const QVariant& cellValue);

	void SetCellValues(const QString& cellReference, const QStringList& cellValues);
	void SetCellValues(const QString& cellReference, const QList<QVariant>& cellValues);

	QColor GetCellInteriorColor(const QString& cellReference);

	// Setting Fonts
	void SetCellBold(const QString& cellReference, bool bold);
	void SetCellFontSize(const QString& cellReference, int fontSize);
	void SetCellFontColor(const QString& cellReference, const QColor& color);

	// Setting Formats
	void AutoFitColumns(const QString& cellReference);
	void SetCellInteriorColor(const QString& cellReference, const QColor& color);
	void SetCellHorizontalAlignment(const QString& cellReference, Excel::XlHAlign horizontalAlignment);
	void SetCellVerticalAlignment(const QString& cellReference, Excel::XlVAlign verticalAlignment);

	// Setting Borders
	void SetCellBorderStyle(const QString& cellReference, Excel::XlLineStyle lineStyle,
		Excel::XlBorderWeight lineWeight, const QColor& color);
	void SetCellBorderElement(const QString& cellReference, quint32 borderMask, Excel::XlLineStyle lineStyle,
		Excel::XlBorderWeight lineWeight, const QColor& color);

private:
	bool OpenExcel(bool visible);

	QVariant GetCellValue(Excel::Range* range, quint32 row, quint32 col);
	void SetCellValue(Excel::Range* range, quint32 row, quint32 col, const QVariant& value);

	Excel::Application*			_application;
	Excel::Workbook*			_workbook;
	Excel::Worksheet*			_workSheet;		
	bool						_comInitialized;
};

#endif
