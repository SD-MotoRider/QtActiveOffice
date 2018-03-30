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

#include "QTExcel.h"
#include "Hack.h"

#include <QtGlobal>

using namespace Excel;

QTExcel::QTExcel() :
	_application(NULL),
	_workbook(NULL),
	_workSheet(NULL),
	_comInitialized(false)
{
}

QTExcel::~QTExcel()
{
	if (_workSheet != NULL)
	{
		delete _workSheet;
		_workSheet = NULL;
	}
	if (_workbook != NULL)
	{
		delete _workbook;
		_workbook = NULL;
	}
	if (_application != NULL)
	{
		_application->Quit();

		delete _application;
		_application = NULL;
	}

	Q_ASSERT_X(_comInitialized == false, "QTExcel::~QTExcel", "QTExcel::Quit not called");
}

bool QTExcel::New
(
	bool visible
)
{
	bool result(false);

	if (OpenExcel(visible))
	{
		_workbook = _application->Workbooks()->Add();
		if (_workbook != NULL)
		{
			result = true;
		}
	}

	return result;
}

bool QTExcel::Open
(
	const QString& excelXLSFile,
	bool visible
)
{
	bool result(false);

	if (OpenExcel(visible))
	{
		_workbook = _application->Workbooks()->Open(excelXLSFile);
		if (_workbook != NULL)
		{
			_workSheet = new Excel::Worksheet(new _Worksheet(_workbook->Sheets()->Item(QVariant(1))));

			result = true;
		}
	}

	return result;
}

bool QTExcel::OpenExcel
(
	bool visible
)
{
	bool result(false);

	if (_comInitialized == false)
	{
		SetupCom();
		_comInitialized = true;
	}

	if (_application == NULL)
	{
		_application = new Excel::Application;
		if (_application != NULL)
		{
			if (_application->isNull())
			{
				delete _application;
			}
			else
			{
				_application->SetVisible(visible);
				result = true;
			}
		}
	}
	else
		result = true;

	return result;
}

void QTExcel::Quit()
{
	if (_workbook != NULL)
	{
		delete _workbook;
		_workbook = NULL;
	}

	if (_workSheet != NULL)
	{
		delete _workSheet;
		_workSheet = NULL;
	}

	if (_application != NULL)
	{
		_application->Quit();
		delete _application;
		_application = NULL;
	}

	if (_comInitialized == true)
	{
		DestroyCom();
		_comInitialized = false;
	}
}

void QTExcel::ScreenUpdating(bool updateScreen)
{

	if (_application != NULL)
	{
		_application->SetScreenUpdating(updateScreen);
	}
}

void QTExcel::DisplayAlerts(bool displayAlerts)
{
	if (_application != NULL)
	{
		_application->SetDisplayAlerts(displayAlerts);
	}
}

bool QTExcel::ActivateWorkbook()
{
	bool result(false);

	if (_workbook != NULL)
	{
		_workbook->Activate();
		result = true;
	}

	return result;
}

QStringList QTExcel::WorkBookNames()
{	
	QStringList result;
	
	if (_application != NULL)
	{
		Excel::Workbooks* workbooks = _application->Workbooks();
		if (workbooks != NULL)
		{
			quint32 size = workbooks->Count();

			for (quint32 index = 1; index <= size; index++)
			{
				Excel::Workbook* workbook = workbooks->Item(QVariant(index));
				if (workbook != NULL)
				{
					QString name = workbook->Name();
					
					result.append(name);
				}
			}

			delete workbooks;
		}
	}

	return result;
}

quint32 QTExcel::WorkBookCount()
{	
	quint32 result(0);

	if (_application != NULL)
	{
		Excel::Workbooks* workbooks = _application->Workbooks();
		if (workbooks != NULL)
		{
			result = workbooks->Count();
			delete workbooks;
		}
	}

	return result;
}

bool QTExcel::SelectWorkbook
(
	quint32 index
)
{
	bool result(false);

	if (_application != NULL)
	{
		Excel::Workbooks* workbooks = _application->Workbooks();
		if (workbooks != NULL)
		{
			Excel::Workbook* workbook = workbooks->Item(QVariant(index));
			if (workbook != NULL)
			{
				if (_workbook != NULL)
					delete _workbook;

				_workbook = workbook;
				result = true;
			}

			delete workbooks;
		}
	}

	return result;
}

bool QTExcel::SelectWorkbook
(
	const QString& sheetName
)
{
	bool result(false);

	if (_application != NULL)
	{
		Excel::Workbooks* workbooks = _application->Workbooks();
		if (workbooks != NULL)
		{
			Excel::Workbook* workbook = workbooks->Item(QVariant(sheetName));
			if (workbook != NULL)
			{
				if (_workbook != NULL)
					delete _workbook;

				_workbook = workbook;
				result = true;
			}

			delete workbooks;
		}
	}

	return result;
}

bool QTExcel::ActivateWorksheet()
{	
	bool result(false);

	if (_workSheet != NULL)
	{
		_workSheet->Activate();
		result = true;
	}

	return result;
}

bool QTExcel::RenameSheet
(
	const QString& newName
)
{	
	bool result(false);

	if (_workSheet != NULL)
	{
		_workSheet->SetName(newName);
		result = true;
	}

	return result;
}

QStringList QTExcel::WorkSheetNames()
{
	QStringList result;

	if (_workbook != NULL)
	{
		Excel::Sheets* sheets = _workbook->Sheets();
		if (sheets != NULL)
		{
			quint32 sheetCount = sheets->Count();
			for (quint32 index = 1; index <= sheetCount; index++)
			{
				IDispatch* sheetItem = sheets->Item(QVariant(index));
				if (sheetItem != NULL)
				{
					Excel::Worksheet* workSheet = new Excel::Worksheet(new Excel::_Worksheet(sheetItem));
					if (workSheet != NULL)
					{
						result.append(workSheet->Name());
						delete workSheet;
					}
				}
			}

			delete sheets;
		}
	}

	return result;
}

quint32 QTExcel::WorkSheetCount()
{
	quint32 result(0);

	if (_workbook != NULL)
	{
		Excel::Sheets* sheets = _workbook->Sheets();
		if (sheets != NULL)
		{
			result = sheets->Count();
			
			delete sheets;
		}
	}

	return result;
}

bool QTExcel::SelectWorksheet
(
	quint32 index
)
{
	bool result(false);

	if (_workbook != NULL)
	{
		Excel::Sheets* sheets = _workbook->Sheets();
		if (sheets != NULL)
		{
			quint32 sheetCount = sheets->Count();
			
			if (index <= sheetCount)
			{
				IDispatch* sheetItem = sheets->Item(QVariant(index));
				if (sheetItem != NULL)
				{
					Excel::Worksheet* workSheet = new Excel::Worksheet(new Excel::_Worksheet(sheetItem));
					if (workSheet != NULL)
					{
						if (_workSheet != NULL)
							delete _workSheet;

						_workSheet = workSheet;

						return true;
					}
				}
			}

			delete sheets;
		}
	}

	return false;
}

bool QTExcel::SelectWorksheet
(
	const QString& sheetName
)
{
	bool result(false);

	if (_workbook != NULL)
	{
		Excel::Sheets* sheets = _workbook->Sheets();
		if (sheets != NULL)
		{
			quint32 sheetCount = sheets->Count();
			for (quint32 index = 1; index <= sheetCount; index++)
			{
				IDispatch* sheetItem = sheets->Item(index);
				if (sheetItem != NULL)
				{
					Excel::Worksheet* workSheet = new Excel::Worksheet(new Excel::_Worksheet(sheetItem));
					if (workSheet != NULL)
					{
						if (workSheet->Name() == sheetName)
						{
							if (_workSheet != NULL)
								delete _workSheet;

							_workSheet = workSheet;

							return true;
						}
					}
				}
			}

			delete sheets;
		}
	}

	return result;
}

void QTExcel::SetWorksheetGridLines
(
	bool showGridlines
)
{
	if (_application != NULL)
	{
		Excel::Window* activeWindow = _application->ActiveWindow();
		if (activeWindow != NULL)
		{
			activeWindow->SetDisplayGridlines(showGridlines);
			delete activeWindow;
		}
	}
}

QVariant QTExcel::GetCellValue
(
	const QString& cellReference
)
{
	QVariant result;

	Excel::Range* range = _workSheet->Range(QVariant(cellReference));
	if (range != NULL)
	{
		result = GetCellValue(range, 1, 1);

		delete range;
	}

	return result;
}

quint32 QTExcel::GetCellValues
(
	const QString& cellReference,
	QList<QVariant>& results
)
{
	results.clear();

	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			quint32 rowBounds = range->Rows()->Count();
			quint32 colBounds = range->Columns()->Count();
			if (rowBounds == 1 &&  colBounds == 1)
			{
				results.push_back(GetCellValue(range, 1, 1));
			}
			else
			{
				for (quint32 row = 1; row <= rowBounds; row++)
				{
					for (quint32 col = 1; col <= colBounds; col++)
					{
						results.push_back(GetCellValue(range, row, col));
					}
				}
			}

			delete range;
		}
	}

	return results.size();
}

QVariant QTExcel::GetCellValue
(
	Excel::Range* range, 
	quint32 row, 
	quint32 col
)
{
	QVariant result;

	QVariant item = range->Item(row, col);

	if (item.isValid())
	{
		QAxObject* cellObject = item.value<QAxObject*>();
		if (cellObject != NULL)
		{
			cellObject->disableMetaObject();
			result = cellObject->dynamicCall("Value");

			if (!result.isValid())
				result = QString("");

			delete cellObject;
		}
	}

	return result;
}

void QTExcel::SetCellValue
(
	const QString& cellReference, 
	const QString& cellValue
)
{
	SetCellValue(cellReference, QVariant(cellValue));
}

void QTExcel::SetCellValue
(
	const QString& cellReference, 
	const QVariant& cellValue
)
{
	Excel::Range* range = _workSheet->Range(QVariant(cellReference));
	if (range != NULL)
	{
		SetCellValue(range, 1, 1, cellValue);

		delete range;
	}
}

void QTExcel::SetCellValues
(
	const QString& cellReference, 
	const QStringList& cellValues
)
{
	QList<QVariant> cellValuesV;

	QList<QString>::const_iterator cellValue = cellValues.begin();

	while (cellValue != cellValues.end())
	{
		cellValuesV.push_back(QVariant(*cellValue));
		cellValue++;
	}

	SetCellValues(cellReference, cellValuesV);
}

void QTExcel::SetCellValues
(
	const QString& cellReference, 
	const QList<QVariant>& cellValues
)
{
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			quint32 rowBounds = range->Rows()->Count();
			quint32 colBounds = range->Columns()->Count();

			QList<QVariant>::const_iterator cellValue = cellValues.begin();

			while (cellValue != cellValues.end())
			{
				if (rowBounds == 1 &&  colBounds == 1)
				{
					SetCellValue(range, 1, 1, *cellValue);
					cellValue = cellValues.end();
				}
				else
				{
					quint32 row = 1;
					quint32 col = 1;

					while (cellValue != cellValues.end())
					{
						SetCellValue(range, row, col, *cellValue);

						row++;
						if (row > rowBounds)
						{
							row = 1;
							col++;
						}

						if (col > colBounds)
						{
							cellValue = cellValues.end();
						}
						else
							cellValue++;
					}
				}
			}

			delete range;
		}
	}
}

void QTExcel::SetCellValue
(
	Excel::Range* range, 
	quint32 row, 
	quint32 col, 
	const QVariant& value
)
{
	QVariant item = range->Item(row, col);

	if (item.isValid())
	{
		QAxObject* cellObject = item.value<QAxObject*>();
		if (cellObject != NULL)
		{
			cellObject->disableMetaObject();
			cellObject->dynamicCall("Value", value);

			delete cellObject;
		}
	}
}

QColor QTExcel::GetCellInteriorColor
(
	const QString& cellReference
)
{
	QColor result(Qt::black);

	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			result = range->Interior()->Color().value<QColor>();
			delete range;
		}
	}

	return result;
}

void QTExcel::AutoFitColumns
(
	const QString& cellReference
)
{	
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			Excel::Range* columns = range->EntireColumn();
			if (columns != NULL)
			{
				columns->AutoFit();
				delete columns;
			}

			delete range;
		}
	}
}

void QTExcel::SetCellBold
(
	const QString& cellReference, 
	bool bold
)
{	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			Excel::Font* font = range->Font();
			if (font != NULL)
			{
				font->SetBold(bold);

				delete font;
			}

			delete range;
		}
	}
}

void QTExcel::SetCellFontSize
(
	const QString& cellReference, 
	int fontSize
)
{	
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			Excel::Font* font = range->Font();
			if (font != NULL)
			{
				font->SetSize(fontSize);

				delete font;
			}

			delete range;
		}
	}
}

void QTExcel::SetCellFontColor
(
	const QString& cellReference, 
	const QColor& color
)
{
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			Excel::Font* font = range->Font();
			if (font != NULL)
			{
				font->SetColor(color);

				delete font;
			}

			delete range;
		}
	}
}

void QTExcel::SetCellInteriorColor
(
	const QString& cellReference, 
	const QColor& color
)
{
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			range->Interior()->SetColor(color);
			delete range;
		}
	}
}

void QTExcel::SetCellBorderStyle
(
	const QString& cellReference, 
	Excel::XlLineStyle lineStyle,
	Excel::XlBorderWeight lineWeight,
	const QColor& color
)
{	
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			range->BorderAround(QVariant(lineStyle), lineWeight, xlColorIndexNone, QVariant(color));

			delete range;
		}
	}
}

void QTExcel::SetCellBorderElement
(
	const QString& cellReference, 
	quint32 borderMask, 
	Excel::XlLineStyle lineStyle,
	Excel::XlBorderWeight lineWeight, 
	const QColor& color
)
{	
	int borders(borderMask);
	int mask = 1;

	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			Excel::XlBordersIndex borderIndex;
			
			Excel::Borders* borders = range->Borders();
			if (borders != NULL)
			{
				while (mask <= kLastBorderMask)
				{  
					if (mask & borderMask)
					{
						switch (mask)
						{
							case kInsideHorizontal: borderIndex = xlInsideHorizontal; break;
							case kInsideVertical: borderIndex = xlInsideVertical;break;
							case kDiagonalDown: borderIndex = xlDiagonalDown; break;
							case kDiagonalUp: borderIndex = xlDiagonalUp; break;
							case kEdgeBottom: borderIndex = xlEdgeBottom; break;
							case kEdgeLeft: borderIndex = xlEdgeLeft; break;
							case kEdgeRight: borderIndex = xlEdgeRight; break;
							case kEdgeTop: borderIndex = xlEdgeTop; break;
							default: borderIndex = (Excel::XlBordersIndex) 0; break;
						}

						if (borderIndex != (Excel::XlBordersIndex) 0) // we have a good border
						{
							Excel::Border* border;

							border = borders->Item(borderIndex);
							if (border != NULL)
							{
								border->SetLineStyle(lineStyle);
								border->SetWeight(lineWeight);
								border->SetColor(color);

								delete border;
							}
						}
					}

					mask = mask << 1;
				}

				delete borders;
			}

			delete range;
		}
	}
}

void QTExcel::SetCellHorizontalAlignment
(
	const QString& cellReference, 
	Excel::XlHAlign horizontalAlignment
)
{	
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			range->SetHorizontalAlignment(horizontalAlignment);

			delete range;
		}
	}
}

void QTExcel::SetCellVerticalAlignment
(
	const QString& cellReference, 
	Excel::XlVAlign verticalAlignment
)
{	
	if (_workSheet != NULL)
	{
		Excel::Range* range = _workSheet->Range(QVariant(cellReference));
		if (range != NULL)
		{
			range->SetVerticalAlignment(verticalAlignment);

			delete range;
		}
	}
}