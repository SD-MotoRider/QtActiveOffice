//The MIT License
//
//Copyright (c) 2008 Michael Simpson
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

#include "QTWord.h"
#include "Hack.h"

#include <QtGlobal>

using namespace Office;
using namespace Word;

QTWord::QTWord() :
	_application(NULL),
	_comInitialized(false)
{
	SetupCom();
}

QTWord::~QTWord()
{
	if (_application != NULL)
	{
		_application->Quit();

		delete _application;
		_application = NULL;
	}

	Q_ASSERT_X(_comInitialized == false, "QTWord::~QTWord", "QTWord::Quit not called");
}

bool QTWord::New
(
	bool visible
)
{
	bool result(false);

	if (OpenWord(visible))
	{
/*		_application->NewDocument()->Add(
		if (wordFile != NULL)
		{
			wordFile->Add
		}
*/
	}

	return result;
}

bool QTWord::Open
(
	const QString& excelXLSFile,
	bool visible
)
{
	bool result(false);

	if (OpenWord(visible))
	{
	}

	return result;
}

bool QTWord::OpenWord
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
		_application = new Word::Application;
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

void QTWord::Quit()
{
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
