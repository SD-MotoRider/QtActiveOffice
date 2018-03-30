#ifndef QTWORD_H
#define QTWORD_H

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

// word.h and word.cpp are created with this command
// Office 2003 use, "dumpcpp {000209FF-0000-0000-C000-000000000046} -o word"
// Office 2016 use, "dumpcpp {00020905-0000-0000-C000-000000000046} -o word"

#include "Word.h"

class QTACTIVEOFFICE_EXPORT QTWord
{
public:
	QTWord();
	~QTWord();

	bool New(bool visible = true);
	bool Open(const QString& docFile, bool visible = true);
	void ScreenUpdating(bool updateScreen);
	void DisplayAlerts(bool displayAlerts);

	void Quit(void);

private:
	bool OpenWord(bool visible);

	Word::Application*			_application;	
	bool						_comInitialized;
};

#endif
