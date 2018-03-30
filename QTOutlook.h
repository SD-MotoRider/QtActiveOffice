#ifndef QTOutlook_H
#define QTOutlook_H

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

#include "qtactiveoffice_global.h"

#include <QList>
#include <QString>
#include <QVariant>

#include "OutlookContact.h"

// outlook.h and outlook.cpp are created with this command
// dumpcpp {00062FFF-0000-0000-C000-000000000046} -o outlook

#include "outlook.h"

class QTACTIVEOFFICE_EXPORT QTOutlook
{
public:
	QTOutlook();
	~QTOutlook();

	bool Open(bool visible = true);
	bool OpenSession(const QString& sessionName = QString(""));

	quint32 GetContacts(OutlookContacts& contacts);

	void Quit(void);

private:
	bool OpenOutlook(bool visible);

	Outlook::Application*		_application;	
	Outlook::_NameSpace*		_session;
	bool						_comInitialized;
};

#endif
