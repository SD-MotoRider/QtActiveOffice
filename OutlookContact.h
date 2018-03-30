#ifndef OUTLOOKCONTACT_H
#define OUTLOOKCONTACT_H

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

class QTACTIVEOFFICE_EXPORT OutlookContact
{
public:
	OutlookContact() {};
	OutlookContact(const OutlookContact& copyMe)
	{
		_name = copyMe._name;
		_email = copyMe._email;
	}

	~OutlookContact() {};

	QString						_name;
	QString						_email;
};

typedef QList<OutlookContact> OutlookContacts;
typedef QList<OutlookContact>::const_iterator OutlookContactIterator;

#endif //OUTLOOKCONTACT_H

