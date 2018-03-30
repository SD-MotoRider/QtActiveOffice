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

#include "QTOutlook.h"
#include "Hack.h"

#include <QtGlobal>

using namespace Office;
using namespace Outlook;

QTOutlook::QTOutlook() :
	_application(NULL),
	_session(NULL),
	_comInitialized(false)
{
	SetupCom();
}

QTOutlook::~QTOutlook()
{
	if (_application != NULL)
		Quit();
}

bool QTOutlook::Open
(
	bool visible
)
{
	bool result(false);

	if (OpenOutlook(visible))
	{
		result = true;
	}

	return result;
}

bool QTOutlook::OpenSession
(
	const QString& sessionName
)
{
	if (_session != NULL)
	{
		delete _session;
		_session = NULL;
	}

	if (sessionName.size() > 0)
	{
		_session = _application->GetNamespace(sessionName);
	}
	else
	{
		_session = _application->Session();
		if (_session != NULL)
		{
			_session->dynamicCall("Logon()");
		}
	}

	return (_session != NULL);
}

quint32 QTOutlook::GetContacts
(
	OutlookContacts& contacts
)
{	
	Q_ASSERT(_application != NULL);

	contacts.clear();

	Outlook::_NameSpace* nameSpace;

	if (_session != NULL)
		nameSpace = _session;
	else
		nameSpace = _application->GetNamespace("mapi");

	if (nameSpace != NULL)
	{
		Outlook::MAPIFolder* folder = nameSpace->GetDefaultFolder(olFolderContacts);
		if (folder != NULL)
		{ 
			Outlook::Items* items = new Outlook::Items(folder->Items());
			if (items != NULL)
			{
				QAxObject *item = items->querySubObject("GetFirst()");
				while (item != NULL) 
				{
					OutlookContact contact;

					//item->disableMetaObject();

					contact._name = item->property("FullName").toString();
					contact._email = item->property("Email1Address").toString();

					contacts.push_back(contact);




//					FirstName, LastName, HomeAddress, Title, Birthday
//					CompanyName, Department, Body, FileAs, BusinessHomePage
//					MailingAddress, BusinessAddress, OfficeLocation
//					Subject, JobTitle

					delete item;

					item = items->querySubObject("GetNext()");
				}

				delete items;
			}


			delete folder;
		}

		if (_session == NULL)
			delete nameSpace;
	}

	return contacts.size();
}

bool QTOutlook::OpenOutlook
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
		_application = new Outlook::Application;
		if (_application != NULL)
		{
			if (_application->isNull())
			{
				delete _application;
				_application = NULL;
			}
			else
			{
				result = true;
			}
		}
	}
	else
		result = true;

	return result;
}

void QTOutlook::Quit()
{	
	if (_session != NULL)
	{
		delete _session;
		_session = NULL;
	}
	
	if (_application != NULL)
	{
		delete _application;
		_application = NULL;
	}

	if (_comInitialized == true)
	{
		DestroyCom();
		_comInitialized = false;
	}
}
