#pragma once

#include <QtCore/qglobal.h>

#ifndef BUILD_STATIC
# if defined(QTACTIVEOFFICE_LIB)
#  define QTACTIVEOFFICE_EXPORT Q_DECL_EXPORT
# else
#  define QTACTIVEOFFICE_EXPORT Q_DECL_IMPORT
# endif
#else
# define QTACTIVEOFFICE_EXPORT
#endif
