//////////////////////////////////////////////////////////////////////
//
// ItfDB.h: Definition of Class CItfDB
//
//////////////////////////////////////////////////////////////////////

#pragma once

#include <sqlext.h>
#include <sstream>
#include "SQLTools.h"
#include "TypesDef.h"


#define MAX_COL_LENGTH 50
#define MAX_FIELD      75

class CItfDB
{
public:
	CItfDB(void);
	~CItfDB(void);

	long UpdateDB(typRecipe *recipe);

private:

	// SQL Handles 
	SQLHENV		m_hndEnv;
	SQLHDBC		m_hndDBConnect;
	SQLHSTMT  	m_hndSQLState;

	// SQL Operations
	CSQLTools	cSQL;

	// Database connection
	char *m_strODBC_Connection;
	char *m_strDBUser;			
	char *m_strDBPassword;

};
