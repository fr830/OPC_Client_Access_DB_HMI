//////////////////////////////////////////////////////////////////////
//
// SQLTools.h: Definition of Class CSQLTools
//
//////////////////////////////////////////////////////////////////////

#pragma once

#include <sqlext.h>
#include <stdio.h>
#include <time.h>  

class CSQLTools  
{
public:

	CSQLTools();
	virtual ~CSQLTools();

	BSTR	SQLConnectDB(SQLCHAR *ServerName,SQLCHAR *UserName, SQLCHAR *Authentication);
	BSTR	SQLSendUpdate(SQLCHAR *StatementText);	
	BSTR	SQLFree();
	BSTR 	SQLGetError(int retcode,int handletype);

private:
	//SQL Handles 
	SQLHENV		m_hndEnv;
	SQLHDBC		m_hndDBConnect;
	SQLHSTMT  	m_hndSQLState;

};

