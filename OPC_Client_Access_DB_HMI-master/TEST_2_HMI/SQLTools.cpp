// SQLTools.cpp: Implementation of CSQLTools Class
//

#include "stdafx.h"
#include "SQLTools.h"


//////////////////////////////////////////////////////////////////////
// Constructor/Destructor
//////////////////////////////////////////////////////////////////////

CSQLTools::CSQLTools()
{
}


CSQLTools::~CSQLTools()
{
}

//****************************************************************************************
BSTR CSQLTools::SQLConnectDB( SQLCHAR *ServerName,SQLCHAR *UserName, SQLCHAR *PW)
{
	int ret =0;

		// Environment handle allokieren
		ret = SQLAllocEnv(&m_hndEnv);
		if ((ret!=SQL_SUCCESS) && (ret != SQL_SUCCESS_WITH_INFO))
			return 	SQLGetError(ret,SQL_HANDLE_ENV);

		// Connection handle allokieren       	
		ret = SQLAllocConnect(m_hndEnv, &m_hndDBConnect);
		if ((ret!=SQL_SUCCESS) && (ret != SQL_SUCCESS_WITH_INFO))
			return 	SQLGetError(ret,SQL_HANDLE_ENV);

		// Verbindung zur Datenbank aufbauen
		ret =  SQLConnect(m_hndDBConnect, ServerName , SQL_NTS, UserName , SQL_NTS, PW , SQL_NTS);
		if ((ret!=SQL_SUCCESS) && (ret != SQL_SUCCESS_WITH_INFO))
			return SQLGetError(ret,SQL_HANDLE_DBC);

		// Statement handle allokieren
		ret =  SQLAllocHandle(SQL_HANDLE_STMT, m_hndDBConnect, &m_hndSQLState);
		if ((ret!=SQL_SUCCESS) && (ret != SQL_SUCCESS_WITH_INFO))
			return SQLGetError(ret,SQL_HANDLE_STMT);

		return NULL;
}

//****************************************************************************************

BSTR CSQLTools::SQLSendUpdate(SQLCHAR *StatementText)
{
	SQLRETURN ret = 0;

	// SQL-Statement an Datenbank absetzen
	ret = SQLExecDirect(m_hndSQLState,StatementText,SQL_NTS);
		if (ret!=SQL_SUCCESS)
			return SQLGetError(ret,SQL_HANDLE_STMT);

	return NULL;
}

//****************************************************************************************

BSTR CSQLTools::SQLFree()

{
		//Free all Database resources
		int	ret=0;

		ret = SQLFreeHandle(SQL_HANDLE_STMT,m_hndSQLState); 
		if (ret!=SQL_SUCCESS)	
			return(SQLGetError(ret,SQL_HANDLE_STMT));

		ret = SQLDisconnect(m_hndDBConnect);		
		if (ret!=SQL_SUCCESS)	
			return(SQLGetError(ret,SQL_HANDLE_DBC));

		ret = SQLFreeHandle(SQL_HANDLE_DBC,m_hndDBConnect); 
		if (ret!=SQL_SUCCESS)	
			return(SQLGetError(ret,SQL_HANDLE_DBC));

		ret = SQLFreeHandle(SQL_HANDLE_ENV,m_hndEnv); 	
		if (ret!=SQL_SUCCESS)
			return(SQLGetError(ret,SQL_HANDLE_ENV));

		return NULL;
}

//****************************************************************************************

BSTR CSQLTools::SQLGetError(int retcode,int handletype )
{

static int		envCounter=1;
static int		dbcCounter=1;
static int		stmtCounter=-1;	

SQLCHAR			SqlState[6],sErrorMsg[SQL_MAX_MESSAGE_LENGTH];
SQLINTEGER		NativeError;
SQLSMALLINT		ErrorMsgLen=0;
BSTR			BStrError;
OLECHAR			wzError[SQL_MAX_MESSAGE_LENGTH];

	switch(handletype)
	{
		case SQL_HANDLE_ENV:
			SQLGetDiagRec(SQL_HANDLE_ENV, m_hndEnv, envCounter++, SqlState, &NativeError, sErrorMsg, sizeof(sErrorMsg), &ErrorMsgLen) ;
			if(ErrorMsgLen > 0)     // uerspruenglich: if (sErrMsg), das war natuerlich
                                    // ein Fehler, weil sErrMsg const != 0 ist!
				MultiByteToWideChar(CP_ACP,0,(const char *)sErrorMsg,-1,wzError, ErrorMsgLen);
            else
                wcscpy_s(wzError,L"Unknown SQL_HANDLE_ENV Error.");
			break;

        case SQL_HANDLE_DBC:
			SQLGetDiagRec(SQL_HANDLE_DBC, m_hndDBConnect, dbcCounter++, SqlState, &NativeError, sErrorMsg, sizeof(sErrorMsg), &ErrorMsgLen) ;
			if(ErrorMsgLen > 0)
				MultiByteToWideChar(CP_ACP,0,(const char *)sErrorMsg,-1,wzError, ErrorMsgLen);
            else
                wcscpy_s(wzError,L"Unknown SQL_HANDLE_DBC Error.");
    		break;

		case SQL_HANDLE_STMT:
			SQLGetDiagField(SQL_HANDLE_STMT, m_hndSQLState, 0, SQL_DIAG_NUMBER, &stmtCounter, SQL_IS_INTEGER, NULL);
			// aus unerfindlichen Gruenden kommt hier bisweilen stmtCounter==0 zurueck
			// SQLGetDiagRec() liefert dann nichts Brauchbares ==> abfangen
			if (stmtCounter<1)
			{
				wcscpy_s(wzError,L"SQLGetDiag: wrong error count");
				break;
			}

			SQLGetDiagRec(SQL_HANDLE_STMT, m_hndSQLState,stmtCounter++ , SqlState, &NativeError, sErrorMsg, sizeof(sErrorMsg), &ErrorMsgLen); 
			
			if(ErrorMsgLen > 0)
				MultiByteToWideChar(CP_ACP,0,(const char *)sErrorMsg,-1,wzError, ErrorMsgLen);
            else
                wcscpy_s(wzError,L"Unknown SQL_HANDLE_STMT Error.");
			break;

        default:
            wcscpy_s(wzError,L"Unknown Error SQL_HANDLE.");
            break;

	}

	BStrError = SysAllocString(wzError);
	return BStrError;
}