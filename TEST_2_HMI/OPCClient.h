//////////////////////////////////////////////////////////////////////
//
// OPCClient.h: Definition of Class COPCClient
//
//////////////////////////////////////////////////////////////////////

#pragma once

#include "opcerror.h"
#include "opc.h"
#include <comutil.h>
#include <iostream>
#include <fstream>
#include "TypesDef.h"
#include "ItfDB.h"

#pragma comment(lib, "comsuppw.lib")

using namespace std;

#define READ_ITEMS   10
#define WRITE_ITEMS  2

// Used by GetErrorString; Language code 
#define LOCALE_ID    0x409	//  ENGLISH


class COPCClient
{
public:
	COPCClient(void);
	~COPCClient(void);
	long Connect(long lngInterval);
	long ReadFromPLC(void);
	void loop();
	long WriteToPLC(void);
	//long ResetData(void);///////////////////////////////////////////////////////TODEL///////
	long Close(void);
	char *getPLCStatusMessage(void);
	void SetRecipe(typRecipe *pRecipe);


private:
	//PLCThread
	HANDLE			m_hPLCLoop;
	DWORD			m_dwPLCLoop;
	void			PLCThreadWorker();
	static void		PLCThread(void *pvParameter)
	{
		(((COPCClient*)pvParameter)->PLCThreadWorker)();
	}
	bool			m_bIsRunning;
	bool			m_bLoopExit;

	//Critical section for the protection of the thread
	CComCriticalSection	m_csSecure;

	//PLC Connection
	bool			m_bPLCConnectionFailed;
	long			m_lngInterval;
	char		   *m_strStatusMessage;

	//OPC Process
	IOPCServer		*m_pIOPCServer;

	OPCITEMDEF		m_ReadItems[READ_ITEMS];
	OPCITEMRESULT	*m_pReadItemResult;
	OPCHANDLE		m_RcvGrpSrvHandle;
	IOPCItemMgt		*m_pIOPCItemMgt_Rcv;
	IOPCSyncIO		*m_pIOPCSyncIO_Rcv;
	OPCHANDLE		*m_phServer;

	OPCITEMDEF		m_WriteItems[WRITE_ITEMS];
	OPCITEMRESULT	*m_pWriteItemResult;
	OPCHANDLE		m_WriteGrpSrvHandle;
	IOPCItemMgt		*m_pIOPCItemMgt_Write;
	IOPCSyncIO		*m_pIOPCSyncIO_Write;

	HRESULT			*m_pRcvErrors;
	HRESULT			*m_pWriteErrors;

	
	/* Debug File */
	ofstream fDebugFile;
	ofstream fDebugFileThread;
	ofstream fDebugFileReadWrite;
	CItfDB m_cIftDB;
	/* Recipe Pointer */
	typRecipe *m_pRecipe;
};
