// OPCClient.cpp: Implementation of COPCClient Class
//

#include "StdAfx.h"
#include "OPCClient.h"


#define DEBUG_FILE

//////////////////////////////////////////////////////////////////////
// Constructors/Destructor
//////////////////////////////////////////////////////////////////////
COPCClient::COPCClient(void)
{
	m_csSecure.Init();	
}

COPCClient::~COPCClient(void)
{
	m_pRecipe = NULL;
}

long COPCClient::Connect(long lngInterval)
{
	// Read/Write Interval (ms)
	m_lngInterval		=  lngInterval;

	//Create OPC Connection
	HRESULT		ret;
	CLSID		clsid;
	LONG		TimeBias = 0;
	FLOAT		PercentDeadband = 0.0;
	DWORD		RevisedUpdateRate;
	LPWSTR		ErrorStr;

#ifdef DEBUG_FILE
	fDebugFile.open("test.txt", std::ofstream::out | std::ofstream::app);
	fDebugFile << "Connection to PLC..." << endl;
#endif

	//	Given a ProgID, this looks up the associated CLSID in the registry
	ret = CLSIDFromProgID(L"OPC.SimaticNET", &clsid);
	if (ret != S_OK)
	{
		m_strStatusMessage = "Error Get ClassID";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		return E_FAIL;
	}

	//	Create the OPC server object and query for the IOPCServer interface of the object
	ret = CoCreateInstance (clsid, NULL, CLSCTX_LOCAL_SERVER ,IID_IOPCServer, (void**)&m_pIOPCServer);
	if (ret != S_OK)
	{						
		m_strStatusMessage = "Connect", "Error Create OPC Server";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		m_pIOPCServer = NULL;
		return E_FAIL;
	}

	//	Add a read group object "ums" and query for interface IOPCItemMgt
	ret=m_pIOPCServer->AddGroup(L"rcvData",		//	[in] group name
						TRUE,					//	[in] active
						lngInterval,			//	[in] request this Update Rate from Server
						1,						//	[in] Client handle
						&TimeBias,				//	[in] no time interval to system UTC time
						&PercentDeadband,		//	[in] no deadband, so all data changes are reported	
						LOCALE_ID,				//	[in] Server uses English language for text values
						&m_RcvGrpSrvHandle,		//	[out] Server handle to identify this group in later calls
						&RevisedUpdateRate,		//	[out] the answer form the Server to the requested update rate
						IID_IOPCItemMgt,		//	[in] requested interface type of the group object
						(LPUNKNOWN*)&m_pIOPCItemMgt_Rcv);	//	[out] pointer to the requested interface

	if (ret == OPC_S_UNSUPPORTEDRATE)
	{	
		sprintf_s(m_strStatusMessage, 100, "Unsupported Update Rate (%d): Revised to %d ms", lngInterval,RevisedUpdateRate );
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
	}
	else
	{
		if (FAILED(ret))
		{
			m_strStatusMessage = "Can't add Read Group to Server!";
#ifdef DEBUG_FILE
			fDebugFile << m_strStatusMessage << endl;
			fDebugFile.close();
#endif
			m_pIOPCServer->Release();
			m_pIOPCServer = NULL;
			return E_FAIL;
		}
	}

	//	define an item table with one item as in-paramter for AddItem 
	LPWSTR			 szReadItemID[READ_ITEMS];

	szReadItemID[0]  = L"S7:[UMS_SPS]DB852,X0.0";	
	szReadItemID[1]  = L"S7:[UMS_SPS]DB852,STRING2.50";
	szReadItemID[2]  = L"S7:[UMS_SPS]DB852,STRING54.50";
	szReadItemID[3]  = L"S7:[UMS_SPS]DB852,STRING106.50";
	szReadItemID[4]  = L"S7:[UMS_SPS]DB852,STRING158.10";
	szReadItemID[5]  = L"S7:[UMS_SPS]DB852,STRING170.10";
	szReadItemID[6]  = L"S7:[UMS_SPS]DB852,INT182,1";
	szReadItemID[7]  = L"S7:[UMS_SPS]DB852,INT184,1";
	szReadItemID[8]  = L"S7:[UMS_SPS]DB852,DINT186,1";
	szReadItemID[9]  = L"S7:[UMS_SPS]DB852,INT190,1";
	szReadItemID[10] = L"S7:[UMS_SPS]DB852,STRING192.20";

	for(int i=0;i<READ_ITEMS;i++)
	{
		m_ReadItems[i].szAccessPath		  = L"";					//	Accesspath not needed
		m_ReadItems[i].szItemID			  = szReadItemID[i];		//	ItemID, see above
		m_ReadItems[i].bActive			  = TRUE;			
		m_ReadItems[i].hClient			  = 1;
		m_ReadItems[i].dwBlobSize		  = 0;
		m_ReadItems[i].pBlob			  = NULL;
		m_ReadItems[i].vtRequestedDataType = 0;				//	return values in native (cannonical) datatype 
	}													//	defined by the item itself

	ret = m_pIOPCItemMgt_Rcv->AddItems(READ_ITEMS,					//	[in] add no of items
							m_ReadItems,						//	[in]  defin
							&m_pReadItemResult,					//	[out] array with additional information about the item
							&m_pRcvErrors);						//	[out] tells which of the items was successfully added.
																//	For any item which failed it provides a reason

	if ( (ret != S_OK) && (ret != S_FALSE) )
	{		
		sprintf_s(m_strStatusMessage, 100, "AddItems (Read) failed! %d", ret);
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		m_pIOPCItemMgt_Rcv->Release();
		m_pIOPCItemMgt_Rcv = NULL;
		m_RcvGrpSrvHandle = NULL;
		return E_FAIL;
	}
	
	if(ret == S_FALSE)
	{
		for(int i=0;i<READ_ITEMS;i++)
		{
			m_pIOPCServer->GetErrorString(m_pRcvErrors[i], LOCALE_ID, &ErrorStr);
			sprintf_s(m_strStatusMessage, 100, "Additem [%d]: %ls", i, ErrorStr);
#ifdef DEBUG_FILE
			fDebugFile << m_strStatusMessage << endl;
#endif
		}

	}
	//	query interface for sync calls on group object
	ret = m_pIOPCItemMgt_Rcv->QueryInterface(IID_IOPCSyncIO, (void**)&m_pIOPCSyncIO_Rcv);
	if (ret < 0)
	{	
		m_strStatusMessage = "Connect Rcv: No IOPCSyncIO found!";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		CoTaskMemFree(m_pReadItemResult);
		m_pIOPCItemMgt_Rcv->Release();
		m_pIOPCItemMgt_Rcv= NULL;
		m_RcvGrpSrvHandle = NULL;
		return E_FAIL;
	}			

	
	if (ret == OPC_S_UNSUPPORTEDRATE)
	{	
		sprintf_s(m_strStatusMessage, 100, "Unsupported Update Rate (%d): Revised to %d ms", lngInterval,RevisedUpdateRate );
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
	}
	else
	{
		if (FAILED(ret))
		{
			m_strStatusMessage = "Can't add Group to Server!";
#ifdef DEBUG_FILE
			fDebugFile << m_strStatusMessage << endl;
			fDebugFile.close();
#endif
			m_pIOPCServer->Release();
			m_pIOPCServer = NULL;
			return E_FAIL;
		}
	}

	//WRITE GROUP

	//	Add a write group object "ums" and query for interface IOPCItemMgt
	ret=m_pIOPCServer->AddGroup(L"writeData",		//	[in] group name
						FALSE,					//	[in] active
						lngInterval,			//	[in] request this Update Rate from Server
						2,						//	[in] Client handle
						&TimeBias,				//	[in] no time interval to system UTC time
						&PercentDeadband,		//	[in] no deadband, so all data changes are reported	
						LOCALE_ID,				//	[in] Server uses English language for text values
						&m_WriteGrpSrvHandle,		//	[out] Server handle to identify this group in later calls
						&RevisedUpdateRate,		//	[out] the answer form the Server to the requested update rate
						IID_IOPCItemMgt,		//	[in] requested interface type of the group object
						(LPUNKNOWN*)&m_pIOPCItemMgt_Write);	//	[out] pointer to the requested interface

	if (ret == OPC_S_UNSUPPORTEDRATE)
	{	
		sprintf_s(m_strStatusMessage, 100, "Unsupported Update Rate (%d): Revised to %d ms", lngInterval,RevisedUpdateRate );
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
	}
	else
	{
		if (FAILED(ret))
		{
			m_strStatusMessage = "Can't add Write Group to Server!";
#ifdef DEBUG_FILE
			fDebugFile << m_strStatusMessage << endl;
			fDebugFile.close();
#endif
			m_pIOPCServer->Release();
			m_pIOPCServer = NULL;
			return E_FAIL;
		}

	}

	//	define an item table with one item as in-paramter for AddItem 

	LPWSTR			szWriteItemID[WRITE_ITEMS];

	
	szWriteItemID[0]  = L"S7:[UMS_SPS]DB852,B0,1";	//bits	

	for(int i=0;i<WRITE_ITEMS;i++)
	{
		m_WriteItems[i].szAccessPath			= L"";					//	Accesspath not needed
		m_WriteItems[i].szItemID				= szWriteItemID[i];		//	ItemID, see above
		m_WriteItems[i].bActive					= TRUE;			
		m_WriteItems[i].hClient					= 2;
		m_WriteItems[i].dwBlobSize				= 0;
		m_WriteItems[i].pBlob					= NULL;
		m_WriteItems[i].vtRequestedDataType		= 0;				//	return values in native (cannonical) datatype 
	}													//	defined by the item itself

	ret = m_pIOPCItemMgt_Write->AddItems(WRITE_ITEMS,				//	[in] add one item
							m_WriteItems,							//	[in] see above
							&m_pWriteItemResult,						//	[out] array with additional information about the item
							&m_pWriteErrors);						//	[out] tells which of the items was successfully added.
																	//	For any item which failed it provides a reason

	if ( (ret != S_OK) && (ret != S_FALSE) )
	{	
		m_pIOPCServer->GetErrorString(m_pWriteErrors[0], LOCALE_ID, &ErrorStr);
		sprintf_s(m_strStatusMessage, 100, "AddItems (Read) failed! %d", ret);
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		m_pIOPCItemMgt_Write->Release();
		m_pIOPCItemMgt_Write = NULL;
		m_WriteGrpSrvHandle = NULL;
		return E_FAIL;
	}
	
	if(ret == S_FALSE)
	{
		for(int i=0;i<WRITE_ITEMS;i++)
		{
			m_pIOPCServer->GetErrorString(m_pWriteErrors[i], LOCALE_ID, &ErrorStr);
			sprintf_s(m_strStatusMessage, 100, "Additem [%d]: %ls", i, ErrorStr);
#ifdef DEBUG_FILE
			fDebugFile << m_strStatusMessage << endl;
#endif
		}

	}

	//	query interface for sync calls on group object
	ret = m_pIOPCItemMgt_Write->QueryInterface(IID_IOPCSyncIO, (void**)&m_pIOPCSyncIO_Write);
	if (ret < 0)
	{	
		m_strStatusMessage = "Connect Write: No IOPCSyncIO found!";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
		fDebugFile.close();
#endif
		CoTaskMemFree(m_pWriteItemResult);
		m_pIOPCItemMgt_Write->Release();
		m_pIOPCItemMgt_Write = NULL;
		m_WriteGrpSrvHandle = NULL;
		
		return E_FAIL;
	}


	HRESULT			*pErrors;
	//Memory allocation really needed, if more than one item is to be written
	//Select item by server handle, received at AddItem
	m_phServer = new OPCHANDLE[WRITE_ITEMS];

	for(int i=0;i < WRITE_ITEMS ;i++)
		m_phServer[i] = m_pWriteItemResult[i].hServer;

	ret = m_pIOPCItemMgt_Write->SetActiveState(WRITE_ITEMS,m_phServer,TRUE,&pErrors);
	if(ret != S_OK)
		return(ret);
	
	m_strStatusMessage = "PLC Connected!";

#ifdef DEBUG_FILE
	fDebugFile << m_strStatusMessage << endl;
#endif

	// Reset Exit Thread Condition
	m_bLoopExit	= false;

	// New thread for permanent reading and writing from and to PLC
	///m_hPLCLoop = CreateThread(NULL,0,(LPTHREAD_START_ROUTINE)PLCThread,this,0,&m_dwPLCLoop);	//start a seperate thread for work!

	if(m_hPLCLoop == NULL)
		m_strStatusMessage = "Error during Thread creation.";
	else
		m_strStatusMessage = "PLC Thread created.";

	m_bIsRunning = true;

#ifdef DEBUG_FILE
	fDebugFile << m_strStatusMessage << endl;
	fDebugFile.close();
#endif
	loop(); ///TODEL
	return S_OK;
}
void COPCClient::loop()
{
	long t1,t2,t3;
	while(1)
	{
		// Stop Thread
		if( m_bLoopExit )
		{
#ifdef DEBUG_FILE
			fDebugFileThread << "Request stop thread" << endl;
#endif			//TODELUSDiag(1,'run',"Loop", "Leave RunLoop");
			return;
		}
		//TODELUSDiag( 6, 'sps', "SpsLoop", "count:%d",count++  );
		t1 = timeGetTime();
		
		//PLC RcvData
		ReadFromPLC();
		t2 = timeGetTime();

		//Quit bit New Coil
		if(!m_pRecipe->bNewCoil && m_pRecipe->bAckNewCoil)
			m_pRecipe->bNewCoil = false;

		//SPS WriteData
		WriteToPLC();	
		t3 = timeGetTime();

		// Sleep
		if(t3-t1<m_lngInterval)
			Sleep(m_lngInterval-(t3-t1));
	}
}
void COPCClient::SetRecipe(typRecipe *pRecipe)
{
	m_pRecipe = pRecipe;
}

void COPCClient::PLCThreadWorker()
{
	long t1,t2,t3;
	long count = 0;

	CoInitializeEx( NULL, COINIT_MULTITHREADED );
#ifdef DEBUG_FILE
	fDebugFileThread.open("PLC_thread.txt", std::ofstream::out | std::ofstream::app);
	fDebugFileThread << "Start PLC Thread..." << endl;
#endif


	//Init of Critical Section
	m_csSecure.Init();	

	m_csSecure.Lock();

	while( 1 )
	{

		m_csSecure.Unlock();
			Sleep( 0 );	// Complete process to set m_bexit 
		m_csSecure.Lock();
		// Stop Thread
		if( m_bLoopExit )
		{
#ifdef DEBUG_FILE
			fDebugFileThread << "Request stop thread" << endl;
#endif			//TODELUSDiag(1,'run',"Loop", "Leave RunLoop");
			return;
		}
		//TODELUSDiag( 6, 'sps', "SpsLoop", "count:%d",count++  );
		t1 = timeGetTime();
		
		//PLC RcvData
		ReadFromPLC();
		t2 = timeGetTime();

		//Quit bit New Coil
		if(!m_pRecipe->bNewCoil && m_pRecipe->bAckNewCoil)
			m_pRecipe->bNewCoil = false;

		//SPS WriteData
		WriteToPLC();	
		t3 = timeGetTime();

		// Sleep
		if(t3-t1<m_lngInterval)
			Sleep(m_lngInterval-(t3-t1));

	}
#ifdef DEBUG_FILE
	fDebugFileThread << "Stop PLC Thread" << endl;
	fDebugFileThread.close();
#endif

	CoUninitialize();
}

/////////////////////////////////////////////////////////////////////////////
long COPCClient::ReadFromPLC(void)
{
	HRESULT			*pErrors;
	OPCHANDLE		*phServer;
	OPCITEMSTATE	*pItemValue;
	HRESULT			ret = 0;
	LPWSTR			ErrorStr;
	long			tstart = 0; //TODELtimeGetTime();

#ifdef DEBUG_FILE
	fDebugFileReadWrite.open("ReadWrite.txt", std::ofstream::out | std::ofstream::app);
	fDebugFileReadWrite << "Start Read From PLC" << endl;
#endif

	//m_pRecipe->bNewCoil = false; ///////////////////////////////////////////////TODEL/////////////////////////////////////////
	//Memory allocation really needed, if more than one item is to be read
	phServer = new OPCHANDLE[READ_ITEMS];

	//Select item by server handle, received at AddItem
	for(int i=0;i < READ_ITEMS;i++)
	{		
		if (m_pRcvErrors[i] != S_OK)		//Item not available
		{
			sprintf_s(m_strStatusMessage, 100, "Error : OPC Item [%d] not available", i);
#ifdef DEBUG_FILE
			fDebugFileReadWrite << m_strStatusMessage << endl;
#endif
		}

		phServer[i] = m_pReadItemResult[i].hServer;
	}
	// Read Data from OPC Server
	ret = m_pIOPCSyncIO_Rcv->Read(OPC_DS_DEVICE, READ_ITEMS, phServer, &pItemValue, &pErrors);
	delete[] phServer;

	// If reading is OK
	if (ret == S_OK) 
	{
		// If one or several item has not a good OPC quality
		for(int i=0;i < READ_ITEMS;i++)
		{
			if(pItemValue[i].wQuality != OPC_QUALITY_GOOD)
			{
				sprintf_s(m_strStatusMessage, 100, "Item-Quality(%d): %d not GOOD", i , pItemValue[i].wQuality);
#ifdef DEBUG_FILE
				fDebugFileReadWrite << m_strStatusMessage << endl;
#endif
			}
		}
		// Affect data values to members
		if(pItemValue[0].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->bNewCoil			=	 (bool)pItemValue[0].vDataValue.bVal;

		if(pItemValue[1].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strOrderName		=	 _com_util::ConvertBSTRToString(pItemValue[1].vDataValue.bstrVal);
		
		if(pItemValue[2].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strCustomerName	=	 _com_util::ConvertBSTRToString(pItemValue[2].vDataValue.bstrVal);

		if(pItemValue[3].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strCoilNumber		=	 _com_util::ConvertBSTRToString(pItemValue[3].vDataValue.bstrVal);
		
		if(pItemValue[4].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strNuance			=	 _com_util::ConvertBSTRToString(pItemValue[4].vDataValue.bstrVal);

		if(pItemValue[5].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strEtat			=	 _com_util::ConvertBSTRToString(pItemValue[5].vDataValue.bstrVal);

		if(pItemValue[6].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->lngThickness		=	(long)pItemValue[6].vDataValue.lVal;

		if(pItemValue[7].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->lngWidth			=	(long)pItemValue[7].vDataValue.lVal;

		if(pItemValue[8].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->lngEModul			=	(long)pItemValue[8].vDataValue.lVal;
		
		if(pItemValue[9].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->lngYieldStrength	=	(long)pItemValue[9].vDataValue.lVal;

		if(pItemValue[10].wQuality == OPC_QUALITY_GOOD)
			m_pRecipe->strCassetteName	=	 _com_util::ConvertBSTRToString(pItemValue[10].vDataValue.bstrVal);		

#ifdef DEBUG_FILE
		fDebugFileReadWrite << "Read from PLC OK" << endl;
#endif
/////////////////////TODEL////////////////////////////////////////////////////////////////////////////////
		
		if(m_pRecipe->bNewCoil && !m_pRecipe->bAckNewCoil)
		{
			m_pRecipe->bAckNewCoil = true;
			m_cIftDB.UpdateDB(m_pRecipe);
#ifdef DEBUG_FILE
			fDebugFileReadWrite << "NEW COIL!!" << endl;
			fDebugFileReadWrite << m_pRecipe->strOrderName << endl;
			fDebugFileReadWrite << m_pRecipe->strCoilNumber << endl;
			fDebugFileReadWrite << m_pRecipe->strCustomerName << endl;
			fDebugFileReadWrite << m_pRecipe->strNuance << endl;
			fDebugFileReadWrite << m_pRecipe->strEtat << endl;
			fDebugFileReadWrite << m_pRecipe->lngThickness << endl;
			fDebugFileReadWrite << m_pRecipe->lngWidth << endl;
			fDebugFileReadWrite << m_pRecipe->lngYieldStrength << endl;
			fDebugFileReadWrite << m_pRecipe->strCassetteName << endl;
#endif
		}
		else
			m_pRecipe->bAckNewCoil = false;
/////////////////////////////////////////////////////////////////////////////////////////////////////
	}

	m_bPLCConnectionFailed = false;

	if (ret == S_FALSE) 
	{
		for(int i=0;i<READ_ITEMS;i++)
		{
			m_pIOPCServer->GetErrorString(pErrors[i], LOCALE_ID, &ErrorStr);
			//sprintf(m_strStatusMessage, "Error (%d): %ls(%d)",i , ErrorStr ,pErrors[i]);
			if(pErrors[i] == -1073446909)
				m_bPLCConnectionFailed = true;
		}
	}

	if (FAILED(ret)) 
	{
		m_strStatusMessage = "Read failed!";
#ifdef DEBUG_FILE
		fDebugFileReadWrite << m_strStatusMessage << endl;
#endif
	}
	else 
	{		
		for(int i=0;i<READ_ITEMS;i++)
			VariantClear(&pItemValue[i].vDataValue);
		//	release out-parameters in case of not failed
		CoTaskMemFree(pErrors);
		CoTaskMemFree(pItemValue);
	}
#ifdef DEBUG_FILE
	fDebugFileReadWrite << "Read from PLC completed" << endl;
	fDebugFileReadWrite.close();
#endif
	return S_OK;
}

long COPCClient::WriteToPLC(void)
{
	HRESULT			*pErrors;
	HRESULT			ret;
	LPWSTR			ErrorStr;
	BYTE			B0=0;
	short			INT0=0;
	VARIANT			vValues[WRITE_ITEMS];
	
#ifdef DEBUG_FILE
	fDebugFileReadWrite.open("ReadWrite.txt", std::ofstream::out | std::ofstream::app);
	fDebugFileReadWrite << "Start Write to PLC" << endl;
#endif

	//Select item by server handle, received at AddItem
	for(int i=0;i < WRITE_ITEMS ;i++)
	{

		VariantClear(&vValues[i]);
        VariantInit(&vValues[i]);
		if (m_pWriteErrors[i] != S_OK)		//Item not available
			sprintf_s(m_strStatusMessage, 100, "Error : OPC Item [%d] not available", i);

		m_phServer[i] = m_pWriteItemResult[i].hServer;
	}


	//Build Send Byte Addr.0:
	//bit 1:	AckNewCoil

	if(m_pRecipe->bAckNewCoil)
		B0   = (BYTE)0x02;
	else
		B0   = (BYTE)0x00;

	vValues[0].vt		= VT_UI1;
	vValues[0].uintVal 	= (BYTE)B0; 

	ret = m_pIOPCSyncIO_Write->Write(WRITE_ITEMS, m_phServer, vValues, &pErrors);

	m_bPLCConnectionFailed = false;

	if (ret == S_FALSE) 
	{
		for(int i=0;i<WRITE_ITEMS;i++)
		{
			m_pIOPCServer->GetErrorString(pErrors[i], LOCALE_ID, &ErrorStr);
			sprintf_s(m_strStatusMessage, 100, "Error (%d): %ls(%d)",i,ErrorStr ,pErrors[i]);
		}
	}

	if (FAILED(ret)) 
	{
		//sprintf_s(m_strStatusMessage, 100, "Error at write: %d", ret );
	}
	else 
	{
		// Release out-parameters in case of not failed	
		CoTaskMemFree(pErrors);
	}

#ifdef DEBUG_FILE
	fDebugFileReadWrite << "Write to PLC completed" << endl;
	fDebugFileReadWrite.close();
#endif
	return S_OK;
}

long COPCClient::Close(void)
{
	HRESULT		ret;
	OPCHANDLE	*phServer;
	HRESULT		*pErrors;
	LPWSTR		ErrorStr;

	m_bIsRunning = false;

#ifdef DEBUG_FILE
	fDebugFile.open("test.txt", std::ofstream::out | std::ofstream::app);
	fDebugFile << "Closing connection..." << endl;
#endif
	// Stop PLC Loop
	m_csSecure.Lock();
		m_bLoopExit = true;
	m_csSecure.Unlock();


	//	Remove Read Items
	phServer = new OPCHANDLE[READ_ITEMS];

	for(int i=0; i<READ_ITEMS;i++)
		phServer[i] = m_pReadItemResult[i].hServer;

	ret = m_pIOPCItemMgt_Rcv->RemoveItems(READ_ITEMS,	//	[in] remove one item
							phServer,					//	[in] identified by server handle
							&pErrors);					//	[out] Errors for each item, returned by OPC server
	
	if ( (ret != S_OK) && (ret != S_FALSE) )
	{
		m_strStatusMessage = "PLC Disconnection ... RemoveItems failed!";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
#endif
	}

	if (ret == S_FALSE)
		m_pIOPCServer->GetErrorString(pErrors[0], LOCALE_ID, &ErrorStr);


	delete[] phServer;

	//	Remove Write Items
	phServer = new OPCHANDLE[WRITE_ITEMS];

	for(int i=0; i<WRITE_ITEMS;i++)
		phServer[i] = m_pWriteItemResult[i].hServer;

	ret = m_pIOPCItemMgt_Rcv->RemoveItems(WRITE_ITEMS,	//	[in] remove one item
							phServer,					//	[in] identified by server handle
							&pErrors);					//	[out] Errors for each item, returned by OPC server
	
	if ( (ret != S_OK) && (ret != S_FALSE) )
	{
		m_strStatusMessage = "PLC Disconnection ... RemoveItems failed!";
#ifdef DEBUG_FILE
		fDebugFile << m_strStatusMessage << endl;
#endif
	}

	if (ret == S_FALSE)
	{		
		m_pIOPCServer->GetErrorString(pErrors[0], LOCALE_ID, &ErrorStr);
		//TODELUSDiag( 0, 'opc', "Result disconnect", "%ls",ErrorStr );
		//TODELm_strStatusMessage = ErrorStr;
	}

	delete[] phServer;

	CoTaskMemFree(pErrors);

	CoTaskMemFree(m_pReadItemResult);
	m_pReadItemResult=NULL;
	CoTaskMemFree(m_pWriteItemResult);
	m_pWriteItemResult=NULL;

	CoTaskMemFree(m_pWriteErrors);
	m_pWriteErrors = NULL;
	CoTaskMemFree(m_pRcvErrors);
	m_pRcvErrors = NULL;

	//	Release interface for sync calls
	m_pIOPCSyncIO_Write->Release();
	m_pIOPCSyncIO_Write = NULL;
	m_pIOPCSyncIO_Rcv->Release();
	m_pIOPCSyncIO_Rcv = NULL;

	//	Release ItemManagement interface
	m_pIOPCItemMgt_Write->Release();
	m_pIOPCItemMgt_Write = NULL;
	m_pIOPCItemMgt_Rcv->Release();
	m_pIOPCItemMgt_Rcv = NULL;
	

	//	Remove Group
	ret = m_pIOPCServer->RemoveGroup(m_RcvGrpSrvHandle, TRUE);
	m_RcvGrpSrvHandle = NULL;

	ret = m_pIOPCServer->RemoveGroup(m_WriteGrpSrvHandle, TRUE);
	m_WriteGrpSrvHandle = NULL;

	//	Release OPC-Server
	m_pIOPCServer->Release();
	m_pIOPCServer = NULL;

	Sleep(100);
	CloseHandle(m_hPLCLoop);
	m_csSecure.Term();
	m_strStatusMessage = "PLC Disconnected!";
#ifdef DEBUG_FILE
	fDebugFile << m_strStatusMessage << endl;
	fDebugFile.close();
#endif
	return S_OK;
}

char * COPCClient::getPLCStatusMessage(void)
{
	return m_strStatusMessage;
}
