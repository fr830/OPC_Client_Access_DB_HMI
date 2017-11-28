/****************************************************************************/
/*  Copyright (C) Siemens AG 1996-97  All Rights Reserved. Confidential     */
/****************************************************************************
 * ------------------------------------------------------------------------  
 *  Project  : SimaticNET OPC Server    OOO    PPPP     CCC               
 *  Filename : opcfrm.h                O   O   P   P   C                   
 *  Date     : 21.04.97 V 1.0          O   O   PPPP    C                   
 *  Update   : 00.00.00                O   O   P       C                 
 *  Author   : Deiretsbacher            OOO    P        CCC               
 *  Compiler : Visual C++ 4.2
 *  OS       : Windows NT 4.0
 *                                                            
 * ------------------------------------------------------------------------  
 *                                                                           
 *  Abstract:  
 *  	Defines for FormatEtc
 *  
 *  Modification Log:
 *   Version     Date    By    Notes
 *   --------   -------- ---   -----
 *   01.00.00   21.04.97 dei
 *                    
 * ------------------------------------------------------------------------  
 *                                                                           
 ****************************************************************************/

#ifndef __OPCFRM_H
#define __OPCFRM_H

#ifdef UNICODE
#define OPC_DATASTREAMFORMAT1  L"OPCSTMFORMATDATA"
#define OPC_DATASTREAMFORMAT2  L"OPCSTMFORMATDATATIME"
#define OPC_DATASTREAMFORMAT3  L"OPCSTMFORMATWRITECOMPLETE"
#else
#define OPC_DATASTREAMFORMAT1  "OPCSTMFORMATDATA"
#define OPC_DATASTREAMFORMAT2  "OPCSTMFORMATDATATIME"
#define OPC_DATASTREAMFORMAT3  "OPCSTMFORMATWRITECOMPLETE"
#endif

#endif // OpcFrm

/****************************************************************************/
/*  Copyright (C) Siemens AG 1996-97  All Rights Reserved. Confidential     */
/****************************************************************************/
