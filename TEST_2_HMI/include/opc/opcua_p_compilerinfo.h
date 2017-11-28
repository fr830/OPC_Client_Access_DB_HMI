/******************************************************************************
**
** This software is licensed under the proprietary Software License Agreement
** (SLA) of Unified Automation GmbH
** You are not allowed to publish, distribute, sublicense, and/or sell the
** Source Code of this Software
** Get full License Agreement at: license@www.unified-automation.com
**
** Copyright (C) 2005-2010 Unified Automation GmbH. All Rights Reserved.
** Web: http://www.unified-automation.com
** This file is provided AS IS with NO WARRANTY OF ANY KIND, INCLUDING THE
** WARRANTY OF DESIGN, MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.
**
** This code is derived work and it was substantially changed!
** The original MIT License can be found here:
** http://opcfoundation.org/License/MIT/1.00/
**
** Project: OPC UA Ansi C Communication Stack Windows Platform Layer
**
** Description: Implementation of the OPC UA Ansi C Communication Stacks
**              Platform Layer Interface
**
******************************************************************************/

#ifndef _OpcUa_P_CompilerInfo_H_
#define _OpcUa_P_CompilerInfo_H_ 1

/* compiler information constants */
#define OPCUA_P_COMPILERNAME_UNKNOWN    "Unknown Compiler"
#define OPCUA_P_COMPILERNAME_MSVC       "Microsoft Visual C/C++"
#define OPCUA_P_COMPILERNAME_MINGNUW    "GNU C++/MINGW"
#define OPCUA_P_COMPILERNAME_GNU        "GNU C++"
#define OPCUA_P_COMPILERNAME_INTEL      "Intel C++"

/* check for known compilers */
#if defined(_MSC_VER)

  /* compiler name */
# if defined(__INTEL_COMPILER)
#  define OPCUA_P_COMPILERNAME OPCUA_P_COMPILERNAME_INTEL
# else
#  define OPCUA_P_COMPILERNAME OPCUA_P_COMPILERNAME_MSVC
# endif
  /* compiler version */
# define OPCUA_P_COMPILERVERSION OPCUA_TOSTRING(_MSC_VER)

#elif defined(__GNUC__)

  /* compiler name */
# if defined(__MINGW32__)
#  define OPCUA_P_COMPILERNAME OPCUA_P_COMPILERNAME_MINGNUW
# elif defined(__INTEL_COMPILER)
#  define OPCUA_P_COMPILERNAME OPCUA_P_COMPILERNAME_INTEL
# else
#  define OPCUA_P_COMPILERNAME OPCUA_P_COMPILERNAME_GNU
# endif
  /* compiler version */
# define OPCUA_P_COMPILERVERSION OPCUA_TOSTRING(__GNUC__)"."OPCUA_TOSTRING(__GNUC_MINOR__)

#else /* compiler */

/* compiler unknown */
# define OPCUA_P_COMPILERNAME       OPCUA_P_COMPILERNAME_UNKNOWN
# define OPCUA_P_COMPILERVERSION    "0"

#endif /* compiler */

/* create defines used by the stack */
#define OPCUA_P_COMPILERINFO OPCUA_P_COMPILERNAME " " OPCUA_P_COMPILERVERSION

#endif /* _OpcUa_P_CompilerInfo_H_ */
