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

/* This are the win32 platform definitions! */

#ifndef _OpcUa_PlatformDefs_H_
#define _OpcUa_PlatformDefs_H_ 1

#ifdef _MSC_VER
  #pragma warning( push )
  #pragma warning( disable : 4001 )
#endif /* _MSC_VER */

/* System Headers */
#include <stdio.h>
#include <stdarg.h>

#ifdef _MSC_VER
  #pragma warning( pop )
  #pragma warning (disable: 4115)
  #pragma warning (disable: 4102)
#endif /* _MSC_VER */

#ifdef __cplusplus
# define OPCUA_BEGIN_EXTERN_C extern "C" {
# define OPCUA_END_EXTERN_C }
#else
# define OPCUA_BEGIN_EXTERN_C
# define OPCUA_END_EXTERN_C
#endif

/*============================================================================
* Types and Mapping
*===========================================================================*/

OPCUA_BEGIN_EXTERN_C

#define OPCUA_P_LITTLE_ENDIAN 1234
#define OPCUA_P_BIG_ENDIAN    4321
#define OPCUA_P_BYTE_ORDER OPCUA_P_LITTLE_ENDIAN

#define OPCUA_P_EXPORT __declspec(dllexport)
#define OPCUA_P_IMPORT __declspec(dllimport)

/* marks functions that need to be exported via functiontable */
#if defined(_UA_STACK_BUILD_DLL)
  #define OPCUA_EXPORT OPCUA_P_EXPORT
#else
  #define OPCUA_EXPORT
#endif

#define OPCUA_EXPORT_SYNC_SERVER_API OPCUA_EXPORT

#if defined(_UA_STACK_USE_DLL)
  #define OPCUA_IMPORT OPCUA_P_IMPORT
#else
  #define OPCUA_IMPORT
#endif

#if defined(_UA_STACK_BUILD_DLL)
    #define OPCUA_IMEXPORT OPCUA_P_EXPORT
#elif defined(_UA_STACK_USE_DLL)
    #define OPCUA_IMEXPORT OPCUA_P_IMPORT
#else
  /* build/using static lib */
  #define OPCUA_IMEXPORT
#endif

/* call exported functions by stdcall convention */
#ifdef _WIN32_WCE
  #define OPCUA_DLLCALL
#else
  #define OPCUA_DLLCALL __stdcall
#endif

/* calling convention used by stack functions that explicitely use cdecl */
#define OPCUA_CDECL __cdecl

/* used ie. for unlimited timespans */
#define OPCUA_INFINITE 0xFFFFFFFF

/* helper macro to convert a numeric define into a string (only use OPCUA_TOSTRING) */
#ifndef OPCUA_TOSTRING_HELP
# define OPCUA_TOSTRING_HELP(xNumber)  #xNumber
#endif

#ifndef OPCUA_TOSTRING
# define OPCUA_TOSTRING(xNumber) OPCUA_TOSTRING_HELP(xNumber)
#endif

/*============================================================================
* Additional basic headers
*===========================================================================*/
#include <string.h>

/* get compiler name and version */
#include <opcua_p_compilerinfo.h>

/*============================================================================
 * platform layer version information
 *===========================================================================*/
/* helper to get build time through buildsystem */
#ifndef OPCUA_P_BUILDINFO_BUILD_TIMESTAMP
#if (defined(__DATE__) && defined(__TIME__))
#define OPCUA_P_TIMESTAMP __DATE__ " " __TIME__
#define OPCUA_P_BUILDINFO_BUILD_TIMESTAMP OPCUA_P_TIMESTAMP
#endif /* __DATE__ && __TIME__ */
#endif /* OPCUA_P_BUILDINFO_BUILD_TIMESTAMP */

/** @brief Name of the platform or operating system this platform layer was created for. */
#ifndef OPCUA_P_BUILDINFO_PLATFORMNAME
# define OPCUA_P_BUILDINFO_PLATFORMNAME             "Windows V6"
#endif /* OPCUA_P_BUILDINFO_PLATFORMNAME */

/** @brief The date and time when the source was last modified. */
#ifndef OPCUA_P_BUILDINFO_SOURCE_TIMESTAMP
# define OPCUA_P_BUILDINFO_SOURCE_TIMESTAMP         "OPCUA_P_BUILDINFO_SOURCE_TIMESTAMP not set"
#endif /* OPCUA_P_BUILDINFO_SOURCE_TIMESTAMP */

/** @brief The date and time when the binary is build. */
#ifndef OPCUA_P_BUILDINFO_BUILD_TIMESTAMP
# define OPCUA_P_BUILDINFO_BUILD_TIMESTAMP          "OPCUA_P_BUILDINFO_BUILD_TIMESTAMP not set"
#endif /* OPCUA_P_BUILDINFO_BUILD_TIMESTAMP */

/** @brief The name of the company which build the binary. */
#ifndef OPCUA_P_BUILDINFO_VENDOR_NAME
# define OPCUA_P_BUILDINFO_VENDOR_NAME              "OPCUA_P_BUILDINFO_VENDOR_NAME not set"
#endif /* OPCUA_P_BUILDINFO_VENDOR_NAME */

/** @brief Additional information from the company, ie. internal revision number. */
#ifndef OPCUA_P_BUILDINFO_VENDOR_INFO
# define OPCUA_P_BUILDINFO_VENDOR_INFO              "OPCUA_P_BUILDINFO_VENDOR_INFO not set"
#endif /* OPCUA_P_BUILDINFO_VENDOR_INFO */

/** @brief Additional information from the company, ie. internal revision number. */
#ifndef OPCUA_P_BUILDINFO_COMPILER
# define OPCUA_P_BUILDINFO_COMPILER                 OPCUA_P_COMPILERINFO
#endif /* OPCUA_P_BUILDINFO_COMPILER */

/** @brief Separator used between different fields in the version string. */
#define OPCUA_P_VERSIONSTRING_SEPARATOR             "\\"

/** @brief The versionstring returned by OpcUa_P_GetVersion. */
#define OPCUA_P_VERSIONSTRING                       OPCUA_P_BUILDINFO_PLATFORMNAME      OPCUA_P_VERSIONSTRING_SEPARATOR \
                                                    OPCUA_P_BUILDINFO_SOURCE_TIMESTAMP  OPCUA_P_VERSIONSTRING_SEPARATOR \
                                                    OPCUA_P_BUILDINFO_BUILD_TIMESTAMP   OPCUA_P_VERSIONSTRING_SEPARATOR \
                                                    OPCUA_P_BUILDINFO_VENDOR_NAME       OPCUA_P_VERSIONSTRING_SEPARATOR \
                                                    OPCUA_P_BUILDINFO_VENDOR_INFO       OPCUA_P_VERSIONSTRING_SEPARATOR \
                                                    OPCUA_P_BUILDINFO_COMPILER

/*============================================================================
* Additional basic headers
*===========================================================================*/
/* configuration switches */
#include <opcua_config.h>

/* basic type mapping */
#include "opcua_p_types.h"

/**********************************************************************************/
/*/  Security Configuration section.                                             /*/
/**********************************************************************************/
#ifndef OPCUA_SUPPORT_SECURITYPOLICY_BASIC128RSA15
# define OPCUA_SUPPORT_SECURITYPOLICY_BASIC128RSA15 OPCUA_CONFIG_YES
#endif  /* OPCUA_SUPPORT_SECURITYPOLICY_BASIC128RSA15 */

#ifndef OPCUA_SUPPORT_SECURITYPOLICY_BASIC256
# define OPCUA_SUPPORT_SECURITYPOLICY_BASIC256      OPCUA_CONFIG_YES
#endif /* OPCUA_SUPPORT_SECURITYPOLICY_BASIC256 */

#ifndef OPCUA_SUPPORT_SECURITYPOLICY_NONE
# define OPCUA_SUPPORT_SECURITYPOLICY_NONE          OPCUA_CONFIG_YES
#endif  /* OPCUA_SUPPORT_SECURITYPOLICY_NONE */

#ifndef OPCUA_SUPPORT_PKI
# define OPCUA_SUPPORT_PKI                          OPCUA_CONFIG_YES
#endif  /* OPCUA_SUPPORT_PKI */

#if OPCUA_SUPPORT_PKI
# ifndef OPCUA_SUPPORT_PKI_OVERRIDE
#  define OPCUA_SUPPORT_PKI_OVERRIDE                OPCUA_CONFIG_NO
# endif /* OPCUA_SUPPORT_PKI_OVERRIDE */
# ifndef OPCUA_SUPPORT_PKI_OPENSSL
#  define OPCUA_SUPPORT_PKI_OPENSSL                 OPCUA_CONFIG_YES
# endif /* OPCUA_SUPPORT_PKI_OPENSSL */
# ifndef OPCUA_SUPPORT_PKI_WIN32
#  if !defined(_WIN32_WCE) && _MSC_VER >= 1400
#   define OPCUA_SUPPORT_PKI_WIN32                  OPCUA_CONFIG_YES
#  else
#   define OPCUA_SUPPORT_PKI_WIN32                  OPCUA_CONFIG_NO
#  endif
# endif /* OPCUA_SUPPORT_PKI_WIN32 */
#endif /* OPCUA_SUPPORT_PKI */

/*============================================================================
 * Memory allocation functions.
 *
 * Note: Realloc and Free behave gracefully if passed a NULL pointer. Changing
 * these functions to a macro call to free will cause problems.
 *===========================================================================*/

/* shortcuts for often used memory functions */
#define OpcUa_Alloc(xSize)                              OpcUa_Memory_Alloc(xSize)
#define OpcUa_ReAlloc(xSrc, xSize)                      OpcUa_Memory_ReAlloc(xSrc, xSize)
#define OpcUa_Free(xSrc)                                OpcUa_Memory_Free(xSrc)
#define OpcUa_MemCpy(xDst, xDstSize, xSrc, xCount)      OpcUa_Memory_MemCpy(xDst, xDstSize, xSrc, xCount)

/* import prototype for direct mapping on memset */
#define OpcUa_MemSet(xDst, xValue, xDstSize)            memset(xDst, xValue, xDstSize)
#define OpcUa_MemSetD(xDst, xValue, xDstSize)           memset(xDst, xValue, xDstSize)

/* import prototype for direct mapping on memcmp */
#define OpcUa_MemCmp(xBuf1, xBuf2, xBufSize)            memcmp(xBuf1, xBuf2, xBufSize)

/*============================================================================
 * String handling functions.
 *===========================================================================*/

/* mapping of ansi string functions on lib functions: */
#define OpcUa_StrCmpA(xStr1, xStr2)                     strcmp(xStr1, xStr2)
#define OpcUa_StrnCmpA(xStr1, xStr2, xCount)            strncmp(xStr1, xStr2, xCount)
#define OpcUa_StriCmpA(xStr1, xStr2)                    stricmp(xStr1, xStr2)
#define OpcUa_StrinCmpA(xStr1, xStr2, xCount)           strnicmp(xStr1, xStr2, xCount)
#define OpcUa_StrLenA(xStr)                             (OpcUa_UInt32)strlen(xStr)
#define OpcUa_StrChrA(xString, xChar)                   strchr(xString, (OpcUa_CharA)xChar)
#define OpcUa_StrrChrA(xString, xChar)                  strrchr(xString, (OpcUa_CharA)xChar)
#define OpcUa_StrStrA(xString, xSubstring)              strstr(xString, xSubstring)

#if OPCUA_USE_SAFE_FUNCTIONS
#define OpcUa_SPrintfA                                  sprintf_s
#else
#define OpcUa_SPrintfA                                  sprintf
#endif

#if OPCUA_USE_SAFE_FUNCTIONS
#define OpcUa_SnPrintfA                                 _snprintf_s
#else
#define OpcUa_SnPrintfA                                 _snprintf
#endif

#if OPCUA_USE_SAFE_FUNCTIONS
#define OpcUa_SScanfA                                   sscanf_s
#else
#define OpcUa_SScanfA                                   sscanf
#endif

/* shortcuts to OpcUa_String functions */
#define OpcUa_StrLen(xStr)                              OpcUa_String_StrLen(xStr)

OPCUA_END_EXTERN_C

#endif /* _OpcUa_PlatformDefs_H_ */
