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

#ifndef _OpcUa_P_PKI_H_
#define _OpcUa_P_PKI_H_ 1

OPCUA_BEGIN_EXTERN_C

/**
  @brief The supported PKIs.
*/
typedef enum _OpcUa_P_PKI_Types
{
    OpcUa_Invalid_PKI   = 0,
    OpcUa_NO_PKI        = 1,
    OpcUa_Override      = 2,
    OpcUa_OpenSSL_PKI   = 3,
    OpcUa_Win32_PKI     = 4
} OpcUa_P_PKI_Types;

/* WIN32 PKI specific flag */
#define WIN32_PKI_USERSTORE 1
#define WIN32_PKI_MACHINESTORE 2
#define WIN32_PKI_SERVICESSTORE 3

/**
  @brief The openssl pki config.
  */
struct _OpcUa_P_OpenSSL_CertificateStore_Config
{
    /*! @brief used PKI type. */
    OpcUa_P_PKI_Types   PkiType;

    /*! @brief The certificate store location. */
    OpcUa_StringA       CertificateTrustListLocation;

    /*! @brief The certificate revocation list. */
    OpcUa_StringA       CertificateRevocationListLocation;

    /*! @brief PKI-specific flags. */
    OpcUa_UInt32        Flags;

    /*! @brief External PKIProvider IF to override default implementation. Checked when Configuration name is "Override" */
    OpcUa_Void*         Override;
};
typedef struct _OpcUa_P_OpenSSL_CertificateStore_Config OpcUa_P_OpenSSL_CertificateStore_Config;


/**
  @brief The certificate und key format enumeration.
*/
typedef enum _OpcUa_P_FileFormat
{
    OpcUa_Crypto_Encoding_Invalid   = 0,
    OpcUa_Crypto_Encoding_DER       = 1,
    OpcUa_Crypto_Encoding_PEM       = 2,
    OpcUa_Crypto_Encoding_PKCS12    = 3
}
OpcUa_P_FileFormat;

/**
  @brief Loads a X.509 certificate from the specified file.
  */
OpcUa_StatusCode OpcUa_P_OpenSSL_X509_LoadFromFile(
    OpcUa_StringA               fileName,
    OpcUa_P_FileFormat          fileFormat,
    OpcUa_StringA               sPassword,      /* optional: just for OpcUa_PKCS12 */
    OpcUa_ByteString*           pCertificate);

/**
  @brief Loads a RSA private key from the specified file.
  */
OpcUa_StatusCode OpcUa_P_OpenSSL_RSA_LoadPrivateKeyFromFile(
    OpcUa_StringA           privateKeyFile,
    OpcUa_P_FileFormat      fileFormat,
    OpcUa_StringA           password,
    OpcUa_ByteString*       pPrivateKey);

OPCUA_END_EXTERN_C

#endif
