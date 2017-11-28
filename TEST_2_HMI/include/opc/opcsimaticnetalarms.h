/****************************************************************************/
/*     Copyright (C) Siemens AG 2003  All Rights Reserved.                  */
/****************************************************************************/
/*                                                                          */
/*  version: OPCSimaticNetAlarms.h    V 1.00                                */
/*                                                                          */
/*  This header file contains defines                                       */
/*  for using the OPC Alarms and Events SIMATIC NET OPC, Siemens AG.        */
/*  to be included in all source files of user applications                 */
/*                                                                          */
/*                                                                          */
/****************************************************************************/



// Supported categories 
// OPC Server <8.0 
#define     EVENTCATEGORY_S7SIMPLE                              =  2    // Simple Events für folgende S7-Alarme: NOTIFY, ALARM, ALARM_8, ALARM_8P, ALARM_S und ALARM_SQ
#define     EVENTCATEGORY_S7TO                                  = 11    // SIMOTION spezifische Alarme
#define     EVENTCATEGORY_DIAGNOSIS                             = 12    // enthalten die Einträge im Diagnosepuffer der Baugruppe
#define     EVENTCATEGORY_S7CONDITION                           = 13    // Condition Events für folgende S7-Alarme: NOTIFY, ALARM, ALARM_8, ALARM_8P, ALARM_S und ALARM_SQ
#define     EVENTCATEGORY_STATEPATH                             = 14    // Alarme für die Anzeige einer unterbrochene Verbindung zu einem Gerät
#define     EVENTCATEGORY_SNMP_GENERICTRAP                      = 20    // SNMP Trap Event (Generischer Trap)
#define     EVENTCATEGORY_SNMP_SPECIFICTRAP                     = 21    // SNMP Trap Event (Spezifischer Trap)
// OPC Server >=8.0 
#define     EVENTCATEGORY_S7LEVEL                               = 40    // Levelalarm
#define     EVENTCATEGORY_S7DEVIATION                           = 41    // Toleranzalarm
#define     EVENTCATEGORY_S7ROC                                 = 42    // RateOfChange-Alarm
#define     EVENTCATEGORY_S7OFFNORMAL                           = 43    // OffNormal-Alarm, signalisiert Abweichungen von einem Normalzustand
#define     EVENTCATEGORY_S7TRIP                                = 44    // Trip-Alarm
#define     EVENTCATEGORY_S7COS                                 = 45    // ChangeOfState-Alarm
#define     EVENTCATEGORY_S7DEVICEFAILURE                       = 46    // Gerätefehler
#define     EVENTCATEGORY_S7SYSTEMFAILURE                       = 47    // Systemfehler
#define     EVENTCATEGORY_S7SYSTEMMESSAGE                       = 60    // Systemnachricht
#define     EVENTCATEGORY_S7USERMESSAGE                         = 60    // Anwendernachricht


// For EVENTCATEGORY_S7

#define     EVENT_ATTR_S7_MSGTEXT                               =    -1 // Messagetext
#define     EVENT_ATTR_S7_ALARMSTATE                            =    -2 // Alarmstatus
#define     EVENT_ATTR_S7_TYPE                                  =    -5 // EVENT_TYPE_CONDITION für EVENTCATEGORY_S7CONDITION und
                                                                        // EVENT_TYPE_SIMPLE für EVENTCATEGORYSIMPLE
#define     EVENT_ATTR_S7_COMMENT                               =   -13 // Comment
#define     EVENT_ATTR_S7_PCTIME                                =  6000 // Zeitstempel des SIMATIC NET Core-Servers
#define     EVENT_ATTR_S7_S7TIME                                =  6001 // Zeitstempel der S7
#define     EVENT_ATTR_S7_STATE                                 =  6002 // S7_Status im AG
#define     EVENT_ATTR_S7_ACK_STATE                             =  6003 // Quittungszustand
#define     EVENT_ATTR_S7_EVENT_STATE                           =  6004 // Ereigniszustand
#define     EVENT_ATTR_S7_NO_DATA                               =  6005 // Anzahl der Begleitwerte
#define     EVENT_ATTR_S7_DATA0_DATATYPE                        =  6006 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA0_DATA_VALUE_LEN                  =  6007 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA0                                 =  6008 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA1_DATATYPE                        =  6010 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA1_DATA_VALUE_LEN                  =  6011 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA1                                 =  6012 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA2_DATATYPE                        =  6014 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA2_DATA_VALUE_LEN                  =  6015 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA2                                 =  6016 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA3_DATATYPE                        =  6018 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA3_DATA_VALUE_LEN                  =  6019 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA3                                 =  6020 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA4_DATATYPE                        =  6022 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA4_DATA_VALUE_LEN                  =  6023 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA4                                 =  6024 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA5_DATATYPE                        =  6026 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA5_DATA_VALUE_LEN                  =  6027 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA5                                 =  6028 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA6_DATATYPE                        =  6030 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA6_DATA_VALUE_LEN                  =  6031 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA6                                 =  6032 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA7_DATATYPE                        =  6034 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA7_DATA_VALUE_LEN                  =  6035 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA7                                 =  6036 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA8_DATATYPE                        =  6038 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA8_DATA_VALUE_LEN                  =  6039 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA8                                 =  6040 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_DATA9_DATATYPE                        =  6042 // S7-Spezifischer Datentyp des Begleitwertes
#define     EVENT_ATTR_S7_DATA9_DATA_VALUE_LEN                  =  6043 // Länge des Begleitwerts
#define     EVENT_ATTR_S7_DATA9                                 =  6044 // Datenpuffer des Begleitwerts
#define     EVENT_ATTR_S7_EVENT_EVENTID                         =  6046 // Alarmnummer
#define     EVENT_ATTR_S7_EVENT_SUBEVENTID                      =  6047 // Alarmunternummer

#define     EVENT_ATTR_S7_CONNECTION                            =  6050 // S7-Verbindungsname
#define     EVENT_ATTR_S7_AREA                                  =  6051 // Projektierte Area

#define     EVENT_ATTR_S7_INFOTEXT                              =  6060 // Projektierter Infotext, sprachabhängig
#define     EVENT_ATTR_S7_TEXT1                                 =  6061 // Projektierter Zusatztext 1, sprachabhängig
#define     EVENT_ATTR_S7_TEXT2                                 =  6062 // Projektierter Zusatztext 2, sprachabhängig
#define     EVENT_ATTR_S7_TEXT3                                 =  6063 // Projektierter Zusatztext 3 sprachabhängig
#define     EVENT_ATTR_S7_TEXT4                                 =  6064 // Projektierter Zusatztext 4, sprachabhängig
#define     EVENT_ATTR_S7_TEXT5                                 =  6065 // Projektierter Zusatztext 5 sprachabhängig
#define     EVENT_ATTR_S7_TEXT6                                 =  6066 // Projektierter Zusatztext 6 sprachabhängig
#define     EVENT_ATTR_S7_TEXT7                                 =  6067 // Projektierter Zusatztext 7, sprachabhängig
#define     EVENT_ATTR_S7_TEXT8                                 =  6068 // Projektierter Zusatztext 8 sprachabhängig
#define     EVENT_ATTR_S7_TEXT9                                 =  6069 // Projektierter Zusatztext 9, sprachabhängig


// S7 Datatypes
#define     S7_DATATYPE_ERROR                                   = 0x00  // Fehler
#define     S7_DATATYPE_BOOLEAN                                 = 0x03  // Boolean
#define     S7_DATATYPE_INTEGER                                 = 0x05  // Integer
#define     S7_DATATYPE_FLOAT                                   = 0x07  // Float
#define     S7_DATATYPE_OCTET_STRING                            = 0x09  // Octet-String
#define     S7_DATATYPE_BITSTRING                               = 0x04  // Bit-String
#define     S7_DATATYPE_DATE                                    = 0x30  // Datum
#define     S7_DATATYPE_TIME_OF_DAY                             = 0x31  // Zeit (ms seit Tagesbegin)
#define     S7_DATATYPE_TIME                                    = 0x32  // Zeit (ms)
#define     S7_DATATYPE_S5TIME                                  = 0x33  // Zeit (BCD-Codiert)

// For EVENTCATEGORY_S7TO
#define     EVENT_ATTR_S7TO_STATE                               = -2    // STATE
#define     EVENT_ATTR_S7TO_TYPE                                = -5    // IEVENT_TYPE_CONDITION
#define     EVENT_ATTR_S7TO_ALARMID                             = 5100  // Eindeutige Alarmnummer
#define     EVENT_ATTR_S7TO_INSTANCEID                          = 5101  // RT-ID der Objekt-Instanz
#define     EVENT_ATTR_S7TO_SEQUENCENO                          = 5102  // Laufende Nummer
#define     EVENT_ATTR_S7TO_GUID                                = 5103  // RT-GUID des TO-Typs
#define     EVENT_ATTR_S7TO_ALARMTYPE                           = 5104  // Simotion-spezifischer Zusatzwert
#define     EVENT_ATTR_S7TO_ACKTYPE                             = 5105  // Simotion-spezifischer Zusatzwert
#define     EVENT_ATTR_S7TO_TIMESTAMP                           = 5106  // Zeitstempel der S7
#define     EVENT_ATTR_S7TO_TIMESTAMPPC                         = 5107  // Zeitstempel des SIMATIC NET Core-Servers,
#define     EVENT_ATTR_S7TO_PARAMDATA                           = 5108  // Simotion-spezifischer Zusatzwert

// For EVENTCATEGORY_DIAGNOSIS
#define     EVENT_ATTR_DIAGNOSIS_MSGTEXT                        =   -1  // Messagetext
#define     EVENT_ATTR_DIAGNOSIS_ALARMSTATE                     =   -2  // Alarmstatus
#define     EVENT_ATTR_DIAGNOSIS_TYPE                           =   -5  // IEVENT_TYPE_SIMPLE
#define     EVENT_ATTR_DIAGNOSIS_SEVERITY                       =   -6  // Severity
#define     EVENT_ATTR_DIAGNOSIS_CATEGORYID                     =   -9  // Category Identifier
#define     EVENT_ATTR_DIAGNOSIS_CATEGORYDESC                   =  -10  // Kategorie Beschreibung
#define     EVENT_ATTR_DIAGNOSIS_COMMENT                        =  -13  // Comment
#define     EVENT_ATTR_DIAGNOSIS_S7_PCTIME                      = 6000  // Zeitstempel des SIMATIC NET Core-Servers
#define     EVENT_ATTR_DIAGNOSIS_S7_S7TIME                      = 6001  // Zeitstempel der S7
#define     EVENT_ATTR_DIAGNOSIS_S7_DIAGNOSISID                 = 6048  // Diagnoseereignisnummer, analog EventID
#define     EVENT_ATTR_DIAGNOSIS_S7_DIAGNOSISDATA               = 6049  // Informationen zum Diagnoseereignis bzw. zu dessen Wirkung, Datenlänge abhängig von der Ereignisklasse
#define     EVENT_ATTR_DIAGNOSIS_CONNECTION                     = 6050  // S7-Verbindungsname
#define     EVENT_ATTR_DIAGNOSIS_AREA                           = 6051  // Projektierte Area
#define     EVENT_ATTR_DIAGNOSIS_INFOTEXT                       = 6060  // Projektierter Infotext, sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT1                          = 6061  // Projektierter Zusatztext 1, sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT2                          = 6062  // Projektierter Zusatztext 2, sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT3                          = 6063  // Projektierter Zusatztext 3 sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT4                          = 6064  // Projektierter Zusatztext 4, sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT5                          = 6065  // Projektierter Zusatztext 5 sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT6                          = 6066  // Projektierter Zusatztext 6 sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT7                          = 6067  // Projektierter Zusatztext 7, sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT8                          = 6068  // Projektierter Zusatztext 8 sprachabhängig
#define     EVENT_ATTR_DIAGNOSIS_TEXT9                          = 6069  // Projektierter Zusatztext 9, sprachabhängig


// For EVENTCATEGORY_STATEPATH
#define     EVENT_ATTR_STATEPATH_MSGTEXT                        = -1    // Messagetext
#define     EVENT_ATTR_STATEPATH_STATE                          = -2    // Statepath Status
#define     EVENT_ATTR_STATEPATH_TYPE                           = -5    // IEVENT_TYPE_CONDITION
#define     EVENT_ATTR_STATEPATH_SEVERITY                       = -6    // Severity
#define     EVENT_ATTR_STATEPATH_CATEGORYID                     = -9    // Category Identifier
#define     EVENT_ATTR_STATEPATH_CATEGORYDESC                   = -10   // Kategorie Beschreibung
#define     EVENT_ATTR_STATEPATH_COMMENT                        = -13   // Comment
#define     EVENT_ATTR_STATEPATH_CONNECTION                     =  6050 // S7-Verbindungsname

// For EVENTCATEGORY_SNMP_GENERICTRAP and EVENTCATEGORY_SNMP_SPECIFICTRAP
#define     EVENT_ATTR_SNMP_TYPE                                = -5    // IEEVENT_TYPE_SIMPLE
#define     EVENT_ATTR_SNMP_PCTIME                              = 6100  // Zeitpunkt, an dem der OPC Event-Server die Meldung bekommen hat.
#define     EVENT_ATTR_SNMP_TIMESTAMP                           = 6101  // Zeitpunkt, zu dem die Meldung auf dem Partnergerät erzeugt wurde.
#define     EVENT_ATTR_SNMP_TRAP                                = 6102  // innerhalb des Geräts eindeutige Trap Nummer
#define     EVENT_ATTR_SNMP_ENTERPRISE                          = 6103  // Gerätespezifischer Trap
#define     EVENT_ATTR_SNMP_AGENTADDRESS                        = 6104  // IP-Adresse Agent
#define     EVENT_ATTR_SNMP_SOURCEADDRESS                       = 6105  // IP-Adresse Trap-Source
#define     EVENT_ATTR_SNMP_VARBINDINGS                         = 6106  // variableBindings des gesendeten Traps

// For EVENTCATEGORY_S7SIMPLE and EVENTCATEGORY_S7CONDITION
#define     EVENT_TYPE_SIMPLE                                   = 0x00000010,  // Simple
#define     EVENT_TYPE_CONDITION                                = 0x00000020,  // Conditional



