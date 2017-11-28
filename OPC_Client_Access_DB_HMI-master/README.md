# OPC_Client_Access_DB_HMI

Plateform: Windows 7 Professionnal SP1 32bits
IDE      : Visual Studio 2008
Language : C++

Test program used to read real-time data from OPC Server (DB852) (Siemens SIMATIC Net Softnet IE V13 SP2) from a S7-317 2PN/DP Siemens PLC
and write this data into an MS Access Database

On startup, m_bNewCoil and m_AckNewCoil are set to false

Scenario:
- PLC sets DB852.X0.0 to 1
- m_bNewCoil is set to true
- and the data read from DB852 (except bits) are written into Access DB.
- m_bNewCoil is reset
- m_bAckNewCoil is set to true
- DB852.X0.1 is set to 1
- PLC read that DB852.X0.1 is set to true and reset  DB852.X0.0
- m_bAckNewCoil is reset

