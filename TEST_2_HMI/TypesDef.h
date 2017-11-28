//////////////////////////////////////////////////////////////////////
//
// TypesDef.h : DataTypes definition
//
//////////////////////////////////////////////////////////////////////

#pragma once

typedef struct {

	//DB852
	bool		bNewCoil;           // DB852.DBX0.0
	bool		bAckNewCoil;        // DB852.DBX0.1
	char	   *strOrderName;       // DB852.DBB2   (STRING[50])
	char	   *strCustomerName;    // DB852.DBB54  (STRING[50])	
	char	   *strCoilNumber;      // DB852.DBB106 (STRING[50])
	char	   *strNuance;          // DB852.DBB158 (STRING[10])
	char	   *strEtat;            // DB852.DBB170 (STRING[10])
	long		lngThickness;       // DB852.DBW182 (0.01mm)
	long		lngWidth;           // DB852.DBW184 (0.1mm)
	long		lngEModul;		    // DB852.DBD186 (N/mm2)
	long		lngYieldStrength;   // DB852.DBW190 (N/mm2)
	char	   *strCassetteName;    // DB852.DBB192 (STRING[20])

} typRecipe;