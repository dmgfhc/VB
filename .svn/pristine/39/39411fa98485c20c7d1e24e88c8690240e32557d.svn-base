CREATE OR REPLACE Package PKG_HSHEET is TYPE V_CUR IS REF CURSOR;
-------------------------------------------------------------------------------
-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-------------------------------------------------------------------------------
-- System Name       Template System
-- Sub_System Name   Common
-- Program Name      Header Sheet
-- Program ID        PKG_HSHEET
-- Document No       Q-00-0010(Specification)
-- Designer          Kim Sung Ho
-- Coder             Kim Sung Ho
-- Date              2003.5.19
-- Description
-------------------------------------------------------------------------------
-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-------------------------------------------------------------------------------
-- VER   DATE     EDITOR       DESCRIPTION
-------------------------------------------------------------------------------
-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-------------------------------------------------------------------------------
 PROCEDURE P_SREFER(
     P_CD_MANA_NO     IN  ZP_CD.CD_MANA_NO%type,

     P_CUR            OUT V_CUR);

PROCEDURE P_SONEROW (
     P_CD_MANA_NO     IN  ZP_CD.CD_MANA_NO%type,
     P_CD             IN  ZP_CD.CD%type,

     P_CUR            OUT V_CUR);

PROCEDURE P_SMODIFY (

     iType	          IN  VARCHAR2,                        -- Type
	   P_CD_MANA_NO     IN  ZP_CD.CD_MANA_NO%type,
     P_CD	            IN  ZP_CD.CD%type,
     P_CD_SHORT_NAME	IN  ZP_CD.CD_SHORT_NAME%type,
     P_CD_NAME	      IN  ZP_CD.CD_NAME%type,
     P_CD_SHORT_ENG   IN  ZP_CD.CD_SHORT_ENG%type,
     P_CD_FULL_ENG	  IN  ZP_CD.CD_FULL_ENG%type,
     P_APLY_STD	      IN  ZP_CD.APLY_STD%type,
     P_INS_EMP	      IN  ZP_CD.INS_EMP%type,

     P_E_CODE         OUT NUMBER,
     P_E_MSG          OUT VARCHAR2);


PROCEDURE P_REFER(
     P_CD_MANA_NO     IN  ZP_CD_MASTER.CD_MANA_NO%type,

     P_CUR            OUT V_CUR);

PROCEDURE P_MODIFY (

     iType	          IN  VARCHAR2,                           -- Type
	   P_CD_MANA_NO	    IN  ZP_CD_MASTER.CD_MANA_NO%type,
     P_CD_MANA_NAME	  IN  ZP_CD_MASTER.CD_MANA_NAME%type,
     P_BIZ_AREA	      IN  ZP_CD_MASTER.BIZ_AREA%type,
     P_CD_LEN	        IN  ZP_CD_MASTER.CD_LEN%type,
     P_CD_DESC	      IN  ZP_CD_MASTER.CD_DESC%type,
     P_INS_EMP	      IN  ZP_CD_MASTER.INS_EMP%type,

     P_E_CODE         OUT NUMBER,
     P_E_MSG          OUT VARCHAR2);

END PKG_HSHEET;
/
CREATE OR REPLACE Package body PKG_HSHEET is
-----------------------------------------------------------------------------------
PROCEDURE P_REFER (
     P_CD_MANA_NO    IN  ZP_CD_MASTER.CD_MANA_NO%type,

     P_CUR           OUT V_cur  ) IS

BEGIN

    OPEN P_CUR FOR
		     SELECT CD_MANA_NO,CD_MANA_NAME,BIZ_AREA,GF_COMNNAMEFIND('Z0001',BIZ_AREA),
                CD_LEN,    CD_DESC,
                INS_DATE,  INS_TIME,    INS_EMP
           FROM ZP_CD_MASTER
          WHERE CD_MANA_NO = P_CD_MANA_NO;

END P_REFER;
-----------------------------------------------------------------------------------
PROCEDURE P_MODIFY (

     iType	          IN  VARCHAR2,                         -- Type
     P_CD_MANA_NO	    IN  ZP_CD_MASTER.CD_MANA_NO%type,
     P_CD_MANA_NAME	  IN  ZP_CD_MASTER.CD_MANA_NAME%type,
     P_BIZ_AREA	      IN  ZP_CD_MASTER.BIZ_AREA%type,
     P_CD_LEN	        IN  ZP_CD_MASTER.CD_LEN%type,
     P_CD_DESC	      IN  ZP_CD_MASTER.CD_DESC%type,
     P_INS_EMP	      IN  ZP_CD_MASTER.INS_EMP%type,

	   P_E_CODE	        OUT NUMBER,
	   P_E_MSG	        OUT VARCHAR2) IS

BEGIN

     P_E_CODE := 0;

     IF iType = 'I' THEN
        GOTO P_INSERT;
     ELSIF iType = 'U' THEN
        GOTO P_UPDATE;
     ELSIF iType = 'D' THEN
        GOTO P_DELETE;
     END IF;

<<P_INSERT>>

  INSERT INTO ZP_CD_MASTER(
           CD_MANA_NO
         , CD_MANA_NAME
         , BIZ_AREA
         , CD_LEN
         , CD_DESC
         , INS_EMP
         , INS_DATE
         , INS_TIME

  )
  VALUES (
           P_CD_MANA_NO
         , P_CD_MANA_NAME
         , P_BIZ_AREA
         , P_CD_LEN
         , P_CD_DESC
         , P_INS_EMP
         , TO_CHAR(SYSDATE,'YYYYMMDD')
         , TO_CHAR(SYSDATE,'HH24MISS')

  );

RETURN;

<<P_UPDATE>>

  UPDATE ZP_CD_MASTER SET

           CD_MANA_NAME = P_CD_MANA_NAME
         , BIZ_AREA     = P_BIZ_AREA
         , CD_LEN       = P_CD_LEN
         , CD_DESC      = P_CD_DESC
         , INS_EMP      = P_INS_EMP
         , INS_DATE     = TO_CHAR(SYSDATE,'YYYYMMDD')
         , INS_TIME     = TO_CHAR(SYSDATE,'HH24MISS')

  WHERE CD_MANA_NO = P_CD_MANA_NO;
  
  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;

RETURN;

<<P_DELETE>>

  DELETE FROM ZP_CD_MASTER WHERE CD_MANA_NO = P_CD_MANA_NO ;

  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;
  
RETURN;

EXCEPTION
	WHEN NO_DATA_FOUND THEN
    P_E_CODE := 1;
		P_E_MSG := 'NO_DATA_FOUND';
		RETURN;
	WHEN DUP_VAL_ON_INDEX THEN
     P_E_CODE := 1;
		 P_E_MSG := 'DUP_VAL_ON_INDEX';
		 RETURN;
  WHEN OTHERS THEN
     P_E_CODE := 1;
		 P_E_MSG := SQLERRM;
		 RETURN;

END P_MODIFY;


PROCEDURE P_SONEROW (

          P_CD_MANA_NO    IN  ZP_CD.CD_MANA_NO%type,
          P_CD            IN  ZP_CD.CD%type,

          P_CUR           OUT V_CUR  ) IS

BEGIN

    OPEN P_CUR FOR
         SELECT CD_MANA_NO,CD,CD_SHORT_NAME,CD_NAME,CD_SHORT_ENG,CD_FULL_ENG,APLY_STD,
                INS_DATE,INS_TIME,INS_EMP
           FROM ZP_CD
          WHERE CD_MANA_NO = P_CD_MANA_NO
            AND CD  = P_CD;

END P_SONEROW;


PROCEDURE P_SREFER (

          P_CD_MANA_NO    IN  ZP_CD.CD_MANA_NO%type,

          P_CUR           OUT V_CUR  ) IS

BEGIN

    OPEN P_CUR FOR

		     SELECT CD_MANA_NO,CD,CD_SHORT_NAME,CD_NAME,CD_SHORT_ENG,CD_FULL_ENG,APLY_STD,
                INS_DATE,INS_TIME,INS_EMP
           FROM ZP_CD
          WHERE CD_MANA_NO like P_CD_MANA_NO || '%';

END P_SREFER;


PROCEDURE P_SMODIFY (

     iType	          IN  VARCHAR2,                        -- Type
	   P_CD_MANA_NO     IN  ZP_CD.CD_MANA_NO%type,
     P_CD	            IN  ZP_CD.CD%type,
     P_CD_SHORT_NAME	IN  ZP_CD.CD_SHORT_NAME%type,
     P_CD_NAME	      IN  ZP_CD.CD_NAME%type,
     P_CD_SHORT_ENG	  IN  ZP_CD.CD_SHORT_ENG%type,
     P_CD_FULL_ENG	  IN  ZP_CD.CD_FULL_ENG%type,
     P_APLY_STD	      IN  ZP_CD.APLY_STD%type,
     P_INS_EMP	      IN  ZP_CD.INS_EMP%type,

	   P_E_CODE	        OUT  NUMBER,
	   P_E_MSG	        OUT  VARCHAR2) IS


BEGIN

     P_E_CODE := 0;

     IF iType = 'I' THEN
        GOTO P_SINSERT;
     ELSIF iType = 'U' THEN
        GOTO P_SUPDATE;
     ELSIF iType = 'D' THEN
        GOTO P_SDELETE;
     END IF;

<<P_SINSERT>>

  INSERT INTO ZP_CD(
           CD_MANA_NO
         , CD
         , CD_SHORT_NAME
         , CD_NAME
         , CD_SHORT_ENG
         , CD_FULL_ENG
         , APLY_STD
         , INS_EMP
         , INS_DATE
         , INS_TIME

  )
  VALUES (
           P_CD_MANA_NO
         , P_CD
         , P_CD_SHORT_NAME
         , P_CD_NAME
         , P_CD_SHORT_ENG
         , P_CD_FULL_ENG
         , P_APLY_STD
         , P_INS_EMP
         , TO_CHAR(SYSDATE,'YYYYMMDD')
         , TO_CHAR(SYSDATE,'HH24MISS')

  );

RETURN;

<<P_SUPDATE>>

  UPDATE ZP_CD SET

           CD_SHORT_NAME = P_CD_SHORT_NAME
         , CD_NAME       = P_CD_NAME
         , CD_SHORT_ENG  = P_CD_SHORT_ENG
         , CD_FULL_ENG   = P_CD_FULL_ENG
         , APLY_STD      = P_APLY_STD
         , INS_EMP       = P_INS_EMP
         , INS_DATE      = TO_CHAR(SYSDATE,'YYYYMMDD')
         , INS_TIME      = TO_CHAR(SYSDATE,'HH24MISS')

  WHERE CD_MANA_NO = P_CD_MANA_NO
    AND CD  = P_CD;
    
  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;    

RETURN;

<<P_SDELETE>>

  DELETE FROM ZP_CD WHERE CD_MANA_NO = P_CD_MANA_NO AND CD = P_CD;

  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;
  
RETURN;

EXCEPTION
	WHEN NO_DATA_FOUND THEN
    P_E_CODE := 1;
		P_E_MSG := 'NO_DATA_FOUND';
		RETURN;
	WHEN DUP_VAL_ON_INDEX THEN
     P_E_CODE := 1;
		 P_E_MSG := 'DUP_VAL_ON_INDEX';
		 RETURN;
  WHEN OTHERS THEN
     P_E_CODE := 1;
		 P_E_MSG := SQLERRM;
		 RETURN;

END P_SMODIFY;

END PKG_HSHEET;
/
