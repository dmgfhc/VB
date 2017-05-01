create or replace package PKG_MSHEET IS TYPE V_CUR IS REF CURSOR;

-------------------------------------------------------------------------------
-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-------------------------------------------------------------------------------
-- System Name       Template System
-- Sub_System Name   Common
-- Program Name      Master Sheet
-- Program ID        PKG_MSHEET
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
PROCEDURE P_REFER(

     P_EMP_ID         IN  varchar2,
     P_EMP_NAME       IN  varchar2,
     P_PGMID          IN  varchar2,

     P_CUR            OUT V_CUR);

PROCEDURE P_ONEROW (

     P_EMP_ID         IN  varchar2,
     P_PGMID          IN  varchar2,

     P_CUR            OUT V_CUR);

PROCEDURE P_MODIFY (

     iType            IN  VARCHAR2,
	   P_EMP_ID	        IN  ZP_AUTHORITY.EMP_ID%type,
     P_PGMID	        IN  ZP_AUTHORITY.PGMID%type,
     P_INQ    	      IN  ZP_AUTHORITY.INQ%type,
     P_INS            IN  ZP_AUTHORITY.INS%type,
     P_UPD            IN  ZP_AUTHORITY.UPD%type,
     P_DEL            IN  ZP_AUTHORITY.DEL%type,
     P_INS_EMP	      IN  ZP_AUTHORITY.INS_EMP%type,

     P_E_CODE         OUT NUMBER,
     P_E_MSG          OUT VARCHAR2);

END PKG_MSHEET;
/
CREATE OR REPLACE package body PKG_MSHEET is

PROCEDURE P_ONEROW (

          P_EMP_ID   IN  varchar2,
          P_PGMID    IN  varchar2,

          P_CUR      OUT V_CUR  ) IS

BEGIN

    OPEN P_CUR FOR
         SELECT EMP_ID,PGMID,INQ,INS,UPD,DEL,INS_DATE,INS_TIME,INS_EMP
           FROM ZP_AUTHORITY
          WHERE  emp_id = P_emp_id
            AND  pgmid  = P_pgmid;

END P_ONEROW;

PROCEDURE P_REFER (
          P_EMP_ID    IN  varchar2,
          P_EMP_NAME  IN  varchar2,
          P_PGMID     IN  varchar2,

          P_CUR       OUT V_CUR  ) IS

BEGIN

    OPEN P_CUR FOR
		     SELECT A.EMP_ID,A.PGMID,A.INQ,A.INS,A.UPD,A.DEL,A.INS_DATE,A.INS_TIME,A.INS_EMP
           FROM ZP_AUTHORITY A, ZP_EMPLOYEE B
          WHERE a.emp_id   like P_emp_id || '%'
            AND b.emp_name like P_emp_name || '%'
            AND a.pgmid    like P_pgmid || '%'
            AND a.emp_id = b.emp_id ;

END P_REFER;

PROCEDURE P_MODIFY (

     iType            IN  VARCHAR2,
     P_EMP_ID	        IN  ZP_AUTHORITY.EMP_ID%type,
     P_PGMID	        IN  ZP_AUTHORITY.PGMID%type,
     P_INQ    	      IN  ZP_AUTHORITY.INQ%type,
     P_INS            IN  ZP_AUTHORITY.INS%type,
     P_UPD            IN  ZP_AUTHORITY.UPD%type,
     P_DEL            IN  ZP_AUTHORITY.DEL%type,
     P_INS_EMP	      IN  ZP_AUTHORITY.INS_EMP%type,

	   P_E_CODE	        OUT  NUMBER,
	   P_E_MSG	        OUT  VARCHAR2) IS

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

  INSERT INTO ZP_AUTHORITY(
           EMP_ID
         , PGMID
         , INQ
         , INS
         , UPD
         , DEL
         , INS_EMP
         , INS_DATE
         , INS_TIME

  )
  VALUES (
           P_EMP_ID
         , P_PGMID
         , P_INQ
         , P_INS
         , P_UPD
         , P_DEL
         , P_INS_EMP
         , TO_CHAR(SYSDATE,'YYYYMMDD')
         , TO_CHAR(SYSDATE,'HH24MISS')

  );

RETURN;

<<P_UPDATE>>

  UPDATE ZP_AUTHORITY SET

           INQ      = P_INQ
         , INS      = P_INS
         , UPD      = P_UPD
         , DEL      = P_DEL
         , INS_EMP  = P_INS_EMP
         , INS_DATE = TO_CHAR(SYSDATE,'YYYYMMDD')
         , INS_TIME = TO_CHAR(SYSDATE,'HH24MISS')

  WHERE EMP_ID = P_EMP_ID
    AND PGMID  = P_PGMID;
    
  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;    

RETURN;

<<P_DELETE>>

  DELETE FROM ZP_AUTHORITY WHERE EMP_ID = P_EMP_ID AND PGMID = P_PGMID;

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

END PKG_MSHEET;
/
