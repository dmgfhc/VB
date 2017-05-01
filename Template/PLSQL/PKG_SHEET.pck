CREATE OR REPLACE Package PKG_SHEET is TYPE V_CUR IS REF CURSOR;

-------------------------------------------------------------------------------
-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-------------------------------------------------------------------------------
-- System Name       Template System
-- Sub_System Name   Common
-- Program Name      Sheet
-- Program ID        PKG_SHEET
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
PROCEDURE P_REFER(P_CUR OUT V_CUR);

PROCEDURE P_ONEROW(

	   P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,

     P_cur            OUT V_cur);

PROCEDURE P_MODIFY (

     iType	          IN  VARCHAR2,                         -- Type
	   P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,          -- 01.EMP_ID
     P_EMP_NAME	      IN  ZP_EMPLOYEE.EMP_NAME%type,        -- 02.EMP_NAME
     P_DEPT	          IN  ZP_EMPLOYEE.DEPT%type,            -- 03.DEPT
     P_PHONE_OFFICE	  IN  ZP_EMPLOYEE.PHONE_OFFICE%type,    -- 04.PHONE_OFFICE
     P_PHONE_MOBILE	  IN  ZP_EMPLOYEE.PHONE_MOBILE%type,    -- 05.PHONE_MOBILE
     P_PHONE_HOME	    IN  ZP_EMPLOYEE.PHONE_HOME%type,      -- 06.PHONE_HOME
     P_PASSWORD	      IN  ZP_EMPLOYEE.PASSWORD%type,        -- 07.PASSWORD
     P_E_MAIL	        IN  ZP_EMPLOYEE.E_MAIL%type,          -- 08.E_MAIL
     P_DESCRIPTION	  IN  ZP_EMPLOYEE.DESCRIPTION%type,     -- 09.DESCRIPTION
     P_SUPERVISOR	    IN  ZP_EMPLOYEE.SUPERVISOR%type,      -- 10.SUPERVISOR
     P_INS_EMP	      IN  ZP_EMPLOYEE.INS_EMP%type,         -- 11.INS_EMP

     P_E_CODE         OUT NUMBER,
     P_E_MSG          OUT VARCHAR2);

end PKG_SHEET;
/
CREATE OR REPLACE Package body PKG_SHEET is

PROCEDURE P_REFER (P_CUR out V_CUR) IS

BEGIN

     OPEN P_CUR FOR
		      SELECT EMP_ID,EMP_NAME,DEPT,PHONE_OFFICE,PHONE_MOBILE,PHONE_HOME,
		  		       PASSWORD,E_MAIL,DESCRIPTION,SUPERVISOR,INS_DATE,INS_TIME,INS_EMP
            FROM zp_employee;

END P_REFER;

PROCEDURE P_ONEROW (

     P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,          -- 01.EMP_ID

	   P_CUR  	        OUT V_CUR) IS

BEGIN

     OPEN P_CUR FOR
          select EMP_ID,EMP_NAME,DEPT,PHONE_OFFICE,PHONE_MOBILE,PHONE_HOME,
		  		       PASSWORD,E_MAIL,DESCRIPTION,SUPERVISOR,INS_DATE,INS_EMP
            from zp_employee
           where emp_id = P_EMP_ID;

END P_ONEROW;


PROCEDURE P_MODIFY (

     iType            IN  VARCHAR2,
     P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,          -- 01.EMP_ID
     P_EMP_NAME	      IN  ZP_EMPLOYEE.EMP_NAME%type,        -- 02.EMP_NAME
     P_DEPT	          IN  ZP_EMPLOYEE.DEPT%type,            -- 03.DEPT
     P_PHONE_OFFICE	  IN  ZP_EMPLOYEE.PHONE_OFFICE%type,    -- 04.PHONE_OFFICE
     P_PHONE_MOBILE	  IN  ZP_EMPLOYEE.PHONE_MOBILE%type,    -- 05.PHONE_MOBILE
     P_PHONE_HOME	    IN  ZP_EMPLOYEE.PHONE_HOME%type,      -- 06.PHONE_HOME
     P_PASSWORD	      IN  ZP_EMPLOYEE.PASSWORD%type,        -- 07.PASSWORD
     P_E_MAIL	        IN  ZP_EMPLOYEE.E_MAIL%type,          -- 08.E_MAIL
     P_DESCRIPTION	  IN  ZP_EMPLOYEE.DESCRIPTION%type,     -- 09.DESCRIPTION
     P_SUPERVISOR	    IN  ZP_EMPLOYEE.SUPERVISOR%type,      -- 10.SUPERVISOR
     P_INS_EMP	      IN  ZP_EMPLOYEE.INS_EMP%type,         -- 11.INS_EMP

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

  INSERT INTO ZP_EMPLOYEE(
           EMP_ID
         , EMP_NAME
         , DEPT
         , PHONE_OFFICE
         , PHONE_MOBILE
         , PHONE_HOME
         , PASSWORD
         , E_MAIL
         , DESCRIPTION
         , SUPERVISOR
         , INS_EMP
         , INS_DATE
         , INS_TIME

  )
  values (
           P_EMP_ID
         , P_EMP_NAME
         , P_DEPT
         , P_PHONE_OFFICE
         , P_PHONE_MOBILE
         , P_PHONE_HOME
         , P_PASSWORD
         , P_E_MAIL
         , P_DESCRIPTION
         , P_SUPERVISOR
         , P_INS_EMP
         , TO_CHAR(SYSDATE,'YYYYMMDD')
         , TO_CHAR(SYSDATE,'HH24MISS')

  );

RETURN;

<<P_UPDATE>>

  UPDATE ZP_EMPLOYEE SET
           EMP_NAME     = P_EMP_NAME
         , DEPT         = P_DEPT
         , PHONE_OFFICE = P_PHONE_OFFICE
         , PHONE_MOBILE = P_PHONE_MOBILE
         , PHONE_HOME   = P_PHONE_HOME
         , PASSWORD     = P_PASSWORD
         , E_MAIL       = P_E_MAIL
         , DESCRIPTION  = P_DESCRIPTION
         , SUPERVISOR   = P_SUPERVISOR
         , INS_EMP      = P_INS_EMP
         , INS_DATE     = TO_CHAR(SYSDATE,'YYYYMMDD')
         , INS_TIME     = TO_CHAR(SYSDATE,'HH24MISS')

         WHERE EMP_ID = P_EMP_ID;
         
  IF SQL%ROWCOUNT = 0 THEN
     RAISE NO_DATA_FOUND;
  END IF;         

RETURN;

<<P_DELETE>>

  DELETE FROM ZP_EMPLOYEE WHERE EMP_ID = P_EMP_ID;

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


END PKG_SHEET;
/
