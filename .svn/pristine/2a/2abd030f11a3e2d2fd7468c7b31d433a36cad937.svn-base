CREATE OR REPLACE Package PKG_TABSHEET is TYPE V_CUR IS REF CURSOR;

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
PROCEDURE P_REFER1(P_CUR OUT V_CUR);

PROCEDURE P_ONEROW1(

	   P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,

     P_cur            OUT V_cur);

PROCEDURE P_MODIFY1 (

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


PROCEDURE P_REFER2(P_CUR OUT V_CUR);

PROCEDURE P_ONEROW2(

     P_EMP_ID         IN  varchar2,
     P_PGMID          IN  varchar2,

     P_CUR            OUT V_CUR);

PROCEDURE P_MODIFY2 (

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


END PKG_TabSheet;
/
CREATE OR REPLACE Package body PKG_TabSheet is

PROCEDURE P_REFER1 (P_CUR out V_CUR) IS

BEGIN

     OPEN P_CUR FOR
		      SELECT EMP_ID,EMP_NAME,DEPT,PHONE_OFFICE,PHONE_MOBILE,PHONE_HOME,
		  		       PASSWORD,E_MAIL,DESCRIPTION,SUPERVISOR,INS_DATE,INS_TIME,INS_EMP
            FROM zp_employee;

END P_REFER1;

PROCEDURE P_ONEROW1 (

     P_EMP_ID	        IN  ZP_EMPLOYEE.EMP_ID%type,          -- 01.EMP_ID

	   P_CUR  	        OUT V_CUR) IS

BEGIN

     OPEN P_CUR FOR
          select EMP_ID,EMP_NAME,DEPT,PHONE_OFFICE,PHONE_MOBILE,PHONE_HOME,
		  		       PASSWORD,E_MAIL,DESCRIPTION,SUPERVISOR,INS_DATE,INS_EMP
            from zp_employee
           where emp_id = P_EMP_ID;

END P_ONEROW1;


PROCEDURE P_MODIFY1 (

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
        GOTO P_INSERT1;
     ELSIF iType = 'U' THEN
        GOTO P_UPDATE1;
     ELSIF iType = 'D' THEN
        GOTO P_DELETE1;
     END IF;

<<P_INSERT1>>

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
  VALUES (
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

<<P_UPDATE1>>

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

<<P_DELETE1>>

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


END P_MODIFY1;


PROCEDURE P_ONEROW2 (

          P_EMP_ID   IN  varchar2,
          P_PGMID    IN  varchar2,

          P_CUR      OUT V_CUR  ) IS

BEGIN

    OPEN P_CUR FOR
         SELECT EMP_ID,PGMID,INQ,INS,UPD,DEL,INS_DATE,INS_TIME,INS_EMP
           FROM ZP_AUTHORITY
          WHERE  emp_id = P_emp_id
            AND  pgmid  = P_pgmid;

END P_ONEROW2;

PROCEDURE P_REFER2 (P_CUR out V_CUR) IS

BEGIN

    OPEN P_CUR FOR
		     SELECT A.EMP_ID,A.PGMID,A.INQ,A.INS,A.UPD,A.DEL,A.INS_DATE,A.INS_TIME,A.INS_EMP
           FROM ZP_AUTHORITY A, ZP_EMPLOYEE B
          WHERE A.EMP_ID = B.EMP_ID;

END P_REFER2;

PROCEDURE P_MODIFY2 (

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
        GOTO P_INSERT2;
     ELSIF iType = 'U' THEN
        GOTO P_UPDATE2;
     ELSIF iType = 'D' THEN
        GOTO P_DELETE2;
     END IF;


<<P_INSERT2>>

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
  values (
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

<<P_UPDATE2>>

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

<<P_DELETE2>>

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

END P_MODIFY2;

END PKG_TABSHEET;
/
