SET SERVEROUTPUT ON;

BEGIN
   EXECUTE IMMEDIATE 'CREATE TABLE ADVANCE (
      ADV_ID VARCHAR2(7) PRIMARY KEY,
      EMP_ID VARCHAR2(6) CONSTRAINT FK_EMPLY REFERENCES EMPLOYEE(EMP_ID),
      PRL_ID VARCHAR2(6) CONSTRAINT FK_PR REFERENCES PAYROLL(PRL_ID),
      ADV_TYPE VARCHAR2(15) NOT NULL,
      ADV_AMT NUMBER(8,2) NOT NULL CHECK (ADV_AMT >= 0),
      EMI NUMBER(7,2) NOT NULL CHECK (EMI >= 0),
      DUES NUMBER(8,2) NOT NULL CHECK (DUES >= 0),
      ADV_RPY_SDT DATE NOT NULL,
      ADV_STATUS CHAR(1) DEFAULT ''A'' NOT NULL CHECK (ADV_STATUS IN (''A'', ''C''))
   )';

   DBMS_OUTPUT.PUT_LINE('ADVANCE TABLE HAS BEEN CREATED SUCCESSFULLY');
EXCEPTION
   WHEN OTHERS THEN
      DBMS_OUTPUT.PUT_LINE('ERROR: ' || SQLERRM);
END;
/
