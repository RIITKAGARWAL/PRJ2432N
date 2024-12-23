SET SERVEROUTPUT ON;

BEGIN
   EXECUTE IMMEDIATE 'CREATE TABLE COMPLIANCE (
   COMP_ID VARCHAR2(7) PRIMARY KEY,
   EMP_ID VARCHAR2(6) CONSTRAINT FK_EM REFERENCES EMPLOYEE(EMP_ID),
   PRL_ID VARCHAR2(6) CONSTRAINT FK_PRL REFERENCES PAYROLL(PRL_ID),
	TDS DECIMAL(6,2) DEFAULT 0 CHECK(TDS >= 0),
   PF DECIMAL(6,2) DEFAULT 0 CHECK(PF >= 0),
   ESI DECIMAL(6,2) DEFAULT 0 CHECK(ESI >= 0),
   PROF_TAX DECIMAL(6,2) DEFAULT 0 CHECK(PROF_TAX >= 0),
   TAXABLE_INC DECIMAL(6,2) NOT NULL CHECK(TAXABLE_INC >= 0)
)';

   DBMS_OUTPUT.PUT_LINE('COMPLIANCE TABLE HAS BEEN CREATED SUCCESSFULLY');
EXCEPTION
   WHEN OTHERS THEN
      DBMS_OUTPUT.PUT_LINE('ERROR: ' || SQLERRM);
END;
/