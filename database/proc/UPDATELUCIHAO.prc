CREATE OR REPLACE PROCEDURE UPDATELUCIHAO (
  P_GAOLU IN NUMBER,P_ZDSJ IN DATE
)
IS
  CURSOR   C_CUR   IS   SELECT  ZDSJ,DKSJ
  FROM  DDLUCI
  WHERE  ZDSJ >= P_ZDSJ AND GAOLU = P_GAOLU
  ORDER BY ZDSJ
  FOR UPDATE OF  LUCI;
   v_luci varchar2(200);
   v_maxluci varchar2(200);
BEGIN
  SELECT MAX(LUCI) INTO v_maxluci FROM DDLUCI WHERE DKSJ IS NOT NULL AND ZDSJ < P_ZDSJ AND GAOLU=P_GAOLU;
  FOR r1 IN C_CUR LOOP
     IF r1.DKSJ IS NULL THEN
             v_luci:=null;
        ELSE 
            BEGIN
                IF v_maxluci IS NULL THEN
                  v_maxluci:='1' || TO_CHAR(P_GAOLU) || '0000000';
                END IF;
                v_maxluci:=TO_CHAR(TO_NUMBER(v_maxluci)+1);
                v_luci:=v_maxluci;
/*                exception
                 when others then
                    null;*/
           END;
        END IF;
       
        UPDATE   DDLUCI
        SET   LUCI=v_luci
        WHERE   CURRENT   OF   C_CUR;

  END LOOP;
  COMMIT;
END UPDATELUCIHAO;
/
