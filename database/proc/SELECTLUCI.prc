CREATE OR REPLACE PROCEDURE "SELECTLUCI2" (RQ IN DATE,
                                        BANCI IN VARCHAR,
                                        BANLUCI IN NUMBER,
                                        G1LUCI OUT DDLUCI.LUCI%TYPE,
                                        G1ZDSJ OUT DDLUCI.ZDSJ%TYPE,
                                        G1DGSJ OUT DDLUCI.DGSJ%TYPE,
                                        G1DKSJ OUT  DDLUCI.DKSJ%TYPE,
                                        G1TZSJ OUT  DDLUCI.TZSJ%TYPE,
                                        G1WDSJ OUT  DDLUCI.WDSJ%TYPE,
                                        G1QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G1FELIANG OUT DDLUCI.FELIANG%TYPE,
                                        G1FEC OUT DDLUCI.FEC%TYPE,
                                        G1FESI OUT  DDLUCI.FESI%TYPE,
                                        G1FEMN OUT  DDLUCI.FEMN%TYPE,
                                        G1FEP OUT  DDLUCI.FEP%TYPE,
                                        G1FES OUT  DDLUCI.FES%TYPE,
                                        G1FETI OUT  DDLUCI.FETI%TYPE,
                                        G1ZHASIO2 OUT  DDLUCI.ZHASIO2%TYPE,
                                        G1ZHACAO OUT  DDLUCI.ZHACAO%TYPE,
                                        G1ZHAMGO OUT  DDLUCI.ZHAMGO%TYPE,
                                        G1ZHAAL2O3 OUT  DDLUCI.ZHAAL2O3%TYPE,
                                        G1ZHAS OUT  DDLUCI.ZHAS%TYPE,
                                        G1ZHATIO2 OUT  DDLUCI.ZHATIO2%TYPE,
                                        G1ZHAR2 OUT  DDLUCI.ZHAR2%TYPE,
                                        G2LUCI OUT  DDLUCI.LUCI%TYPE,
                                        G2ZDSJ OUT  DDLUCI.ZDSJ%TYPE,
                                        G2DGSJ OUT  DDLUCI.DGSJ%TYPE,
                                        G2DKSJ OUT  DDLUCI.DKSJ%TYPE,
                                        G2TZSJ OUT  DDLUCI.TZSJ%TYPE,
                                        G2WDSJ OUT  DDLUCI.WDSJ%TYPE,
                                        G2QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G2FELIANG OUT  DDLUCI.FELIANG%TYPE,
                                        G2FEC OUT  DDLUCI.FEC%TYPE,
                                        G2FESI OUT  DDLUCI.FESI%TYPE,
                                        G2FEMN OUT  DDLUCI.FEMN%TYPE,
                                        G2FEP OUT DDLUCI.FEP%TYPE,
                                        G2FES OUT DDLUCI.FES%TYPE,
                                        G2FETI OUT DDLUCI.FETI%TYPE,
                                        G2ZHASIO2 OUT DDLUCI.ZHASIO2%TYPE,
                                        G2ZHACAO OUT DDLUCI.ZHACAO%TYPE,
                                        G2ZHAMGO OUT DDLUCI.ZHAMGO%TYPE,
                                        G2ZHAAL2O3 OUT DDLUCI.ZHAAL2O3%TYPE,
                                        G2ZHAS OUT DDLUCI.ZHAS%TYPE,
                                        G2ZHATIO2 OUT DDLUCI.ZHATIO2%TYPE,
                                        G2ZHAR2 OUT DDLUCI.ZHAR2%TYPE,
                                        G3LUCI OUT  DDLUCI.LUCI%TYPE,
                                        G3ZDSJ OUT  DDLUCI.ZDSJ%TYPE,
                                        G3DGSJ OUT  DDLUCI.DGSJ%TYPE,
                                        G3DKSJ OUT  DDLUCI.DKSJ%TYPE,
                                        G3TZSJ OUT  DDLUCI.TZSJ%TYPE,
                                        G3WDSJ OUT  DDLUCI.WDSJ%TYPE,
                                        G3QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G3FELIANG OUT DDLUCI.FELIANG%TYPE,
                                        G3FEC OUT DDLUCI.FEC%TYPE,
                                        G3FESI OUT DDLUCI.FESI%TYPE,
                                        G3FEMN OUT DDLUCI.FEMN%TYPE,
                                        G3FEP OUT DDLUCI.FEP%TYPE,
                                        G3FES OUT DDLUCI.FES%TYPE,
                                        G3FETI OUT DDLUCI.FETI%TYPE,
                                        G3ZHASIO2 OUT DDLUCI.ZHASIO2%TYPE,
                                        G3ZHACAO OUT DDLUCI.ZHACAO%TYPE,
                                        G3ZHAMGO OUT DDLUCI.ZHAMGO%TYPE,
                                        G3ZHAAL2O3 OUT DDLUCI.ZHAAL2O3%TYPE,
                                        G3ZHAS OUT DDLUCI.ZHAS%TYPE,
                                        G3ZHATIO2 OUT DDLUCI.ZHATIO2%TYPE,
                                        G3ZHAR2 OUT DDLUCI.ZHAR2%TYPE,
                                        G4LUCI OUT DDLUCI.LUCI%TYPE,
                                        G4ZDSJ OUT DDLUCI.ZDSJ%TYPE,
                                        G4DGSJ OUT DDLUCI.DGSJ%TYPE,
                                        G4DKSJ OUT DDLUCI.DKSJ%TYPE,
                                        G4TZSJ OUT DDLUCI.TZSJ%TYPE,
                                        G4WDSJ OUT DDLUCI.WDSJ%TYPE,
                                        G4QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G4FELIANG OUT DDLUCI.FELIANG%TYPE,
                                        G4FEC OUT DDLUCI.FEC%TYPE,
                                        G4FESI OUT DDLUCI.FESI%TYPE,
                                        G4FEMN OUT DDLUCI.FEMN%TYPE,
                                        G4FEP OUT DDLUCI.FEP%TYPE,
                                        G4FES OUT DDLUCI.FES%TYPE,
                                        G4FETI OUT DDLUCI.FETI%TYPE,
                                        G4ZHASIO2 OUT DDLUCI.ZHASIO2%TYPE,
                                        G4ZHACAO OUT DDLUCI.ZHACAO%TYPE,
                                        G4ZHAMGO OUT DDLUCI.ZHAMGO%TYPE,
                                        G4ZHAAL2O3 OUT DDLUCI.ZHAAL2O3%TYPE,
                                        G4ZHAS OUT DDLUCI.ZHAS%TYPE,
                                        G4ZHATIO2 OUT DDLUCI.ZHATIO2%TYPE,
                                        G4ZHAR2 OUT DDLUCI.ZHAR2%TYPE,
                                        G5LUCI OUT DDLUCI.LUCI%TYPE,
                                        G5ZDSJ OUT DDLUCI.ZDSJ%TYPE,
                                        G5DGSJ OUT DDLUCI.DGSJ%TYPE,
                                        G5DKSJ OUT DDLUCI.DKSJ%TYPE,
                                        G5TZSJ OUT DDLUCI.TZSJ%TYPE,
                                        G5WDSJ OUT DDLUCI.WDSJ%TYPE,
                                        G5QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G5FELIANG OUT DDLUCI.FELIANG%TYPE,
                                        G5FEC OUT DDLUCI.FEC%TYPE,
                                        G5FESI OUT DDLUCI.FESI%TYPE,
                                        G5FEMN OUT DDLUCI.FEMN%TYPE,
                                        G5FEP OUT DDLUCI.FEP%TYPE,
                                        G5FES OUT DDLUCI.FES%TYPE,
                                        G5FETI OUT DDLUCI.FETI%TYPE,
                                        G5ZHASIO2 OUT DDLUCI.ZHASIO2%TYPE,
                                        G5ZHACAO OUT DDLUCI.ZHACAO%TYPE,
                                        G5ZHAMGO OUT DDLUCI.ZHAMGO%TYPE,
                                        G5ZHAAL2O3 OUT DDLUCI.ZHAAL2O3%TYPE,
                                        G5ZHAS OUT DDLUCI.ZHAS%TYPE,
                                        G5ZHATIO2 OUT DDLUCI.ZHATIO2%TYPE,
                                        G5ZHAR2 OUT DDLUCI.ZHAR2%TYPE,
                                         G6LUCI OUT DDLUCI.LUCI%TYPE,
                                        G6ZDSJ OUT DDLUCI.ZDSJ%TYPE,
                                        G6DGSJ OUT DDLUCI.DGSJ%TYPE,
                                        G6DKSJ OUT DDLUCI.DKSJ%TYPE,
                                        G6TZSJ OUT DDLUCI.TZSJ%TYPE,
                                        G6WDSJ OUT DDLUCI.WDSJ%TYPE,
                                        G6QUCHU OUT DDLUCI.QUCHU%TYPE,
                                        G6FELIANG OUT DDLUCI.FELIANG%TYPE,
                                        G6FEC OUT DDLUCI.FEC%TYPE,
                                        G6FESI OUT DDLUCI.FESI%TYPE,
                                        G6FEMN OUT DDLUCI.FEMN%TYPE,
                                        G6FEP OUT DDLUCI.FEP%TYPE,
                                        G6FES OUT DDLUCI.FES%TYPE,
                                        G6FETI OUT DDLUCI.FETI%TYPE,
                                        G6ZHASIO2 OUT DDLUCI.ZHASIO2%TYPE,
                                        G6ZHACAO OUT DDLUCI.ZHACAO%TYPE,
                                        G6ZHAMGO OUT DDLUCI.ZHAMGO%TYPE,
                                        G6ZHAAL2O3 OUT DDLUCI.ZHAAL2O3%TYPE,
                                        G6ZHAS OUT DDLUCI.ZHAS%TYPE,
                                        G6ZHATIO2 OUT DDLUCI.ZHATIO2%TYPE,
                                        G6ZHAR2 OUT DDLUCI.ZHAR2%TYPE
                                        
)
IS
vZDSJ DATE;
vRQ DATE;
vBANCI VARCHAR(20);
vBANLUCI number;
CURSOR sluci2(pRQ DATE,pBANCI VARCHAR,pBANLUCI NUMBER,gl NUMBER DEFAULT 1) IS SELECT LUCI,DGSJ,DKSJ,TZSJ,WDSJ,QUCHU,FELIANG,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHAR2,ZHATIO2
             FROM DDLUCI WHERE ZDSJ=pRQ AND BANCI=pBANCI AND BANLUCI=pBANLUCI AND GAOLU=gl;
BEGIN
IF RQ>=TO_DATE('2009-05-15','YYYY-MM-DD') THEN
 vRQ:=RQ;
 vBANCI:=BANCI;
 vBANLUCI:=BANLUCI;
 IF (BANCI='Ò¹°à') AND BANLUCI=6 THEN
  vBANCI:='°×°à';
  vBANLUCI:=1;
 END IF;
 IF (BANCI='ÖÐ°à') AND BANLUCI=6 THEN
  vBANCI:='Ò¹°à';
  vBANLUCI:=1;
  vRQ:=RQ+NUMTODSINTERVAL(1,'DAY');
 END IF;
 IF (vBANCI='Ò¹°à') THEN
    vZDSJ:=vRQ+NUMTODSINTERVAL(vBANLUCI*90-40,'MINUTE');
 ELSIF (vBANCI='°×°à') THEN
    vZDSJ:=vRQ+NUMTODSINTERVAL((vBANLUCI+5)*90-40,'MINUTE');
 ELSIF (vBANCI='ÖÐ°à') THEN
    vZDSJ:=vRQ+NUMTODSINTERVAL((vBANLUCI+11)*90-40,'MINUTE');
 END IF;
 G1ZDSJ:=vZDSJ;
 G2ZDSJ:=vZDSJ;
 G3ZDSJ:=vZDSJ;
 IF (vBANCI='ÖÐ°à' and vBANLUCI=5) THEN
 G4ZDSJ:=vZDSJ+NUMTODSINTERVAL(39,'MINUTE');
 G5ZDSJ:=vZDSJ+NUMTODSINTERVAL(39,'MINUTE');
 G6ZDSJ:=vZDSJ+NUMTODSINTERVAL(39,'MINUTE');
 ELSE
 G4ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
 G5ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
 G6ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
 END IF;
ELSE
IF (BANCI='Ò¹°à') THEN
    vZDSJ:=RQ+NUMTODSINTERVAL(BANLUCI*80-50,'MINUTE');
ELSIF (BANCI='°×°à') THEN
    vZDSJ:=RQ+NUMTODSINTERVAL((BANLUCI+6)*80-50,'MINUTE');
ELSIF (BANCI='ÖÐ°à') THEN
    vZDSJ:=RQ+NUMTODSINTERVAL((BANLUCI+12)*80-50,'MINUTE');
END IF;
G1ZDSJ:=vZDSJ;
G2ZDSJ:=vZDSJ;
IF vZDSJ>TO_DATE('2006-08-12 13','YYYY-MM-DD HH24') THEN
G3ZDSJ:=vZDSJ;
ELSE
G3ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
END IF;
G4ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
G5ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
G6ZDSJ:=vZDSJ+NUMTODSINTERVAL(40,'MINUTE');
vRQ:=RQ;
vBANCI:=BANCI;
vBANLUCI:=BANLUCI;
END IF;
 IF vRQ>=TO_DATE('2009-06-18','YYYY-MM-DD') THEN
   IF (vBANCI='°×°à' and vBANLUCI>1) Or (vBANCI='ÖÐ°à' and vBANLUCI<5)  THEN
     G1ZDSJ:=G1ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
     G2ZDSJ:=G2ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
     G3ZDSJ:=G3ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
     G4ZDSJ:=G4ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
     G5ZDSJ:=G5ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
     G6ZDSJ:=G6ZDSJ+NUMTODSINTERVAL(10,'MINUTE');
   END IF;
 END IF;
 
OPEN sluci2(G1ZDSJ,vBANCI,vBANLUCI,1);
FETCH sluci2 INTO G1LUCI,G1DGSJ, G1DKSJ, G1TZSJ, G1WDSJ, G1QUCHU, G1FELIANG, G1FEC, G1FESI,
      G1FEMN, G1FEP, G1FES, G1FETI, G1ZHASIO2, G1ZHACAO, G1ZHAMGO, G1ZHAAL2O3, G1ZHAS, G1ZHAR2,G1ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=1;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (1, G1ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

OPEN sluci2(G2ZDSJ,BANCI,BANLUCI,2);
FETCH sluci2 INTO G2LUCI,G2DGSJ, G2DKSJ, G2TZSJ, G2WDSJ, G2QUCHU, G2FELIANG, G2FEC, G2FESI,
      G2FEMN, G2FEP, G2FES, G2FETI, G2ZHASIO2, G2ZHACAO, G2ZHAMGO, G2ZHAAL2O3, G2ZHAS, G2ZHAR2,G2ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=2;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (2, G2ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

OPEN sluci2(G3ZDSJ,BANCI,BANLUCI,3);
FETCH sluci2 INTO G3LUCI,G3DGSJ, G3DKSJ, G3TZSJ, G3WDSJ, G3QUCHU, G3FELIANG, G3FEC, G3FESI,
      G3FEMN, G3FEP, G3FES, G3FETI, G3ZHASIO2, G3ZHACAO, G3ZHAMGO, G3ZHAAL2O3, G3ZHAS, G3ZHAR2,G3ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=3;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (3, G3ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

OPEN sluci2(G4ZDSJ,BANCI,BANLUCI,4);
FETCH sluci2 INTO G4LUCI,G4DGSJ, G4DKSJ, G4TZSJ, G4WDSJ, G4QUCHU, G4FELIANG, G4FEC, G4FESI,
      G4FEMN, G4FEP, G4FES, G4FETI, G4ZHASIO2, G4ZHACAO, G4ZHAMGO, G4ZHAAL2O3, G4ZHAS, G4ZHAR2,G4ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=4;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (4, G4ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

OPEN sluci2(G5ZDSJ,BANCI,BANLUCI,5);
FETCH sluci2 INTO G5LUCI,G5DGSJ, G5DKSJ, G5TZSJ, G5WDSJ, G5QUCHU, G5FELIANG, G5FEC, G5FESI,
      G5FEMN, G5FEP, G5FES, G5FETI, G5ZHASIO2, G5ZHACAO, G5ZHAMGO, G5ZHAAL2O3, G5ZHAS, G5ZHAR2,G5ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=5;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (5, G5ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

OPEN sluci2(G6ZDSJ,BANCI,BANLUCI,6);
FETCH sluci2 INTO G6LUCI,G6DGSJ, G6DKSJ, G6TZSJ, G6WDSJ, G6QUCHU, G6FELIANG, G6FEC, G6FESI,
      G6FEMN, G6FEP, G6FES, G6FETI, G6ZHASIO2, G6ZHACAO, G6ZHAMGO, G6ZHAAL2O3, G6ZHAS, G6ZHAR2,G6ZHATIO2;
IF sluci2%NOTFOUND THEN
--SELECT COUNT(*) INTO vCount FROM DDLUCI WHERE TRUNC(ZDSJ)=TRUNC(vRQ) AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=5;
--IF vCount=0 THEN
INSERT INTO DDLUCI(GAOLU, ZDSJ,BANCI,BANLUCI) VALUES (6, G6ZDSJ,vBANCI,vBANLUCI);
--END IF;
END IF;
CLOSE sluci2;

END SELECTLUCI2;
/
