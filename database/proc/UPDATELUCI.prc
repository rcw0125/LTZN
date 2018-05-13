CREATE OR REPLACE PROCEDURE "UPDATELUCI2" (RQ IN DATE,
                                        BANCI IN VARCHAR,
                                        BANLUCI IN NUMBER,
                                        G1DGSJ IN  DDLUCI.DGSJ%TYPE,
                                        G1DKSJ IN   DDLUCI.DKSJ%TYPE,
                                        G1TZSJ IN   DDLUCI.TZSJ%TYPE,
                                        G1QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G1FELIANG IN  DDLUCI.FELIANG%TYPE,
                                        G1FEC IN  DDLUCI.FEC%TYPE,
                                        G1FESI IN   DDLUCI.FESI%TYPE,
                                        G1FEMN IN   DDLUCI.FEMN%TYPE,
                                        G1FEP IN   DDLUCI.FEP%TYPE,
                                        G1FES IN   DDLUCI.FES%TYPE,
                                        G1FETI IN   DDLUCI.FETI%TYPE,
                                        G1ZHASIO2 IN   DDLUCI.ZHASIO2%TYPE,
                                        G1ZHACAO IN   DDLUCI.ZHACAO%TYPE,
                                        G1ZHAMGO IN   DDLUCI.ZHAMGO%TYPE,
                                        G1ZHAAL2O3 IN   DDLUCI.ZHAAL2O3%TYPE,
                                        G1ZHAS IN   DDLUCI.ZHAS%TYPE,
                                        G1ZHATIO2 IN   DDLUCI.ZHATIO2%TYPE,
                                        G2DGSJ IN   DDLUCI.DGSJ%TYPE,
                                        G2DKSJ IN   DDLUCI.DKSJ%TYPE,
                                        G2TZSJ IN   DDLUCI.TZSJ%TYPE,
                                        G2QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G2FELIANG IN   DDLUCI.FELIANG%TYPE,
                                        G2FEC IN   DDLUCI.FEC%TYPE,
                                        G2FESI IN   DDLUCI.FESI%TYPE,
                                        G2FEMN IN   DDLUCI.FEMN%TYPE,
                                        G2FEP IN  DDLUCI.FEP%TYPE,
                                        G2FES IN  DDLUCI.FES%TYPE,
                                        G2FETI IN  DDLUCI.FETI%TYPE,
                                        G2ZHASIO2 IN  DDLUCI.ZHASIO2%TYPE,
                                        G2ZHACAO IN  DDLUCI.ZHACAO%TYPE,
                                        G2ZHAMGO IN  DDLUCI.ZHAMGO%TYPE,
                                        G2ZHAAL2O3 IN  DDLUCI.ZHAAL2O3%TYPE,
                                        G2ZHAS IN  DDLUCI.ZHAS%TYPE,
                                        G2ZHATIO2 IN  DDLUCI.ZHATIO2%TYPE,
                                        G3DGSJ IN   DDLUCI.DGSJ%TYPE,
                                        G3DKSJ IN   DDLUCI.DKSJ%TYPE,
                                        G3TZSJ IN   DDLUCI.TZSJ%TYPE,
                                        G3QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G3FELIANG IN  DDLUCI.FELIANG%TYPE,
                                        G3FEC IN  DDLUCI.FEC%TYPE,
                                        G3FESI IN  DDLUCI.FESI%TYPE,
                                        G3FEMN IN  DDLUCI.FEMN%TYPE,
                                        G3FEP IN  DDLUCI.FEP%TYPE,
                                        G3FES IN  DDLUCI.FES%TYPE,
                                        G3FETI IN  DDLUCI.FETI%TYPE,
                                        G3ZHASIO2 IN  DDLUCI.ZHASIO2%TYPE,
                                        G3ZHACAO IN  DDLUCI.ZHACAO%TYPE,
                                        G3ZHAMGO IN  DDLUCI.ZHAMGO%TYPE,
                                        G3ZHAAL2O3 IN  DDLUCI.ZHAAL2O3%TYPE,
                                        G3ZHAS IN  DDLUCI.ZHAS%TYPE,
                                        G3ZHATIO2 IN  DDLUCI.ZHATIO2%TYPE,
                                        G4DGSJ IN  DDLUCI.DGSJ%TYPE,
                                        G4DKSJ IN  DDLUCI.DKSJ%TYPE,
                                        G4TZSJ IN  DDLUCI.TZSJ%TYPE,
                                        G4QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G4FELIANG IN  DDLUCI.FELIANG%TYPE,
                                        G4FEC IN  DDLUCI.FEC%TYPE,
                                        G4FESI IN  DDLUCI.FESI%TYPE,
                                        G4FEMN IN  DDLUCI.FEMN%TYPE,
                                        G4FEP IN  DDLUCI.FEP%TYPE,
                                        G4FES IN  DDLUCI.FES%TYPE,
                                        G4FETI IN  DDLUCI.FETI%TYPE,
                                        G4ZHASIO2 IN  DDLUCI.ZHASIO2%TYPE,
                                        G4ZHACAO IN  DDLUCI.ZHACAO%TYPE,
                                        G4ZHAMGO IN  DDLUCI.ZHAMGO%TYPE,
                                        G4ZHAAL2O3 IN  DDLUCI.ZHAAL2O3%TYPE,
                                        G4ZHAS IN  DDLUCI.ZHAS%TYPE,
                                        G4ZHATIO2 IN  DDLUCI.ZHATIO2%TYPE,
                                        G5DGSJ IN  DDLUCI.DGSJ%TYPE,
                                        G5DKSJ IN  DDLUCI.DKSJ%TYPE,
                                        G5TZSJ IN  DDLUCI.TZSJ%TYPE,
                                        G5QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G5FELIANG IN  DDLUCI.FELIANG%TYPE,
                                        G5FEC IN  DDLUCI.FEC%TYPE,
                                        G5FESI IN  DDLUCI.FESI%TYPE,
                                        G5FEMN IN  DDLUCI.FEMN%TYPE,
                                        G5FEP IN  DDLUCI.FEP%TYPE,
                                        G5FES IN  DDLUCI.FES%TYPE,
                                        G5FETI IN  DDLUCI.FETI%TYPE,
                                        G5ZHASIO2 IN  DDLUCI.ZHASIO2%TYPE,
                                        G5ZHACAO IN  DDLUCI.ZHACAO%TYPE,
                                        G5ZHAMGO IN  DDLUCI.ZHAMGO%TYPE,
                                        G5ZHAAL2O3 IN  DDLUCI.ZHAAL2O3%TYPE,
                                        G5ZHAS IN  DDLUCI.ZHAS%TYPE,
                                        G5ZHATIO2 IN  DDLUCI.ZHATIO2%TYPE,
                                        G6DGSJ IN  DDLUCI.DGSJ%TYPE,
                                        G6DKSJ IN  DDLUCI.DKSJ%TYPE,
                                        G6TZSJ IN  DDLUCI.TZSJ%TYPE,
                                        G6QUCHU IN  DDLUCI.QUCHU%TYPE,
                                        G6FELIANG IN  DDLUCI.FELIANG%TYPE,
                                        G6FEC IN  DDLUCI.FEC%TYPE,
                                        G6FESI IN  DDLUCI.FESI%TYPE,
                                        G6FEMN IN  DDLUCI.FEMN%TYPE,
                                        G6FEP IN  DDLUCI.FEP%TYPE,
                                        G6FES IN  DDLUCI.FES%TYPE,
                                        G6FETI IN  DDLUCI.FETI%TYPE,
                                        G6ZHASIO2 IN  DDLUCI.ZHASIO2%TYPE,
                                        G6ZHACAO IN  DDLUCI.ZHACAO%TYPE,
                                        G6ZHAMGO IN  DDLUCI.ZHAMGO%TYPE,
                                        G6ZHAAL2O3 IN  DDLUCI.ZHAAL2O3%TYPE,
                                        G6ZHAS IN  DDLUCI.ZHAS%TYPE,
                                        G6ZHATIO2 IN  DDLUCI.ZHATIO2%TYPE,
                                        G1WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G2WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G3WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G4WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G5WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G6WDSJ IN DDLUCI.WDSJ%TYPE,
                                        G1ZHAR2 IN DDLUCI.WDSJ%TYPE,
                                        G2ZHAR2 IN DDLUCI.WDSJ%TYPE,
                                        G3ZHAR2 IN DDLUCI.WDSJ%TYPE,
                                        G4ZHAR2 IN DDLUCI.WDSJ%TYPE,
                                        G5ZHAR2 IN DDLUCI.WDSJ%TYPE,
                                        G6ZHAR2 IN DDLUCI.WDSJ%TYPE
)
IS
vZDSJ1 DATE;
vZDSJ2 DATE;
vZDSJ3 DATE;
vZDSJ4 DATE;
vZDSJ5 DATE;
vZDSJ6 DATE;
lp number;
pDKSJ DATE;
vBANCI VARCHAR(20);
vBANLUCI number;
maxLUCI1 NUMBER;
maxLUCI2 NUMBER;
maxLUCI3 NUMBER;
maxLUCI4 NUMBER;
maxLUCI5 NUMBER;
maxLUCI6 NUMBER;
maxZDSJ1 DATE;
maxZDSJ2 DATE;
maxZDSJ3 DATE;
maxZDSJ4 DATE;
maxZDSJ5 DATE;
maxZDSJ6 DATE;
maxRQ   DATE;
maxBANCI VARCHAR(20);
maxBANLUCI NUMBER;
old_DKSJ1 DATE;
old_DKSJ2 DATE;
old_DKSJ3 DATE;
old_DKSJ4 DATE;
old_DKSJ5 DATE;
old_DKSJ6 DATE;
err varchar2(2000);
BEGIN
IF RQ>=TO_DATE('2009-05-15','YYYY-MM-DD') THEN
 IF (BANCI='夜班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL(BANLUCI*90-40,'MINUTE');
 ELSIF (BANCI='白班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL((BANLUCI+5)*90-40,'MINUTE');
 ELSIF (BANCI='中班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL((BANLUCI+11)*90-40,'MINUTE');
 END IF;
 vZDSJ2:=vZDSJ1;
 vZDSJ3:=vZDSJ1;
 IF (BANCI='中班' and BANLUCI=5) THEN
 vZDSJ4:=vZDSJ1+NUMTODSINTERVAL(39,'MINUTE');
 vZDSJ5:=vZDSJ1+NUMTODSINTERVAL(39,'MINUTE');
 ELSE
  vZDSJ4:=vZDSJ1+NUMTODSINTERVAL(40,'MINUTE');
  vZDSJ5:=vZDSJ1+NUMTODSINTERVAL(40,'MINUTE');
 END IF;
 vBANCI:=BANCI;
 vBANLUCI:=BANLUCI;
 IF (BANCI='夜班') AND BANLUCI=6 THEN
  vBANCI:='白班';
  vBANLUCI:=1;
 END IF;
 IF (BANCI='中班') AND BANLUCI=6 THEN
  vBANCI:='夜班';
  vBANLUCI:=1;
 END IF;
ELSE
IF (BANCI='夜班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL(BANLUCI*80-50,'MINUTE');
END IF;
IF (BANCI='白班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL((BANLUCI+6)*80-50,'MINUTE');
END IF;
IF (BANCI='中班') THEN
    vZDSJ1:=RQ+NUMTODSINTERVAL((BANLUCI+12)*80-50,'MINUTE');
END IF;
vZDSJ2:=vZDSJ1;
IF vZDSJ1>TO_DATE('2006-08-12 13','YYYY-MM-DD HH24') THEN
vZDSJ3:=vZDSJ1;
ELSE
vZDSJ3:=vZDSJ1+NUMTODSINTERVAL(40,'MINUTE');
END IF;
vZDSJ4:=vZDSJ1+NUMTODSINTERVAL(40,'MINUTE');
vZDSJ5:=vZDSJ1+NUMTODSINTERVAL(40,'MINUTE');
vBANCI:=BANCI;
vBANLUCI:=BANLUCI;
END IF;
 IF RQ>=TO_DATE('2009-06-18','YYYY-MM-DD') THEN
   IF (vBANCI='白班' and vBANLUCI>1) Or (vBANCI='中班' and vBANLUCI<5)  THEN
     vZDSJ1:=vZDSJ1+NUMTODSINTERVAL(10,'MINUTE');
     vZDSJ2:=vZDSJ2+NUMTODSINTERVAL(10,'MINUTE');
     vZDSJ3:=vZDSJ3+NUMTODSINTERVAL(10,'MINUTE');
     vZDSJ4:=vZDSJ4+NUMTODSINTERVAL(10,'MINUTE');
     vZDSJ5:=vZDSJ5+NUMTODSINTERVAL(10,'MINUTE');
   END IF;
 END IF;
 
 vZDSJ6:=vZDSJ5;
 
--1高炉
--计算料批数
lp:=null;
IF (G1DKSJ is not null) THEN
 BEGIN
   select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ1 and gaolu=1 and dksj is not null;
   exception
    when no_data_found then
      lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
    select count(distinct pishu) into lp from lt_liao where gaolu=1 and t>pDKSJ and t<=G1DKSJ;
    exception
    when no_data_found then
      lp:=null;
  END;
 END IF;
END IF;
--更新数据
BEGIN
SELECT DKSJ into old_DKSJ1 FROM DDLUCI WHERE (ZDSJ=vZDSJ1 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=1);
  IF SQL%FOUND THEN
    UPDATE DDLUCI
    SET   DGSJ =G1DGSJ, DKSJ =G1DKSJ,TZSJ =G1TZSJ, QUCHU =G1QUCHU, FELIANG =G1FELIANG,
          FEC =G1FEC,FESI =G1FESI, FEMN =G1FEMN, FES =G1FES, FEP =G1FEP, FETI =G1FETI,
          ZHASIO2 =G1ZHASIO2, ZHACAO =G1ZHACAO, ZHAMGO =G1ZHAMGO,ZHAAL2O3 =G1ZHAAL2O3, ZHAS =G1ZHAS,ZHATIO2=G1ZHATIO2,
          WDSJ=G1WDSJ,ZHAR2=G1ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ1,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
    WHERE (ZDSJ=vZDSJ1 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=1);
   END IF;
    EXCEPTION
   WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
            (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ, QUCHU, FELIANG,
             FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
              ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
      VALUES (1,vZDSJ1,G1DGSJ, G1DKSJ, G1TZSJ, G1QUCHU, G1FELIANG,
              G1FEC, G1FESI,G1FEMN, G1FEP, G1FES, G1FETI,GETR3(G1ZHACAO,G1ZHAMGO,G1ZHASIO2),GETR4(G1ZHACAO,G1ZHAMGO,G1ZHASIO2,G1ZHAAL2O3),GETMGOALO(G1ZHAMGO,G1ZHAAL2O3),
              G1ZHASIO2, G1ZHACAO, G1ZHAMGO, G1ZHAAL2O3, G1ZHAS,G1ZHATIO2,BANCI,BANLUCI,G1WDSJ,G1ZHAR2,lp);
    END;
END;
IF (G1DKSJ is null and old_DKSJ1 is not null) or (G1DKSJ is not null and old_DKSJ1 is null) THEN
 updatelucihao(1,vZDSJ1);
END IF;

--2高炉
--计算料批数
lp:=null;
IF (G2DKSJ is not null) THEN
 BEGIN
   select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ2 and gaolu=2 and dksj is not null;
   exception
    when no_data_found then
      lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
    select count(distinct pishu) into lp from lt_liao where gaolu=2 and t>pDKSJ and t<=G2DKSJ;
    exception
    when no_data_found then
      lp:=null;
  END;
 END IF;
END IF;
--更新数据
BEGIN
SELECT DKSJ into old_DKSJ2 FROM DDLUCI WHERE (ZDSJ=vZDSJ2 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=2);
  IF SQL%FOUND THEN
      UPDATE DDLUCI
      SET    DGSJ =G2DGSJ, DKSJ =G2DKSJ,TZSJ =G2TZSJ, QUCHU =G2QUCHU, FELIANG =G2FELIANG,
              FEC =G2FEC, FESI =G2FESI, FEMN =G2FEMN, FES =G2FES,FEP =G2FEP, FETI =G2FETI,
              ZHASIO2 =G2ZHASIO2,ZHACAO =G2ZHACAO, ZHAMGO =G2ZHAMGO, ZHAAL2O3 =G2ZHAAL2O3,ZHAS =G2ZHAS,ZHATIO2=G2ZHATIO2,
              WDSJ=G2WDSJ,ZHAR2=G2ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ2,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
      WHERE (ZDSJ=vZDSJ2 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=2);
  END IF;
  EXCEPTION
   WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
            (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ,QUCHU, FELIANG,
             FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
             ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
      VALUES (2, vZDSJ2, G2DGSJ, G2DKSJ, G2TZSJ, G2QUCHU, G2FELIANG,
             G2FEC, G2FESI,G2FEMN, G2FEP, G2FES, G2FETI,GETR3(G2ZHACAO,G2ZHAMGO,G2ZHASIO2),GETR4(G2ZHACAO,G2ZHAMGO,G2ZHASIO2,G2ZHAAL2O3),GETMGOALO(G2ZHAMGO,G2ZHAAL2O3),
             G2ZHASIO2, G2ZHACAO, G2ZHAMGO, G2ZHAAL2O3, G2ZHAS,G2ZHATIO2,BANCI,BANLUCI,G2WDSJ,G2ZHAR2,lp);
    END;
END;
IF (G2DKSJ is null and old_DKSJ2 is not null) or (G2DKSJ is not null and old_DKSJ2 is null) THEN
 updatelucihao(2,vZDSJ2);
END IF;

--3高炉
--计算料批数
lp:=null;
IF (G3DKSJ is not null) THEN
 BEGIN
   select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ3 and gaolu=3 and dksj is not null;
   exception
    when no_data_found then
      lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
    select count(distinct pishu) into lp from lt_liao where gaolu=3 and t>pDKSJ and t<=G3DKSJ;
    exception
    when no_data_found then
      lp:=null;
  END;
 END IF;
END IF;
--更新数据
BEGIN
  SELECT DKSJ into old_DKSJ3 FROM DDLUCI WHERE (ZDSJ=vZDSJ3 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=3);
  IF SQL%FOUND THEN
      UPDATE DDLUCI
      SET    DGSJ =G3DGSJ, DKSJ =G3DKSJ,TZSJ =G3TZSJ, QUCHU =G3QUCHU, FELIANG =G3FELIANG,
             FEC =G3FEC, FESI =G3FESI, FEMN =G3FEMN, FES =G3FES,FEP =G3FEP, FETI =G3FETI,
             ZHASIO2 =G3ZHASIO2,ZHACAO =G3ZHACAO, ZHAMGO =G3ZHAMGO,ZHAAL2O3 =G3ZHAAL2O3,ZHAS =G3ZHAS,ZHATIO2=G3ZHATIO2,
             WDSJ=G3WDSJ,ZHAR2=G3ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ3,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
      WHERE (ZDSJ=vZDSJ3 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=3);
  END IF;
  EXCEPTION
   WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
            (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ,QUCHU,FELIANG,
             FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
             ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
      VALUES (3, vZDSJ3, G3DGSJ, G3DKSJ, G3TZSJ,G3QUCHU, G3FELIANG,
              G3FEC, G3FESI,G3FEMN, G3FEP, G3FES, G3FETI,GETR3(G3ZHACAO,G3ZHAMGO,G3ZHASIO2),GETR4(G3ZHACAO,G3ZHAMGO,G3ZHASIO2,G3ZHAAL2O3),GETMGOALO(G3ZHAMGO,G3ZHAAL2O3),
              G3ZHASIO2, G3ZHACAO, G3ZHAMGO, G3ZHAAL2O3, G3ZHAS,G3ZHATIO2,BANCI,BANLUCI,G3WDSJ,G3ZHAR2,lp);
    END;
END;

IF (G3DKSJ is null and old_DKSJ3 is not null) or (G3DKSJ is not null and old_DKSJ3 is null) THEN
 updatelucihao(3,vZDSJ3);
END IF;

--4高炉
--计算料批数
lp:=null;
IF (G4DKSJ is not null) THEN
 BEGIN
   select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ4 and gaolu=4 and dksj is not null;
   exception
    when no_data_found then
      lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
    select count(distinct pishu) into lp from lt_liao where gaolu=4 and t>pDKSJ and t<=G4DKSJ;
    exception
    when no_data_found then
      lp:=null;
  END;
 END IF;
END IF;
--更新数据
BEGIN
SELECT DKSJ into old_DKSJ4 FROM DDLUCI WHERE (ZDSJ=vZDSJ4 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=4);
  IF SQL%FOUND THEN
      UPDATE DDLUCI
      SET    DGSJ =G4DGSJ,DKSJ =G4DKSJ, TZSJ =G4TZSJ, QUCHU =G4QUCHU,FELIANG =G4FELIANG,
             FEC =G4FEC, FESI =G4FESI,FEMN =G4FEMN, FES =G4FES, FEP =G4FEP,FETI =G4FETI,
             ZHASIO2 =G4ZHASIO2,ZHACAO =G4ZHACAO, ZHAMGO =G4ZHAMGO,ZHAAL2O3 =G4ZHAAL2O3, ZHAS =G4ZHAS,ZHATIO2=G4ZHATIO2,
             WDSJ=G4WDSJ,ZHAR2=G4ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ4,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
      WHERE (ZDSJ=vZDSJ4 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=4);
  END IF;
  EXCEPTION
   WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
            (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ,QUCHU, FELIANG,
             FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
             ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
      VALUES ( 4, vZDSJ4, G4DGSJ, G4DKSJ, G4TZSJ,G4QUCHU, G4FELIANG,
              G4FEC, G4FESI,G4FEMN, G4FEP, G4FES, G4FETI,GETR3(G4ZHACAO,G4ZHAMGO,G4ZHASIO2),GETR4(G4ZHACAO,G4ZHAMGO,G4ZHASIO2,G4ZHAAL2O3),GETMGOALO(G4ZHAMGO,G4ZHAAL2O3),
              G4ZHASIO2, G4ZHACAO, G4ZHAMGO, G4ZHAAL2O3, G4ZHAS,G4ZHATIO2,BANCI,BANLUCI,G4WDSJ,G4ZHAR2,lp);
    END;
END;
IF (G4DKSJ is null and old_DKSJ4 is not null) or (G4DKSJ is not null and old_DKSJ4 is null) THEN
 updatelucihao(4,vZDSJ4);
END IF;

--5高炉
--计算料批数
lp:=null;
IF (G5DKSJ is not null) THEN
 BEGIN
     select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ5 and gaolu=5 and dksj is not null;
     exception
      when no_data_found then
        lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
      select count(distinct pishu) into lp from lt_liao where gaolu=5 and t>pDKSJ and t<=G5DKSJ;
      exception
      when no_data_found then
        lp:=null;
  END;
 END IF;
END IF;
--更新数据
BEGIN
  SELECT DKSJ into old_DKSJ5 FROM DDLUCI WHERE (ZDSJ=vZDSJ5 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=5);
  IF SQL%FOUND THEN
      UPDATE DDLUCI
      SET    DGSJ =G5DGSJ,DKSJ =G5DKSJ, TZSJ =G5TZSJ,QUCHU =G5QUCHU, FELIANG =G5FELIANG,
           FEC =G5FEC, FESI =G5FESI, FEMN =G5FEMN,FES =G5FES, FEP =G5FEP, FETI =G5FETI,
           ZHASIO2 =G5ZHASIO2, ZHACAO =G5ZHACAO,ZHAMGO =G5ZHAMGO, ZHAAL2O3 =G5ZHAAL2O3,ZHAS =G5ZHAS,ZHATIO2=G5ZHATIO2,
           WDSJ=G5WDSJ,ZHAR2=G5ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ5,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
      WHERE (ZDSJ=vZDSJ5 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=5);
  END IF;
  EXCEPTION
   WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
            (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ,QUCHU, FELIANG,
             FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
             ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
      VALUES (5, vZDSJ5, G5DGSJ, G5DKSJ, G5TZSJ,G5QUCHU, G5FELIANG,
              G5FEC, G5FESI,G5FEMN, G5FEP, G5FES, G5FETI,GETR3(G5ZHACAO,G5ZHAMGO,G5ZHASIO2),GETR4(G5ZHACAO,G5ZHAMGO,G5ZHASIO2,G5ZHAAL2O3),GETMGOALO(G5ZHAMGO,G5ZHAAL2O3),
              G5ZHASIO2, G5ZHACAO, G5ZHAMGO, G5ZHAAL2O3, G5ZHAS,G5ZHATIO2,BANCI,BANLUCI,G5WDSJ,G5ZHAR2,lp);
    END;
END;

IF (G5DKSJ is null and old_DKSJ5 is not null) or (G5DKSJ is not null and old_DKSJ5 is null) THEN
 updatelucihao(5,vZDSJ5);
END IF;

--6高炉
--计算料批数
lp:=null;
IF (G6DKSJ is not null) THEN
 BEGIN
   select max(dksj) into pDKSJ from ddluci where zdsj<vZDSJ6 and gaolu=6 and dksj is not null;
   exception
    when no_data_found then
      lp:=null;
  END;
 IF (pDKSJ is not null) THEN
  BEGIN
    select count(distinct pishu) into lp from lt_liao where gaolu=6 and t>pDKSJ and t<=G6DKSJ;
    exception
    when no_data_found then
      lp:=null;
  END;
 END IF;
END IF;

--更新数据
BEGIN
 SELECT DKSJ into old_DKSJ6 FROM DDLUCI WHERE (ZDSJ=vZDSJ6 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=6);
 IF SQL%FOUND THEN
  UPDATE DDLUCI
  SET  DGSJ =G6DGSJ,DKSJ =G6DKSJ, TZSJ =G6TZSJ,QUCHU =G6QUCHU, FELIANG =G6FELIANG,
       FEC =G6FEC, FESI =G6FESI, FEMN =G6FEMN,FES =G6FES, FEP =G6FEP, FETI =G6FETI,
       ZHASIO2 =G6ZHASIO2, ZHACAO =G6ZHACAO,ZHAMGO =G6ZHAMGO, ZHAAL2O3 =G6ZHAAL2O3,ZHAS =G6ZHAS,ZHATIO2=G6ZHATIO2,
       WDSJ=G6WDSJ,ZHAR2=G6ZHAR2,LIAOPI=lp,ZDSJ=vZDSJ6,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)
 WHERE (ZDSJ=vZDSJ6 AND BANCI=vBANCI AND BANLUCI=vBANLUCI AND GAOLU=6);
 END IF;
 EXCEPTION
    WHEN NO_DATA_FOUND THEN
    BEGIN
      INSERT INTO DDLUCI
       (GAOLU, ZDSJ, DGSJ, DKSJ, TZSJ,QUCHU, FELIANG,
       FEC, FESI, FEMN, FEP,FES, FETI,ZHAR3,ZHAR4,ZHAMGOALO,
       ZHASIO2, ZHACAO, ZHAMGO, ZHAAL2O3, ZHAS,ZHATIO2,BANCI,BANLUCI,WDSJ,ZHAR2,LIAOPI)
        VALUES (6, vZDSJ6, G6DGSJ, G6DKSJ, G6TZSJ,G6QUCHU, G6FELIANG,
        G6FEC, G6FESI,G6FEMN, G6FEP, G6FES, G6FETI,GETR3(G6ZHACAO,G6ZHAMGO,G6ZHASIO2),GETR4(G6ZHACAO,G6ZHAMGO,G6ZHASIO2,G6ZHAAL2O3),GETMGOALO(G6ZHAMGO,G6ZHAAL2O3),
        G6ZHASIO2, G6ZHACAO, G6ZHAMGO, G6ZHAAL2O3, G6ZHAS,G6ZHATIO2,BANCI,BANLUCI,G6WDSJ,G6ZHAR2,lp);
    END;
END;
IF (G6DKSJ is null and old_DKSJ6 is not null) or (G6DKSJ is not null and old_DKSJ6 is null) THEN
 updatelucihao(6,vZDSJ6);
END IF;

/*IF vZDSJ5>TO_DATE('2008-11-09','YYYY-MM-DD') THEN
UPDATE DDLUCI a SET LUCI=TO_CHAR((SELECT COUNT(*)+150131449 FROM DDLUCI WHERE ZDSJ<a.ZDSJ AND ZDSJ>TO_DATE('2008-11-09','YYYY-MM-DD')  AND GAOLU=5 AND DKSJ IS NOT NULL)) WHERE DKSJ IS NOT NULL AND ZDSJ>=vZDSJ5 AND GAOLU=5;
END IF;*/

IF vZDSJ5>TO_DATE('2008-11-09','YYYY-MM-DD') THEN
IF (G1DKSJ is null and old_DKSJ1 is not null) or (G1DKSJ is not null and old_DKSJ1 is null) or (G2DKSJ is null and old_DKSJ2 is not null) or (G2DKSJ is not null and old_DKSJ2 is null) or  (G3DKSJ is null and old_DKSJ3 is not null) or (G3DKSJ is not null and old_DKSJ3 is null) or(G4DKSJ is null and old_DKSJ4 is not null) or (G4DKSJ is not null and old_DKSJ4 is null) or(G5DKSJ is null and old_DKSJ5 is not null) or (G5DKSJ is not null and old_DKSJ5 is null)  THEN
--UPDATE DDLUCI  SET LUCI=null WHERE DKSJ IS NULL AND ZDSJ>=vZDSJ1;
--在后面插入几个预计炉次
select trunc(zdsj) as rq,banci,banluci into maxRQ,maxBANCI,maxBANLUCI from (select zdsj,banci,banluci from ddluci where dksj is not null order by zdsj desc) where rownum=1;
select to_number(max(luci)) into maxLUCI1 from ddluci where dksj is not null and gaolu=1;
select to_number(max(luci)) into maxLUCI2 from ddluci where dksj is not null and gaolu=2;
select to_number(max(luci)) into maxLUCI3 from ddluci where dksj is not null and gaolu=3;
select to_number(max(luci)) into maxLUCI4 from ddluci where dksj is not null and gaolu=4;
select to_number(max(luci)) into maxLUCI5 from ddluci where dksj is not null and gaolu=5;
select to_number(max(luci)) into maxLUCI6 from ddluci where dksj is not null and gaolu=6;

IF RQ>=TO_DATE('2009-05-15','YYYY-MM-DD') THEN

if maxBanci='夜班' and maxBanluci>=5 then
maxBanci:='白班';
maxBanluci:=1;
elsif maxBanci='白班' and maxBanluci>=6 then
maxBanci:='中班';
maxBanluci:=1;
elsif maxBanci='中班' and maxBanluci>=5 then
maxBanci:='夜班';
maxRQ:=maxRQ+1;
maxBanluci:=1;
else
maxBanluci:=maxBanluci+1;
end if;
ELSE
if maxBanluci<6 then
maxBanluci:=maxBanluci+1;
else
if maxBanci='夜班' then
maxBanci:='白班';
elsif maxBanci='白班' then
maxBanci:='中班';
elsif maxBanci='中班' then
maxBanci:='夜班';
maxRQ:=maxRQ+1;
end if;
maxBanluci:=1;
end if;
end if;
maxLUCI1:=maxLUCI1+1;
maxLUCI2:=maxLUCI2+1;
maxLUCI3:=maxLUCI3+1;
maxLUCI4:=maxLUCI4+1;
maxLUCI5:=maxLUCI5+1;
maxLUCI6:=maxLUCI6+1;

maxZDSJ1 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,1);
update ddluci set luci='*'||to_char(maxluci1),zdsj=maxZDSJ1 where gaolu=1 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(1,maxZDSJ1,maxBanci,maxBanluci,'*'||to_char(maxluci1));
end if;

maxZDSJ2 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,2);
update ddluci set luci='*'||to_char(maxluci2),zdsj=maxZDSJ2 where gaolu=2 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(2,maxZDSJ2,maxBanci,maxBanluci,'*'||to_char(maxluci2));
end if;

maxZDSJ3 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,3);
update ddluci set luci='*'||to_char(maxluci3),zdsj=maxZDSJ3 where gaolu=3 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(3,maxZDSJ3,maxBanci,maxBanluci,'*'||to_char(maxluci3));
end if;

maxZDSJ4 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,4);
update ddluci set luci='*'||to_char(maxluci4),zdsj=maxZDSJ4 where gaolu=4 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(4,maxZDSJ4,maxBanci,maxBanluci,'*'||to_char(maxluci4));
end if;

maxZDSJ5 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,5);
update ddluci set luci='*'||to_char(maxluci5),zdsj=maxZDSJ5 where gaolu=5 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(5,maxZDSJ5,maxBanci,maxBanluci,'*'||to_char(maxluci5));
end if;

maxZDSJ6 :=GETZDSJ(maxRQ,maxBanci,maxBanluci,6);
update ddluci set luci='*'||to_char(maxluci6),zdsj=maxZDSJ6 where gaolu=6 and trunc(zdsj)=maxRQ and banci=maxbanci and banluci=maxbanluci;
if sql%notfound then
insert into ddluci(gaolu,zdsj,banci,banluci,luci) values(6,maxZDSJ6,maxBanci,maxBanluci,'*'||to_char(maxluci6));
end if;

END IF;
END IF;
COMMIT;
EXCEPTION
  WHEN OTHERS THEN
   begin
     err := substrb(sqlerrm,1,2000);
     INSERT INTO ERRLOG VALUES(sysdate,err);
   end;
END UPDATELUCI2;
/
