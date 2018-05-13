declare
t Date;
begin
t:=to_date('2006-07-01','YYYY-MM-DD');
while t<to_date('2008-07-01','YYYY-MM-DD') loop
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t,'Ò¹°à','±û°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t,'°×°à','¶¡°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+1,'Ò¹°à','¼×°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+1,'°×°à','ÒÒ°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+2,'Ò¹°à','¶¡°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+2,'°×°à','±û°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+3,'Ò¹°à','ÒÒ°à');
insert into ¸ßÂ¯´ó°àÅÅ°à±í values(t+3,'°×°à','¼×°à');
t:=t+4;
end loop;
commit;
end;