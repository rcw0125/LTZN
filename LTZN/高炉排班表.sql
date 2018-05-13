declare
t Date;
begin
t:=to_date('2006-07-06','YYYY-MM-DD');
while t<to_date('2008-07-01','YYYY-MM-DD') loop
insert into ¸ßÂ¯ÅÅ°à±í values(t,'Ò¹°à','±û°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t,'ÖÐ°à','¶¡°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t,'°×°à','¶¡°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+1,'Ò¹°à','¼×°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+1,'ÖÐ°à','ÒÒ°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+1,'°×°à','ÒÒ°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+2,'Ò¹°à','¶¡°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+2,'ÖÐ°à','±û°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+2,'°×°à','ÒÒ°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+3,'Ò¹°à','ÒÒ°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+3,'°×°à','¼×°à');
insert into ¸ßÂ¯ÅÅ°à±í values(t+3,'°×°à','ÒÒ°à');
t:=t+4;
end loop;
commit;
end;