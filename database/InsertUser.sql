conn lf/lf@liantie;
set echo on;
spool log/InsertUser.log;

Delete from UserRole where UserName='admin';
Delete from Users where yonghu='admin';
Insert Into Users(yonghu,mima) values('admin','admin');
Insert Into UserRole(UserName,Role) values('admin','π‹¿Ì‘±');

spool off;
exit;