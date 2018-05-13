conn lf/lf@liantie;
set echo on;
spool log/UserRole.log;

drop table UserRole;
drop table Roles;

--角色表
create table Roles(
  Name varchar2(255),
  Description varchar2(255),
  CONSTRAINT RolesKey PRIMARY KEY(Name) 
);
insert into Roles(name) values('管理员');
insert into Roles(name) values('调度');
insert into Roles(name) values('统计');
insert into Roles(name) values('普通');
insert into Roles(name) values('1高炉');
insert into Roles(name) values('2高炉');
insert into Roles(name) values('3高炉');
insert into Roles(name) values('4高炉');
insert into Roles(name) values('5高炉');
--用户角色关系表
create table UserRole(
 UserName varchar2(255),
 Role varchar2(255),
 CONSTRAINT UserRoleKey PRIMARY KEY(UserName,Role), 
 CONSTRAINT UserRoleForeignUser FOREIGN KEY(UserName) 
    REFERENCES Users(yonghu), 
 CONSTRAINT UserRoleForeignRole FOREIGN KEY(Role) 
    REFERENCES Roles(Name)
 );
spool off;
exit;