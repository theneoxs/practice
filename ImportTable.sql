go
sp_configure 'show advanced options', 1;
RECONFIGURE;
GO
sp_configure 'Ad Hoc Distributed Queries', 1;
RECONFIGURE;
go
declare @path nvarchar(100);
set @path = 'C:\Users\Lulz\Desktop\Targets - Sample.xlsx';

declare @year varchar(10);
set @year = '2017';
print(@path);

/*Создание переменной с таблицей*/
declare @table TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table2 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table3 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table4 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table5 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table6 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);
declare @table7 TABLE(co int, fullname nvarchar(50), groups nvarchar(50), jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real);

/*
Вопрос, как сделать?
declare @sql nvarchar(max);
set @sql = N'insert into @tableOUT select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, '', '', ''*''), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', ''Excel 12.0; Database='+@path+''', ''SELECT * FROM [Invoice$A2:N57]'')'
print(@sql);

EXEC sp_executesql @sql, N'@tableOUT TABLE(co int, fullname nvarchar(50), groups nvarchar(50), 
jan real, feb real, mar real, apr real, may real, jun real, jul real, aug real, sep real, oct real, [now] real, [dec] real)';
select * from @table;
*/

/*Импорт таблиц из файла с заменой в имени точки с запятой на звездочку для дальнейшего разбиения на фамилию имя и нумерацией всех строк в алфавитном порядке*/
insert into @table select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [Invoice$A2:N57]')
insert into @table2 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [Prod Amts$A2:N57]');
insert into @table3 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [Production Hours$A2:N57]');
insert into @table4 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [GP$A2:N57]');
insert into @table5 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [AR Aging$A2:N57]');
insert into @table6 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [Net Bill Rates$A2:N57]');
insert into @table7 select ROW_NUMBER() over(order by Partner asc), REPLACE(Partner, ', ', '*'), [Practice Group], [January], [February], [March], [April], [May], [June], [July], [August], [September], [October], [November], [December] 
from OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database=C:\Users\Lulz\Desktop\Targets - Sample.xlsx', 'SELECT * FROM [WAR Realization$A2:N57]');
/*Добавление уникальных групп в таблицу*/
insert into practice.dbo.PracticeGroup (GroupName) (select DISTINCT [groups] from @table EXCEPT select GroupName from practice.dbo.PracticeGroup);

/*Инициализация переменной-счетчика*/
declare @count int;
set @count = 0;
declare @count_ins int;
set @count_ins = 1;

while @count < (select count(*) from @table)
begin
	set @count=@count + 1;
	/*Проверка условия наличия данного имени и фамилии в таблице*/
	if not exists (select * from practice.dbo.Partner where FirstName in 
	(select value from @table cross apply string_split(fullname, '*') where co = @count) and LastName in
	(select value from @table cross apply string_split(fullname, '*') where co = @count))
		begin
		/*Если запись отсутствует, то добавить запись, запрашивая ID группы по названию группы, значение фамилии после парсинга будет в первой строчке, а имени - во второй*/
		insert into practice.dbo.Partner (Practice_Group_ID, FirstName, LastName) VALUES (
		(select ID from practice.dbo.PracticeGroup where GroupName = (select groups from @table where co = @count)),
		(select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2),
		(select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1));
		end
	else
		begin
		if exists (select * from practice.dbo.PartnerTarget where partner_id = (select ID from practice.dbo.Partner 
		where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
		and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1))
		and YEAR(use_Date) = @year)
			begin
				delete from practice.dbo.PartnerTarget where partner_id = (select ID from practice.dbo.Partner 
				where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
				and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1))
				and YEAR(use_Date) = @year;
			end
		end
	print('done');
	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-01',
	(select jan from @table where co = @count),(select jan from @table2 where co = @count), (select jan from @table3 where co = @count), (select jan from @table4 where co = @count),
	(select jan from @table5 where co = @count), (select jan from @table6 where co = @count), (select jan from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-02',
	(select feb from @table where co = @count),(select feb from @table2 where co = @count), (select feb from @table3 where co = @count), (select feb from @table4 where co = @count),
	(select feb from @table5 where co = @count), (select feb from @table6 where co = @count), (select feb from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-03',
	(select mar from @table where co = @count),(select mar from @table2 where co = @count), (select mar from @table3 where co = @count), (select mar from @table4 where co = @count),
	(select mar from @table5 where co = @count), (select mar from @table6 where co = @count), (select mar from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-04',
	(select apr from @table where co = @count),(select apr from @table2 where co = @count), (select apr from @table3 where co = @count), (select apr from @table4 where co = @count),
	(select apr from @table5 where co = @count), (select apr from @table6 where co = @count), (select apr from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-05',
	(select may from @table where co = @count),(select may from @table2 where co = @count), (select may from @table3 where co = @count), (select may from @table4 where co = @count),
	(select may from @table5 where co = @count), (select may from @table6 where co = @count), (select may from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-06',
	(select jun from @table where co = @count),(select jun from @table2 where co = @count), (select jun from @table3 where co = @count), (select jun from @table4 where co = @count),
	(select jun from @table5 where co = @count), (select jun from @table6 where co = @count), (select jun from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-07',
	(select jul from @table where co = @count),(select jul from @table2 where co = @count), (select jul from @table3 where co = @count), (select jul from @table4 where co = @count),
	(select jul from @table5 where co = @count), (select jul from @table6 where co = @count), (select jul from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-08',
	(select aug from @table where co = @count),(select aug from @table2 where co = @count), (select aug from @table3 where co = @count), (select aug from @table4 where co = @count),
	(select aug from @table5 where co = @count), (select aug from @table6 where co = @count), (select aug from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-09',
	(select sep from @table where co = @count),(select sep from @table2 where co = @count), (select sep from @table3 where co = @count), (select sep from @table4 where co = @count),
	(select sep from @table5 where co = @count), (select sep from @table6 where co = @count), (select sep from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-10',
	(select oct from @table where co = @count),(select oct from @table2 where co = @count), (select oct from @table3 where co = @count), (select oct from @table4 where co = @count),
	(select oct from @table5 where co = @count), (select oct from @table6 where co = @count), (select oct from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-11',
	(select [now] from @table where co = @count),(select [now] from @table2 where co = @count), (select [now] from @table3 where co = @count), (select [now] from @table4 where co = @count),
	(select [now] from @table5 where co = @count), (select [now] from @table6 where co = @count), (select [now] from @table7 where co = @count));

	insert into practice.dbo.PartnerTarget values((select ID from practice.dbo.Partner 
	where FirstName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 2) 
	and LastName = (select value from (select ROW_NUMBER() over(order by co asc) as 'num', value from @table cross apply string_split(fullname, '*') where co = @count) a where num = 1)),
	''+@year+'-01-12',
	(select [dec] from @table where co = @count),(select [dec] from @table2 where co = @count), (select [dec] from @table3 where co = @count), (select [dec] from @table4 where co = @count),
	(select [dec] from @table5 where co = @count), (select [dec] from @table6 where co = @count), (select [dec] from @table7 where co = @count));
end

go