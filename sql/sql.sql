--------------------------------------------------------------------------------
-- managing database
--------------------------------------------------------------------------------
if exists(select name from master.dbo.sysdatabases where name = '_products')
begin
    use master
    alter database products set offline
    exec sp_detach_db 'products', 'true'
    drop database products

    -- only with privileges
    exec xp_cmdshell 'del c:\users\...\products.mdf'
    exec xp_cmdshell 'del c:\users\...\products_log.ldf'
end

if not exists(select name from master.dbo.sysdatabases where name = 'products')
    create database products

go
    use products -- force statement

--------------------------------------------------------------------------------
-- managing tables
--------------------------------------------------------------------------------
if exists(select name from products.sys.tables where name = 'suppliers')
    drop table suppliers

create table suppliers
(
    id          int not null,
    loc         varchar(100) not null,
    since       date,
    deliverable smallint,
    -- constraints introduces transparency and simplifies troubleshooting
    -- multiple unique constraints possible, but only one primary key
    -- unique constraint creates unique non-clustered index automatically (may nullable)
    -- primary key generates unique clustered index (not nullable)
    constraint pk_rule primary key clustered(id asc, loc) -- clustering: physical storage order
);

insert into suppliers (id, loc, since, deliverable)
values (934, 'London', '2020-04-10', 3),
       (224, 'Berlin', '2010-12-05', 5),
       (148, 'New York', '2015-08-02', 2),
       (223, 'Miami', '2012-10-27', 2)

-- short form
insert /* into */ suppliers /* columns, ... */
values (882, 'Tokyo', '2019-01-30', 4) /* has to match table definition */

if exists(select name from products.sys.tables where name = 'customers')
    drop table customers

create table customers
(
    id      int identity(1,1) primary key,
    company varchar(50)
);

insert into customers (company) values ('a'), ('b'), ('c');

if exists(select name from products.sys.tables where name = 'bookings')
    drop table bookings

create table bookings
(
    id   int,
    item char(3)
);

insert into bookings (id, item)
values (1, 'x'), (2, 'x'), (1, 'y'), (1, 'x'), (3, 'z'), (2, 'z'), (2, 'x');

if exists(select name from products.sys.tables where name = 'fruits')
    drop table fruits

create table fruits
(
    product_id int identity(10000,1),
    product_name char(50) not null, -- shan't support unicode, might not displayed correctly
    entry_date date default getdate() not null,
    category_id int check (category_id between 1 and 100), -- plausibility check
    discount decimal(5,2) default 0.05,
    constraint ruleCheckDiscount check(discount >= 0 and discount <= 0.1),
    constraint product_pk primary key(product_id)
);

insert into fruits (product_name, category_id)
values ('Pear', 50),
       ('Banana', 50),
       ('Orange', 50),
       ('Apple', 50),
       ('Bread', 75),
       ('Sliced Ham', 25),
       ('Kleenex', null)

-- error management
begin try
    insert into fruits (product_name, category_id, discount)
    values ('Pinapple', 50, 0.03),
           ('Lemon', 50, 0.3)
end try
begin catch
    print 'error'
    raiserror('invalid discount', 10, 1)
end catch

-- random data
if exists(select name from products.sys.tables where name = 'goods')
    drop table goods

create table goods(id int, price money, groupname text, code varchar(100));
create index ix_code on goods(code);

declare @count int
begin
    set @count = 0
    while @count < 1000
    begin
        set @count = @count + 1
        insert into goods
        values
        (
            @count,
            round(rand() * 1000, 2),
            'group_' + cast(@count / 2 as varchar),
            cast(NEWID() as varchar(100))
        )
    end
end

if exists(select name from products.sys.tables where name = 'groups')
    drop table groups

create table groups
(
    id text,
    description text
);

insert into groups 
values ('group_0', 'electrical equipment'),
       ('group_1', 'furniture'),
       ('group_2', 'tools'),
       ('group_3', 'storage systems'),
       ('group_4', 'detergents')

if exists(select name from products.sys.tables where name = 'pizza_companies')
begin
    drop table pizza_companies
end

create table pizza_companies
(
    id int identity(1,1) primary key clustered,
    name varchar(50),
    city varchar(30)  
);

set identity_insert pizza_companies on; -- insert despite of identity rule
insert into pizza_companies (id, name, city)
values (1, 'Dominos', 'Los Angeles'), 
       (2, 'Pizza Hut', 'San Francisco'),
       (3, 'Papa Johns', 'San Diego'),
       (4, 'Ah Pizz', 'Fremont'),
       (5, 'Nino Pizza', 'Las Vegas'),
       (6, 'Pizzeria', 'Boston'),
       (7, 'Chuck e cheese', 'Chicago')

if exists(select name from products.sys.tables where name = 'foods')
    drop table foods

create table foods
(
    id int primary key clustered, 
    name varchar(50), 
    unitssold int,
    company int,
    -- foreign key constraint
    constraint
    fk_company foreign key (company) references pizza_companies(id)
)

alter table foods
drop constraint fk_company

insert into foods (id, name, unitssold, company)
values (1, 'Large Pizza', 5, 2),
       (2, 'Garlic Knots', 6, 3),
       (3, 'Large Pizza', 3, 3),
       (4, 'Medium Pizza', 8, 4),
       (5, 'Breadsticks', 7, 1),
       (6, 'Medium Pizza', 11, 1),
       (7, 'Small Pizza', 9, 6),
       (8, 'Small Pizza', 6, 7)

/* primary/foreign key */

-- primary
drop table if exists pri
create table pri
(
    /* identifier type [identity] [primary key] */
    a int primary key,
    b int identity(1,1)
);

drop table if exists pri
create table pri
(
    a int,

    -- constraint
    primary key (a)
)

drop table if exists pri
create table pri
(
    a int identity(1,1),
    primary key (a)
)

drop table if exists pri
create table pri
(
    a int identity(1,1),

    -- constraint
    constraint a_pk_rule primary key (a)
)

drop table if exists pri
create table pri
(
    a int primary key
);

drop table if exists sec
create table sec
(
    a int references pri(a) -- short form foreign key
);

drop table if exists sec

create table sec
(
    a int foreign key references pri(a)
);

drop table if exists sec

create table sec
(
    a int

    -- constraint
    foreign key references pri(a)
);

drop table if exists sec

-- multiple primary keys require constraint
drop table if exists pri
create table pri
(
    a int,
    b int,
    primary key clustered (a asc, b asc) -- clustering may increase performance
);

drop table if exists sec
create table sec
(
    a int,
    b int,
    foreign key(a, b) references pri(a, b)
);

drop table if exists sec

drop table if exists pri
create table pri
(
    a int,
    b int,
    constraint mpk_rule primary key clustered (a asc, b asc) -- clustering may increase performance
);

drop table if exists sec
create table sec
(
    a int,
    b int,
    constraint fk_rule foreign key(a, b) references pri(a, b) on update cascade on delete cascade
);

drop table if exists sec

/* login */
if 0 = 1
begin
    create login userlogin with password = 'test';
    create user group_0 for login userlogin
    exec sp_addrolemember 'db_owner', 'group_0'
    exec sp_addrolemember 'db_datareader', 'group_0'
    grant select on groups to group_0
end

/* bulk insert */
drop table if exists items
create table items (id int, spec text);

if 0 = 1
begin
    bulk insert items
        from '...\items.txt'
        with
        (
            firstrow        = 1,
            fieldterminator = ';',
            rowterminator   = '\n',
            codepage        = 'acp',
            tablock
        )
end

--------------------------------------------------------------------------------
-- queries
--------------------------------------------------------------------------------

-- fruits table
if 0 = 1
    select * from fruits

-- relation magnitude: u x v x w
if 0 = 1
    select * from bookings, customers, suppliers

-- merge with union
-- merging distinct rows from bookings with customers
-- apply labels from first table
-- values from both tables
if 0 = 1
    select *, 'Bookings' as fromTable from bookings
    union
    select *, 'Customers' as fromTable from customers

/* wildcards */
if 0 = 1
begin
    select * from suppliers
    select * from suppliers where since like '20_0%'    -- % any symbol, _ any but one symbol
    select * from suppliers where since like '201[05]%' -- [] operator
    select * from suppliers where loc   like '_[^eo]%'  -- ^ not

    select product_name from fruits
    where product_name like '[a,b]%'
    order by product_name

    select category_id, product_name from fruits
    where category_id like '[25]%'

    select product_name from fruits
    where product_name like '[a-b]_[^e]%' -- a or b, one symbol, not e, any symbol
end

/* case when */
if 0 = 1
    select id, 
    case
        when id > 2 then 'v11'
        when id < 2 then 'v13'
    end as 'vendor'
    from bookings

/* between */
if 0 = 1
begin
    select city from pizza_companies 
    where city between 'c' and 'l' or city like 'l%'
    order by city

    select city from pizza_companies 
    where city between 'm' and 'c' -- wrong order → no data
end

/* charindex(searchstr, sourcestr) */
-- returns the starting position of the found item
-- finds only the first occurence of the match is returned
-- returns 0 if the search string couldn't be found
if 0 = 1
    select product_name, name from fruits, pizza_companies
    where product_name between 'A' and 'C' and CHARINDEX('pizz', name) > 0
    order by product_name desc

-- considering spelling/capitalization
if 1 = 0
begin
    drop table vendors
    create table vendors(nickname varchar(30));
    insert vendors values ('tom'), ('Tim'), ('theodore')

    -- CS: case sensitive
    select * from vendors
    where CHARINDEX('t', nickname COLLATE Latin1_General_CS_AS) > 0

    -- CI: case insensitive
    select * from vendors
    where CHARINDEX('t', nickname COLLATE Latin1_General_CI_AS) > 0
end

/* update */
if 0 = 1
begin
    select * from foods
    update foods set unitssold = unitssold + 1
    where company = 3
    select * from foods
end

/* select first n rows */
if 0 = 1
begin
    select top 3 *    from foods -- all fields
    select top 3 name from foods -- field [name]
    select top 50 percent name from foods
    select top 50 percent *    from foods
    select top 5 sum(unitssold) from foods
end

/* aggregation */
if 0 = 1
begin
    select max(unitssold) from foods
    select min(unitssold) from foods
    select top 10 min(unitssold) from foods
end

/* grouping */
-- counting number of equal rows per id/item group
if 0 = 1
    select *, count(id) as "# groups" from bookings
    group by id, item

if 0 = 1
    select substring(groupname, 1, 5) /* as */ groupname, sum(price) /* as */ sum
    from goods
    group by substring(groupname, 1, 5) -- can't group on text type directly

/* join */

-- inner join (equi join, join)
-- equal ids between customers and bookings
if 0 = 1
begin
    select * from customers inner join bookings on customers.id = bookings.id
    -- same result
    select * from customers       join bookings on customers.id = bookings.id
end

-- left join
if 0 = 1
begin
    select * from bookings left outer join suppliers on bookings.id = deliverable
    -- same result
    select * from bookings left       join suppliers on bookings.id = deliverable
end

-- right join
if 0 = 1
begin
    select deliverable, bookings.id from bookings right outer join suppliers
    on deliverable = bookings.id
    -- same result
    select deliverable, bookings.id from bookings right       join suppliers
    on deliverable = bookings.id
end

-- full join
if 0 = 1
    select deliverable, bookings.id from bookings full join suppliers
    on deliverable = bookings.id

-- multiple sources in from clause linked by condition alternatively
if 0 = 1
    select bookings.id, item, company
    from bookings, customers
    where bookings.id = customers.id
    order by item

/* subquery, nested query */
-- counting different item groups per id
-- asterix in the enclosing expression can't be used here
-- nested expression need to be aliased
if 0 = 1
( -- instead of begin/end, only for single expressions (select)
    select id, count(id)
    from
    (
        select *
        from bookings
        group by id, item
    ) as items
    group by id
)

if 0 = 1
    -- # booked items per id linked to deliverable items by suppliers
    select distinct loc, deliverable, _bookings.numBookings from suppliers
    inner join
    (
        select id, count(id) as numBookings -- aliasing mandatory here 
        from bookings 
        group by id
    ) as _bookings
    on _bookings.numBookings = deliverable
    -- returns one (distinct) data set (matching deliverable = numBookings = 3)

if 0 = 1
    select count(product_name)
    from
    (
        select product_name
        from fruits
        where category_id = 50
    ) as subq

if 0 = 1
    select power(unitssold - avq.av, 2)
    from
        foods,
        (
            select avg(unitssold) as av from foods
        ) as avq 

-- standard deviation
if 0 = 1
begin
    drop table measurement
    create table measurement (val decimal(10,4)); -- precision, scale
    insert measurement values (4.34), (9.23), (1.43), (5.24), (3.42)
    select * from measurement

    -- optimized counting (single field)
    select sqrt(sum(power(val - subq.av, 2)) / count(1))
    from
        measurement,
        (
            select avg(val) av from measurement
        ) subq
end

-- subquery not permitted to return more than one value if followed by operator
if 0 = 1
begin
    select id from foods where (select distinct company  from foods) = company -- error
    select id from foods where (select      avg(company) from foods) = company -- ok
end

-- selfjoin
if 0 = 1
select distinct f.company from
    foods f,
    (
        select company  from foods where company = 4
    ) subq
    where f.company = subq.company

if 0 = 1
begin
    drop table _foods
    select distinct id i1, id i2 into _foods from foods
    select * from _foods
    select f1.i1, f2.i2 from _foods f1 join _foods f2 on f1.i1 = f2.i2 and f1.i1 < 4
end

/* formatting */
if 0 = 1
    select format(price, 'C', 'th-TH') 'Thailand',
           format(price, 'C', 'de-DE') 'Germany',
           format(price, 'C', 'en-gb') 'British',
           format(price, 'C', 'en-us') 'US',
           format(price, 'C', 'en-in') 'India',
           format(price, 'C', 'fr-FR') 'France',
           format(price, 'C', 'zh-cn') 'China'
    from goods
