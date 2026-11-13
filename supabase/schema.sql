-- 體育組差旅概算 － Supabase 資料表
-- 在 Supabase Dashboard → SQL Editor 貼上整份執行一次即可（可重複執行，會清空重建）。
-- 資料來源：基本補助費概算表格式.xlsm 的「競賽資訊」與「費率對照」【一】【二】【三】。

drop table if exists competitions, regions, city_regions, fare_routes cascade;

-- 競賽資訊：每列一場比賽
create table competitions (
  id          uuid primary key default gen_random_uuid(),
  sort        integer not null default 0,
  name        text,
  start_date  date,
  end_date    date,
  location    text,
  players     integer,
  staff       integer,
  prep        text default '否',          -- 需事前移地準備：是／否
  mode        text default '大眾運輸',    -- 大眾運輸／租賃車輛／每日通勤
  fare        numeric,                    -- 單趟票價（空白＝查【三】）
  items       numeric,                    -- 每人項目數
  per_item    numeric,                    -- 每項報名費
  clean_fee   numeric,                    -- 每人清潔費
  fund_source text default '基本補助費',
  updated_at  timestamptz not null default now()
);

-- 【一】地區費率表
create table regions (
  id          uuid primary key default gen_random_uuid(),
  sort        integer not null default 0,
  name        text,
  meal        numeric,   -- 膳費 元/人/日
  room        numeric,   -- 住宿費 元/人/晚
  fare_way    text,      -- 每日往返一次／往返一次
  car_cap     numeric,   -- 租賃車每日上限
  commute_cap numeric,   -- 每日通勤上限 元/人/日
  updated_at  timestamptz not null default now()
);

-- 【二】縣市關鍵字 → 地區（依 sort 由小到大比對，先命中者為準）
create table city_regions (
  id         uuid primary key default gen_random_uuid(),
  sort       integer not null default 0,
  keyword    text,
  region     text,
  note       text,       -- 含「★待確認」者概算時會提醒
  updated_at timestamptz not null default now()
);

-- 【三】路線車資表（出發地：楊梅；「＊預設」為查無路線時的後備）
create table fare_routes (
  id         uuid primary key default gen_random_uuid(),
  sort       integer not null default 0,
  keyword    text,
  fare       numeric,    -- 單趟票價
  detail     text,       -- 路線明細（可多行）
  extra      text,       -- 地點補充說明
  updated_at timestamptz not null default now()
);

-- 只有登入的帳號可以讀寫（Authentication → Users 新增帳號）
do $$
declare t text;
begin
  foreach t in array array['competitions','regions','city_regions','fare_routes'] loop
    execute format('alter table %I enable row level security', t);
    execute format('create policy "authenticated read write" on %I for all to authenticated using (true) with check (true)', t);
  end loop;
end $$;

-- 初始資料
insert into competitions (sort,name,start_date,end_date,location,players,staff,prep,mode,fare,items,per_item,clean_fee,fund_source) values
(0,'全國青少年跆拳道錦標賽','2026-09-08','2026-09-10','台南市立新營體育館',3,1,'否','大眾運輸',null,1,100,200,'基本補助費'),
(10,'柔道錦標賽','2026-10-10','2026-10-15','新竹火車站',5,1,'否','大眾運輸',null,1,100,200,'基本補助費');

insert into regions (sort,name,meal,room,fare_way,car_cap,commute_cap) values
(0,'桃園市',250,0,'每日往返一次',6000,0),
(10,'北部地區',300,0,'每日往返一次',8000,0),
(20,'中部、南部及東部地區',300,700,'往返一次',12000,700),
(30,'外島地區',300,700,'往返一次',6000,0);

insert into city_regions (sort,keyword,region,note) values
(0,'桃園','桃園市',null),
(10,'臺北','北部地區',null),
(20,'台北','北部地區',null),
(30,'新北','北部地區',null),
(40,'基隆','北部地區',null),
(50,'新竹','北部地區',null),
(60,'苗栗','中部、南部及東部地區',null),
(70,'臺中','中部、南部及東部地區',null),
(80,'台中','中部、南部及東部地區',null),
(90,'彰化','中部、南部及東部地區',null),
(100,'南投','中部、南部及東部地區',null),
(110,'雲林','中部、南部及東部地區',null),
(120,'嘉義','中部、南部及東部地區',null),
(130,'臺南','中部、南部及東部地區',null),
(140,'台南','中部、南部及東部地區',null),
(150,'高雄','中部、南部及東部地區',null),
(160,'屏東','中部、南部及東部地區',null),
(170,'宜蘭','中部、南部及東部地區',null),
(180,'花蓮','中部、南部及東部地區',null),
(190,'臺東','中部、南部及東部地區',null),
(200,'台東','中部、南部及東部地區',null),
(210,'澎湖','外島地區',null),
(220,'金門','外島地區',null),
(230,'連江','外島地區',null),
(240,'馬祖','外島地區',null),
(250,'中壢','桃園市',null),
(260,'楊梅','桃園市',null),
(270,'平鎮','桃園市',null),
(280,'八德','桃園市',null),
(290,'龜山','桃園市',null),
(300,'蘆竹','桃園市',null),
(310,'龍潭','桃園市',null),
(320,'大園','桃園市',null),
(330,'新屋','桃園市',null),
(340,'竹北','北部地區',null),
(350,'竹東','北部地區',null),
(360,'湖口','北部地區',null),
(370,'板橋','北部地區',null),
(380,'三重','北部地區',null),
(390,'新莊','北部地區',null),
(400,'淡水','北部地區',null),
(410,'汐止','北部地區',null),
(420,'竹南','中部、南部及東部地區',null),
(430,'頭份','中部、南部及東部地區',null),
(440,'豐原','中部、南部及東部地區',null),
(450,'沙鹿','中部、南部及東部地區',null),
(460,'員林','中部、南部及東部地區',null),
(470,'斗六','中部、南部及東部地區',null),
(480,'虎尾','中部、南部及東部地區',null),
(490,'新營','中部、南部及東部地區',null),
(500,'岡山','中部、南部及東部地區',null),
(510,'鳳山','中部、南部及東部地區',null),
(520,'羅東','中部、南部及東部地區',null),
(530,'馬公','外島地區',null);

insert into fare_routes (sort,keyword,fare,detail,extra) values
(0,'新竹',64,'(區間)楊梅-新竹64元',null),
(10,'竹北',64,'(區間)楊梅-新竹64元',null),
(20,'臺北',133,E'(區間)楊梅-中壢24元\n(自強)中壢-臺北89元\n(公車)20元',null),
(30,'台北',133,E'(區間)楊梅-中壢24元\n(自強)中壢-臺北89元\n(公車)20元',null),
(40,'新北',133,E'(區間)楊梅-中壢24元\n(自強)中壢-臺北89元\n(公車)20元',null),
(50,'基隆',133,E'(區間)楊梅-中壢24元\n(自強)中壢-臺北89元\n(公車)20元',null),
(60,'桃園',18,'(區間)楊梅-賽場18元',null),
(70,'中壢',18,'(區間)楊梅-賽場18元',null),
(80,'楊梅',18,'(區間)楊梅-賽場18元',null),
(90,'平鎮',18,'(區間)楊梅-賽場18元',null),
(100,'新營',709,E'(區間)楊梅-新竹64元\n(自強)新竹-新營619元\n(大台南公車棕3線)新營車站-新營體育館26元','臺南市新營區長榮路二段78號(約217公里)'),
(110,'臺中',362,E'(區間)楊梅-新竹64元\n(自強)新竹-臺中278元\n(公車)20元',null),
(120,'台中',362,E'(區間)楊梅-新竹64元\n(自強)新竹-臺中278元\n(公車)20元',null),
(130,'新左營',693,E'(區間)楊梅-新竹64元\n(自強)新竹-新左營609元\n(高雄捷運)新左營-世運站20元','世運主場館'),
(140,'＊預設',0,'（查無路線：請於「競賽資訊」I欄填入實際單趟票價，並檢附票價圖）',null);

