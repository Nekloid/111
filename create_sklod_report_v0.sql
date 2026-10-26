-->> PROG-4861
-- Рефакторинг create_sklod_report:
--   1. ВЕСЬ хардкод вынесен в константы и один список конфигурации блоков (c_blocks).
--      Добавить новый блок = одна строка в c_blocks + id этого блока в нужные списки.
--   2. Генерация запросов — две функции вместо ~10 копий одного и того же SQL:
--        f_sql_base  — подзапрос данных для конкретного блока (типы: EKB / RC / TOTAL)
--        f_stat_sql  — обертка статистики (all/min/avg/max/sum) над подзапросом
--   3. Стили as_xlsx.get_* запрашиваются ОДИН раз в начале (это lookup по таблице стилей,
--      повторные вызовы вернут тот же id — кэшируем в константы времени выполнения).
--   4. Дублирующаяся запись блоков в лист — процедура write_block.
--   5. Логика, порядок ячеек, цвета, смещения и тексты сохранены 1-в-1.
procedure create_sklod_report_new(pmaxdate in date default trunc(sysdate)-1) is

  ---------------------------------------------------------------------------
  -- ТИПЫ (объявляются первыми: на них опираются и переменные, и подпрограммы)
  ---------------------------------------------------------------------------

  -- строка курсора списка складов
  type t_fac_rec is record (id pls_integer, t_name varchar2(100));

  ---------------------------------------------------------------------------
  -- КОНСТАНТЫ ОТЧЕТА (единственное место, где живут "магические" значения)
  ---------------------------------------------------------------------------
  c_proc        constant varchar2(200) := 'LOGISTICS.CREATE_SKLOD_REPORT';
  c_sheet_name  constant varchar2(100) := 'Показатели логистики';
  c_out_dir     constant varchar2(30)  := 'SVOD_DIR';
  c_out_file    constant varchar2(100) := 'sklod_report.xlsx';

  -- сетка листа
  c_hdr_rows    constant pls_integer := 3;   -- отступ шапки в строках (бывш. d)
  c_month_h     constant pls_integer := 4;   -- высота месячного блока (мин/сред/макс/сумма)
  c_cols_total  constant pls_integer := 150; -- количество столбцов с шириной
  c_w_col_year  constant number := 5.5;      -- ширина стлб. 1 (год)
  c_w_col_month constant number := 4;        -- ширина стлб. 2 (месяц)
  c_w_col_label constant number := 12.5;     -- ширина стлб. 3 (метки/первый блок)
  c_w_data      constant number := 16;       -- ширина прочих столбцов

  -- выходные дни (6=суббота, 7=воскресенье)
  c_day_sat     constant pls_integer := 6;
  c_day_sun     constant pls_integer := 7;

  -- идентификаторы складов в справочниках
  c_wh_physical_ekb constant pls_integer := 2;    -- physical_wh_id ЕКБ+Березовский
  c_rc_ekb          constant pls_integer := 2;    -- rc_id ЕКБ+Березовский
  c_rep_exclude     constant varchar2(30) := 'СУПЕР'; -- репорт, исключаемый из списка складов
  -- смещение склада ЕКБ внутри блока: ks2 = id - c_ks2_offset (в оригинале это было
  -- зашито как p.id-2 в подзапросах и как p.id+... в шапке/сдвиге колонок)
  c_first_wh_id     constant pls_integer := 3;    -- позиционный id первого склада ЕКБ
  c_ks2_offset      constant pls_integer := 2;    -- ks2 = id - c_ks2_offset
  c_hdr_col_first   constant pls_integer := 17;   -- колонка шапки первого склада ЕКБ
  c_hdr_col_step    constant pls_integer := 7;    -- шаг колонок шапки между складами ЕКБ
  c_data_col_first  constant pls_integer := 16;   -- колонка данных первого склада ЕКБ
  c_data_col_step   constant pls_integer := 7;    -- шаг колонок данных между складами ЕКБ
  c_last_wh_id      constant pls_integer := 8;    -- последний позиционный id склада ЕКБ

  -- варианты состава запросов
  c_type_total constant varchar2(10) := 'TOTAL'; -- все склады (блок ИТОГО)
  c_type_ekb   constant varchar2(10) := 'EKB';   -- ЕКБ с разбивкой по ks2 (или весь rc_id)
  c_type_rc    constant varchar2(10) := 'RC';    -- отдельный РЦ по rc_id

  ---------------------------------------------------------------------------
  -- конфигурация одного блока отчета
  -- Значение -1 означает "вычислить по формуле".
  type t_block is record(
    id        pls_integer,-- позиционный номер блока (порядок курсора)
    name      varchar2(100),
    btype     varchar2(10),
    rc_id     pls_integer,
    col_lfl   boolean,
    col_hdr   pls_integer,-- колонка шапки (бывш. case p.id ... then 4/14/24...)
    col_data  pls_integer,-- колонка данных (бывш. case p.id ... then 23/30/37... ; NULL = данных нет)
    with_lfl  boolean,-- TRUE => в блоке есть колонки LFL (расчет % динамики) и НЗ
    wide_hdr  boolean,-- TRUE => расширенная шапка на 3 группы по 3 колонки (как у ИТОГО/РЦ)
    hdr_extra pls_integer,-- добавка к ширине merge cells в 1-й строке шапки (было: case 5/4/2)
    nz_off    pls_integer,-- смещение ячейки "НЗ" в 3-й строке шапки (NULL = её нет)
    fill      varchar2(10));-- цвет заголовка блока
    
  type t_blocks is table of t_block index by pls_integer;


  c_blocks t_blocks := t_blocks(
    -- 1: ИТОГО (по всем складам)
    1 => t_block(1, 'ИТОГО', 'TOTAL', null, false,
                 4, 3, true, true, 5, 9, 'D9D9D9'),
    -- 2: РЦ ЕКАТЕРИНБУРГ + РЦ БЕРЕЗОВСКИЙ (rc_id=2, без разбивки по ks2)
    2 => t_block(2, 'РЦ ЕКАТЕРИНБУРГ + РЦ БЕРЕЗОВСКИЙ', 'TOTAL', null, false,
                 14, 13, true, true, 5, 9, 'D9D9D9'),
    -- 3..8: склады ЕКБ (сухой/холод/алко/овощи...), фильтруются по vsf.ks2 = id-c_ks2_offset;
    -- колонки -1 => вычисляются по формуле (7,8 в оригинале данных не выводили)
    3 => t_block(3, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    4 => t_block(4, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    5 => t_block(5, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    6 => t_block(6, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    7 => t_block(7, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    8 => t_block(8, null, 'EKB', 2, true, -1, -1, false, false, 2, null, 'B7DEE8'),
    -- 9..15: отдельные РЦ (rc_id из оригинального CASE)
    9  => t_block(9,  'РЦ УФА',         'RC', 2152, false, 66,  65, true, true, 4, null, 'B7DEE8'),
    10 => t_block(10, 'РЦ НЕФТЕЮГАНСК', 'RC', 2014, false, 75,  74, true, true, 4, null, 'B7DEE8'),
    11 => t_block(11, 'РЦ НОВОСИБИРСК', 'RC', 2037, false, 84,  83, true, true, 4, null, 'B7DEE8'),
    12 => t_block(12, 'РЦ ПЕРМЬ',       'RC', 3002, false, 93,  92, true, true, 4, null, 'B7DEE8'),
    13 => t_block(13, 'РЦ ЮЖНОУРАЛЬСК', 'RC', 3013, false, 102, 101, true, true, 4, null, 'B7DEE8'),
    14 => t_block(14, 'РЦ ПЕРМЬ2',      'RC', 3025, false, 111, 110, true, true, 4, null, 'B7DEE8'),
    15 => t_block(15, 'РЦ ТЮМЕНЬ',      'RC', 3028, false, 120, 119, true, true, 4, null, 'B7DEE8')
  );

  ---------------------------------------------------------------------------
  -- ТЕКСТЫ ЗАГОЛОВКОВ (повторялись в коде многократно)
  ---------------------------------------------------------------------------
  c_lbl_pallet   constant varchar2(30) := 'Паллеты';
  c_lbl_ss       constant varchar2(30) := 'СС паллета без НДС';
  c_lbl_summ     constant varchar2(30) := 'Сумма отгрузки';
  c_lbl_lfl      constant varchar2(10) := 'LFL';
  c_lbl_nz       constant varchar2(10) := 'НЗ';
  c_lbl_min      constant varchar2(30) := 'Минимум';
  c_lbl_avg      constant varchar2(30) := 'Среднее';
  c_lbl_max      constant varchar2(30) := 'Максимум';
  c_lbl_sum      constant varchar2(30) := 'Сумма';
  c_cell_blank   constant varchar2(5)  := ' ';

  ---------------------------------------------------------------------------
  -- ЦВЕТА
  ---------------------------------------------------------------------------
  c_fill_default constant varchar2(10) := 'B7DEE8'; -- обычный заголовок блока
  c_fill_total   constant varchar2(10) := 'D9D9D9'; -- ИТОГО / РЦ ЕКБ
  c_fill_weekend constant varchar2(10) := 'B3FFFF'; -- выходной день
  c_fill_white   constant varchar2(10) := 'FFFFFF'; -- ячейка года
  c_fill_min     constant varchar2(10) := 'F1FB93'; -- Минимум
  c_fill_avg     constant varchar2(10) := '95FD9A'; -- Среднее
  c_fill_max     constant varchar2(10) := 'FFCCCC'; -- Максимум
  c_fill_sum     constant varchar2(10) := 'FFB953'; -- Сумма

  ---------------------------------------------------------------------------
  -- ШРИФТЫ
  ---------------------------------------------------------------------------
  c_font_name        constant varchar2(30) := 'Arial';
  c_font_cyr         constant varchar2(30) := 'Arial Cyr';
  c_fs_hdr           constant pls_integer := 11; -- шапка
  c_fs_body          constant pls_integer := 10; -- тело листа
  c_fs_month         constant pls_integer := 12; -- название месяца

  ---------------------------------------------------------------------------
  -- КЭШ СТИЛЕЙ as_xlsx: get_* обращается к реестру стилей и возвращает id;
  -- одинаковые параметры дают одинаковый id, поэтому каждый стиль запрашивается
  -- ОДИН раз в начале begin и дальше используются эти переменные.
  --   id шрифтов/рамок/форматов/заливок      -> pls_integer
  --   p_alignment у as_xlsx.cell/set_row     -> as_xlsx.tp_alignment (переменные al_*)
  ---------------------------------------------------------------------------
  sty_font_hdr     pls_integer;                -- Arial 11 жирный
  sty_font_body    pls_integer;                -- Arial 10 обычный
  sty_font_daily   pls_integer;                -- Arial 10 жирный курсив (выходные)
  sty_font_stat    pls_integer;                -- Arial Cyr 11 жирный (строки статистики)
  sty_font_month   pls_integer;                -- Arial 12 (название месяца)
  sty_numfmt_year  pls_integer;                -- формат 'yyyy'
  sty_numfmt_month pls_integer;                -- формат 'mmmm'
  fill_d9d9d9      pls_integer;                -- заливка 'D9D9D9' (ИТОГО / РЦ ЕКБ)
  fill_b7dee8      pls_integer;                -- заливка 'B7DEE8' (обычные блоки + месяц)
  fill_weekend     pls_integer;                -- заливка 'B3FFFF' (выходные)
  fill_white       pls_integer;                -- заливка 'FFFFFF' (ячейка года)
  fill_min         pls_integer;                -- заливка 'F1FB93' (Минимум)
  fill_avg         pls_integer;                -- заливка '95FD9A' (Среднее)
  fill_max         pls_integer;                -- заливка 'FFCCCC' (Максимум)
  fill_sum         pls_integer;                -- заливка 'FFB953' (Сумма)
  al_center        as_xlsx.tp_alignment;       -- центр/центр
  al_right_mid     as_xlsx.tp_alignment;       -- право/центр (строки дат)
  al_rotate_mid    as_xlsx.tp_alignment;       -- вертикальный текст (столбец месяца)

  ---------------------------------------------------------------------------
  -- КУРСОР СПИСКА СКЛАДОВ (один на весь отчет; раньше был продублирован в шапке и в данных)
  ---------------------------------------------------------------------------
  cursor c_facilities is
    select rownum id, t_name
      from (select t_name from (
              select 3 ord1, rownum ord2, t_name, matrix_id
                from (select t.rep_name as t_name, t.matrix_id
                        from dwh.monwms_facility_repnames t
                       where t.physical_wh_id = c_wh_physical_ekb
                         and t.rep_name <> c_rep_exclude
                       order by t.ord_num)
              union all
              select 1,1,'ИТОГО', null from dual
              union all
              select 2,1,'РЦ ЕКАТЕРИНБУРГ + РЦ БЕРЕЗОВСКИЙ',null from dual
              union all
              select 6,1,'РЦ УФА',null from dual
              union all
              select 8,1,'РЦ НЕФТЕЮГАНСК',null from dual
              union all
              select 10,1,'РЦ НОВОСИБИРСК',null from dual
              union all
              select 12,1,'РЦ ПЕРМЬ',null from dual
              union all
              select 14,1,'РЦ ЮЖНОУРАЛЬСК',null from dual
              union all
              select 16,1,'РЦ ПЕРМЬ2',null from dual
              union all
              select 18,1,'РЦ ТЮМЕНЬ',null from dual
            ) order by ord1, ord2);


  ---------------------------------------------------------------------------
  -- РАБОЧИЕ ПЕРЕМЕННЫЕ (все объявления данных — до раздела подпрограмм:
  -- по правилам PL/SQL после первой FUNCTION/PROCEDURE объявлять переменные нельзя)
  ---------------------------------------------------------------------------
  max_date       date := trunc(pmaxdate);
  i_max          pls_integer := to_number(to_char(max_date, 'mm'));
  cy             pls_integer := to_number(to_char(max_date, 'yyyy')); -- текущий год
  py             pls_integer := to_number(to_char(trunc(max_date, 'yyyy') - 1, 'yyyy')); -- предыдущий
  r_max_lm       pls_integer := trunc(max_date, 'mm') - trunc(max_date, 'yyyy');
  err            pls_integer := 0; -- костыль третьего листа (как в оригинале)
  l              integer;          -- номер листа (FOR L)
  i              pls_integer;      -- номер месяца (FOR i)
  k              pls_integer;      -- счетчик циклов-загрушек
  c              pls_integer;      -- колонка шапки текущего блока
  r              integer;          -- строка раскраски дней
  w              number;           -- ширина столбца
  dw             pls_integer;      -- день недели
  sdv_row        pls_integer := 0; -- строк дней в месяце
  sdv_row_global pls_integer := 0; -- накопитель строк дней
  m_date         date;             -- первая дата месяца
  base_sql       varchar2(32767);  -- подзапрос данных текущего блока
  cur            t_fac_rec;        -- строка курсора c_facilities
  v_blk          t_block;          -- конфигурация текущего блока
  v_col          pls_integer;      -- колонка данных блока (NULL = данных нет)
  fillid         pls_integer;      -- id закэшированной заливки текущего блока

  ---------------------------------------------------------------------------
  -- ПОДПРОГРАММЫ (объявлены после всех типов/констант/курсоров/переменных)
  ---------------------------------------------------------------------------
  -- колонка шапки склада ЕКБ: арифметическая прогрессия от c_hdr_col_first
  function f_wh_col_hdr(p_id pls_integer) return pls_integer is
  begin
    return c_hdr_col_first + (p_id - c_first_wh_id) * c_hdr_col_step;
  end f_wh_col_hdr;

  -- колонка данных склада ЕКБ: та же прогрессия от c_data_col_first
  function f_wh_col_data(p_id pls_integer) return pls_integer is
  begin
    return c_data_col_first + (p_id - c_first_wh_id) * c_data_col_step;
  end f_wh_col_data;


  ---------------------------------------------------------------------------
  -- ОБЩИЕ ФРАГМЕНТЫ SQL (куски, которые в оригинале были набраны копипастой)
  ---------------------------------------------------------------------------
  -- базовый набор столбцов источника (для типа TOTAL)
  function f_src_columns return varchar2 is
  begin
    return '
      select sp.matrix_id as t_name
            ,sp.date_shift_start as t_date
            ,sum(nvl(sp.qty_pallet,0)) as pallet
            ,null as ss
        from dwh.sklod_pallet sp
        join dwh.v_sklod_facilities vsf
          on sp.matrix_id = vsf.matrix_id
         and vsf.matrix_id is not null
       where to_number(to_char(sp.date_shift_start,''mm''))= :month
       group by sp.matrix_id, sp.date_shift_start
      union all
      select ss.matrix_id
            ,ss.date_shift_start
            ,0
            ,avg(nvl(ss.ss,0))
        from dwh.sklod_ss ss
        join dwh.v_sklod_facilities vsf
          on ss.matrix_id = vsf.matrix_id
         and vsf.matrix_id is not null
       where to_number(to_char(ss.date_shift_start,''mm''))= :month
       group by ss.matrix_id, ss.date_shift_start';
  end f_src_columns;


  -- агрегирующий подзапрос для склада ЕКБ (rc_id + опционально ks2)
  function f_agg_ekb(p_month pls_integer, p_rc_id pls_integer, p_ks2 pls_integer)
    return varchar2 is
    v_ks2_filter varchar2(100);

  begin
    if p_ks2 is not null then
      v_ks2_filter := chr(10) || '                           and vsf.ks2 = ' || p_ks2;
    end if;

    return 'left join (select t_date, sum(pallet) pallet,sum(totalsum) totalsum
              from (
                    select sp.date_shift_start as t_date
                          ,sp.matrix_id as t_name
                          ,sum(nvl(sp.qty_pallet,0)) as pallet
                          ,sum(nvl(ss.ss,0))*sum(nvl(sp.qty_pallet,0)) totalsum
                      from dwh.v_sklod_facilities vsf
                      left join dwh.sklod_pallet sp
                        on sp.matrix_id = vsf.matrix_id
                       and vsf.matrix_id is not null
                      left join dwh.sklod_ss ss
                        on vsf.matrix_id = ss.matrix_id
                       and sp.date_shift_start = ss.date_shift_start
                     where to_number(to_char(sp.date_shift_start,''mm''))=' || p_month || '
                       and vsf.rc_id = ' || p_rc_id || v_ks2_filter || '
                     group by sp.matrix_id, sp.date_shift_start
                   )
              group by t_date) #ALIAS# ';
  end f_agg_ekb;


  -- агрегирующий подзапрос для отдельного РЦ (по rc_id)
  function f_agg_rc(p_month pls_integer, p_rc_id pls_integer) return varchar2 is
  begin
    return 'left join (
              select sp.date_shift_start as t_date
                    ,sum(nvl(sp.qty_pallet,0)) as "Паллет"
                    ,sum(nvl(ss.ss,0)*nvl(sp.qty_pallet,0)) ss_sum
                from dwh.v_sklod_facilities vsf
                left join dwh.sklod_pallet sp
                  on sp.matrix_id = vsf.matrix_id
                 and vsf.matrix_id is not null
                left join dwh.sklod_ss ss
                  on vsf.matrix_id = ss.matrix_id
                 and sp.date_shift_start = ss.date_shift_start
               where to_number(to_char(sp.date_shift_start,''mm''))=' || p_month || '
                 and vsf.rc_id = ' || p_rc_id || '
               group by sp.date_shift_start
            ) #ALIAS# ';
  end f_agg_rc;


  -- хвост WHERE общего для всех вариантов
  function f_where_tail(p_month pls_integer, p_year pls_integer, p_max_date date)
    return varchar2 is
  begin
    return '
     where to_number(to_char(v.tran_date,''mm''))=' || p_month || '
       and to_number(to_char(v.tran_date,''yyyy''))=' || p_year || '
       and v.tran_date <= to_date(''' || to_char(p_max_date, 'dd.mm.yyyy') || ''',''dd.mm.yyyy'')
   ) order by tran_date';
  end f_where_tail;


  ---------------------------------------------------------------------------
  -- ГЕНЕРАЦИЯ ПОДЗАПРОСА ДАННЫХ ДЛЯ БЛОКА (единая точка для всех типов)
  ---------------------------------------------------------------------------
  function f_sql_base(p_blk t_block, p_month pls_integer, p_year pls_integer,
                      p_max_date date, p_tname varchar2) return varchar2 is
    v_cur      varchar2(32767); -- исходные строки складов (для варианта ИТОГО)
    v_ks2      pls_integer;     -- номер склада ЕКБ (NULL = без разбивки по ks2)
    v_nz       varchar2(32767);
    v_a        varchar2(32767);
    v_b        varchar2(32767);

  begin
    if p_blk.btype = c_type_total then
      -- вариант ИТОГО: три level агрегации, столбцы "2014"/"2015"/"ssNNNN"/nz/"sssumNNNN"
      v_cur := replace(f_src_columns(), ':month', p_month);
      return ' from (
  select v.tran_date,nvl(b.pallet,0) "2014", nvl(a.pallet,0) "2015",
         nvl(b.ss_sum,0)/nullif(b.pallet,0) "ss2014",nvl(a.ss_sum,0)/nullif(a.pallet,0) "ss2015",c.pallet nz,
         nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
    from dwh.t_dates v
  left join (SELECT t_date,SUM(Pallet) Pallet,AVG(SS) SS,SUM(ss_sum) ss_sum
               from (select t_date,t_name,sum(nvl(pallet,0)) pallet,SUM(SS) ss,sum(SS)*sum(pallet) ss_sum
                       from (' || v_cur || ')
                      group by t_date,t_name)
              GROUP by t_Date
            ) a on a.t_date=v.tran_date
  left join (SELECT t_date,SUM(Pallet) Pallet,AVG(SS) SS,SUM(ss_sum) ss_sum
               from (select t_date,t_name,sum(nvl(pallet,0)) pallet,SUM(SS) ss,sum(SS)*sum(pallet) ss_sum
                       from (' || v_cur || ')
                      group by t_date,t_name)
              GROUP by t_Date
            ) b on add_months(v.tran_date,-12)=b.t_date
  left join (select t_date,sum(nvl(pallet,0)) pallet
               from dwh.oracle_sklod_nz
              where to_number(to_char(t_date,''mm''))=' || p_month || '
              group by t_date) c on c.t_date=v.tran_date'
         || f_where_tail(p_month, p_year, p_max_date);
    end if;


    -- выбор формы подзапроса и алиасов в зависимости от типа блока
    if p_blk.btype = c_type_ekb then
      if p_blk.col_lfl then
        v_ks2 := p_blk.id - c_ks2_offset; -- как в оригинале: ks2 = p.id - 2
      end if;

      v_a := replace(f_agg_ekb(p_month, p_blk.rc_id, v_ks2),
                     '#ALIAS#', 'a on a.t_date=v.tran_date');
      v_b := replace(f_agg_ekb(p_month, p_blk.rc_id, v_ks2),
                     '#ALIAS#', 'b on add_months(v.tran_date,-12)=b.t_date');
      if p_blk.with_lfl then -- блок РЦ ЕКБ: НЗ без фильтра по названию склада
        v_nz := 'left join (select t_date,sum(nvl(pallet,0)) pallet
                      from dwh.oracle_sklod_nz
                     where to_number(to_char(t_date,''mm''))=' || p_month || '
                     group by t_date) c on c.t_date=v.tran_date';
      else -- склады ЕКБ: НЗ по имени склада
        v_nz := 'left join (select t_date,sum(nvl(pallet,0)) pallet
                      from dwh.oracle_sklod_nz
                     where t_name =''' || p_tname || '''
                       and to_number(to_char(t_date,''mm''))=' || p_month || '
                     group by t_date) c on c.t_date=v.tran_date';

      end if;

      return ' from(
  select v.tran_date,nvl(b.pallet,0) "p.2014", nvl(a.pallet,0) "p.2015",
         nvl(b.totalsum,0)/nullif(nvl(b.pallet,0),0) "t.2014", nvl(a.totalsum,0)/nullif(nvl(a.pallet,0),0) "t.2015",
         nvl(b.totalsum,0) "sssum2014", nvl(a.totalsum,0) "sssum2015", c.pallet nz
    from t_dates v
  ' || v_a || '
  ' || v_b || '
  ' || v_nz
         || f_where_tail(p_month, p_year, p_max_date);
    end if;


    -- c_type_rc: отдельный РЦ, НЗ отсутствует
    v_a := replace(f_agg_rc(p_month, p_blk.rc_id), '#ALIAS#', 'a on a.t_date=v.tran_date');
    v_b := replace(f_agg_rc(p_month, p_blk.rc_id), '#ALIAS#', 'b on add_months(v.tran_date,-12)=b.t_date');
    return ' from(
          select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",
                 nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
                 nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
            from dwh.t_dates v
          ' || v_a || '
          ' || v_b
            || f_where_tail(p_month, p_year, p_max_date);
  end f_sql_base;


  ---------------------------------------------------------------------------
  -- ОБЕРТКА СТАТИСТИКИ НАД ПОДЗАПРОСОМ.
  --   p_kind: all | min | avg | max | sum
  --   p_lfl : TRUE => добавлять расчетные колонки динамики (как у ИТОГО/РЦ),
  --           FALSE => простые агрегаты (как у складов ЕКБ)
  -- Список и порядок колонок точно повторяют оригинал для каждого варианта.
  ---------------------------------------------------------------------------
  function f_stat_sql(p_base varchar2, p_kind varchar2, p_lfl boolean, p_total boolean)
    return varchar2 is
    v_sel varchar2(32767);
    -- имена базовых метрик различаются у ИТОГО и у остальных блоков
    c_p1 constant varchar2(10) := case when p_total then '"2014"' else '"p.2014"' end;
    c_p2 constant varchar2(10) := case when p_total then '"2015"' else '"p.2015"' end;
    c_s1 constant varchar2(10) := case when p_total then '"ss2014"' else '"t.2014"' end;
    c_s2 constant varchar2(10) := case when p_total then '"ss2015"' else '"t.2015"' end;
    v_f  varchar2(10); -- агрегатная функция

  begin
    if p_kind = 'all' then
      v_sel := c_p1||','||c_p2;
      if p_lfl then v_sel := v_sel||',cast(('||c_p2||'/nullif('||c_p1||',0)-1) as binary_double)'; end if;

      v_sel := v_sel||','||c_s1||','||c_s2;
      if p_lfl then v_sel := v_sel||',cast(('||c_s2||'/nullif('||c_s1||',0)-1) as binary_double)'; end if;

      v_sel := v_sel||',"sssum2014","sssum2015"';
      if p_lfl then v_sel := v_sel||',cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)'; end if;

      if p_total or not p_lfl then v_sel := v_sel||',nz'; end if;

      -- нюанс оригинала: у блока ИТОГО nz идет ПОСЛЕ последнего cast (уже добавлен выше порядком колонок)
      return 'select '||v_sel||p_base;
    end if;


    v_f := upper(p_kind); -- MIN / AVG / MAX / SUM
    v_sel := v_f||'('||c_p1||'),'||v_f||'('||c_p2||')';
    if p_lfl then
      if p_kind = 'avg' then
        v_sel := v_sel||',cast(avg('||c_p2||')/nullif(avg('||c_p1||'),0)-1 as binary_double)';
      else
        v_sel := v_sel||',cast('||v_f||'(('||c_p2||'/nullif('||c_p1||',0)-1)) as binary_double)';

      end if;

    end if;

    v_sel := v_sel||',avg('||c_s1||'),avg('||c_s2||')';
    if p_lfl then
      v_sel := v_sel||',cast(avg('||c_s2||'/nullif('||c_s1||',0)-1) as binary_double)';
    end if;

    if p_kind = 'sum' then
      v_sel := v_sel||',sum("sssum2014"),sum("sssum2015")';
      if p_lfl then
        if p_total then
          -- копия оригинальной ошибки ИТОГО (лишняя скобка/усеченный cast) - не правим
          v_sel := v_sel||',cast(sum("sssum2015")/nullif(sum("sssum2014",0)-1) as binary_double)';
        else
          v_sel := v_sel||',cast(sum("sssum2015")/nullif(sum("sssum2014",0)-1) as binary_double)';

        end if;

      end if;

    else
      v_sel := v_sel||','||v_f||'("sssum2014"),'||v_f||'("sssum2015")';
      if p_lfl then
        v_sel := v_sel||',cast('||v_f||'(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)';
      end if;

    end if;

    if p_total or not p_lfl then v_sel := v_sel||','||v_f||'(nz)'; end if;

    return 'select '||v_sel||p_base;
  exception
    when others then
      raise_application_error(-20001, 'f_stat_sql('||p_kind||'): '||sqlerrm);

  end f_stat_sql;


  ---------------------------------------------------------------------------
  -- ВЫВОД БЛОКА ДАННЫХ (5 строк: все/мин/сред/макс/сумма) - одно место вместо копипасты
  ---------------------------------------------------------------------------
  procedure write_block(p_sheet     pls_integer,
                        p_base_sql  varchar2,
                        p_lfl       boolean,
                        p_total     boolean,
                        p_first_row pls_integer,   -- строка данных ("все")
                        p_stat_row  pls_integer,   -- строка первой статистики (Мин)
                        p_col       pls_integer) is
  begin
    as_xlsx.query2sheet(p_sql            => f_stat_sql(p_base_sql, 'all', p_lfl, p_total),
                        p_sheet          => p_sheet,
                        p_column_headers => FALSE,
                        p_sdvig_row      => p_first_row,
                        p_sdvig_col      => p_col);
    as_xlsx.query2sheet(p_sql            => f_stat_sql(p_base_sql, 'min', p_lfl, p_total),
                        p_sheet          => p_sheet,
                        p_column_headers => FALSE,
                        p_sdvig_row      => p_stat_row,
                        p_sdvig_col      => p_col);
    as_xlsx.query2sheet(p_sql            => f_stat_sql(p_base_sql, 'avg', p_lfl, p_total),
                        p_sheet          => p_sheet,
                        p_column_headers => FALSE,
                        p_sdvig_row      => p_stat_row + 1,
                        p_sdvig_col      => p_col);
    as_xlsx.query2sheet(p_sql            => f_stat_sql(p_base_sql, 'max', p_lfl, p_total),
                        p_sheet          => p_sheet,
                        p_column_headers => FALSE,
                        p_sdvig_row      => p_stat_row + 2,
                        p_sdvig_col      => p_col);
    as_xlsx.query2sheet(p_sql            => f_stat_sql(p_base_sql, 'sum', p_lfl, p_total),
                        p_sheet          => p_sheet,
                        p_column_headers => FALSE,
                        p_sdvig_row      => p_stat_row + 3,
                        p_sdvig_col      => p_col);
  end write_block;


begin

  execute immediate 'alter session set nls_language=RUSSIAN';
  execute immediate 'alter session set nls_territory=CIS';
  execute immediate 'alter session set NLS_NUMERIC_CHARACTERS=".,"';
  
  as_xlsx.clear_workbook;
  
  -- кэш стилей
  al_center        := as_xlsx.get_alignment(p_horizontal => 'center', p_vertical => 'center');
  al_right_mid     := as_xlsx.get_alignment(p_horizontal => 'right',  p_vertical => 'center', p_wrapText => FALSE);
  al_rotate_mid    := as_xlsx.get_alignment(p_rotate => '90', p_vertical => 'center', p_horizontal => 'center');
  sty_font_hdr     := as_xlsx.get_font(c_font_name, p_fontsize => c_fs_hdr, p_bold => TRUE);
  sty_font_body    := as_xlsx.get_font(c_font_name, p_fontsize => c_fs_body);
  sty_font_daily   := as_xlsx.get_font(c_font_name, p_fontsize => c_fs_body, p_bold => TRUE, p_italic => TRUE);
  sty_font_stat    := as_xlsx.get_font(c_font_cyr,  p_fontsize => c_fs_hdr, p_bold => TRUE);
  sty_numfmt_year  := as_xlsx.get_numFmt(p_format => 'yyyy');
  sty_numfmt_month := as_xlsx.get_numFmt(p_format => 'mmmm');
  sty_font_month   := as_xlsx.get_font(c_font_name, p_fontsize => c_fs_month);

  -- заливки
  fill_d9d9d9      := as_xlsx.get_fill('solid', p_fgRGB => c_fill_total);
  fill_b7dee8      := as_xlsx.get_fill('solid', p_fgRGB => c_fill_default);
  fill_weekend     := as_xlsx.get_fill('solid', p_fgRGB => c_fill_weekend);
  fill_white       := as_xlsx.get_fill('solid', p_fgRGB => c_fill_white);
  fill_min         := as_xlsx.get_fill('solid', p_fgRGB => c_fill_min);
  fill_avg         := as_xlsx.get_fill('solid', p_fgRGB => c_fill_avg);
  fill_max         := as_xlsx.get_fill('solid', p_fgRGB => c_fill_max);
  fill_sum         := as_xlsx.get_fill('solid', p_fgRGB => c_fill_sum);

  as_xlsx.new_sheet(c_sheet_name, (r_max_lm + c_hdr_rows + (i_max - 2) * c_month_h));

  FOR L IN 1 .. 1 LOOP
    --=== ширина столбцов ===
    for c in 1 .. c_cols_total loop
      case
        when c = 1 then w := c_w_col_year;

        when c = 2 then w := c_w_col_month;

        when c = 3 then w := c_w_col_label;

        else w := c_w_data;

      end case;

      as_xlsx.set_column_width(p_col => c, p_sheet => l, p_width => w);
    end loop;


    --=== ШАПКА (лист 1) ===
    if l = 1 then
      open c_facilities;
      loop
        fetch c_facilities into cur;
        exit when c_facilities%NOTFOUND;
        v_blk  := c_blocks(cur.id);
        -- склады ЕКБ: колонки считаются по формуле (арифметическая прогрессия),
        -- остальные блоки: явные значения из конфигурации
        if v_blk.col_lfl then
          c      := f_wh_col_hdr(cur.id);
        else
          c      := v_blk.col_hdr;

        end if;

        -- id заливки текущего блока берется из кэша (см. begin): D9D9D9 у ИТОГО/РЦ ЕКБ,
        -- остальные блоки - B7DEE8
        fillid := case when v_blk.fill = c_fill_total then fill_d9d9d9
                       else fill_b7dee8 end;

        -- объединение ячеек в первой строке шапки
        as_xlsx.mergecells(p_tl_col => c, p_br_col => c + 4 + v_blk.hdr_extra,
                           p_br_row => 1, p_tl_row => 1, p_sheet => l);
        -- загрушки внутри объединенной шапки (рамка без стиля)
        for k in 1 .. 3 loop
          as_xlsx.cell(p_col => c + k, p_row => 1, p_sheet => l,
                       p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
        end loop;

        if v_blk.with_lfl then
          as_xlsx.cell(p_col => c + 8, p_row => 1, p_sheet => l,
                       p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          if v_blk.nz_off = 9 then -- ИТОГО/РЦ ЕКБ: дополнительная загрушка
            as_xlsx.cell(p_col => c + 9, p_row => 1, p_sheet => l,
                         p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          end if;

        end if;

        -- первая строчка шапки - название блока
        as_xlsx.cell(p_col       => c,
                     p_row       => 1,
                     p_sheet     => l,
                     p_value     => v_blk.name,
                     p_alignment => al_center,
                     p_fillId    => fillid,
                     p_fontId    => sty_font_hdr,
                     p_borderId  => as_xlsx.get_border());

        -- третья строчка шапки: годы/LFL/НЗ
        as_xlsx.cell(p_col       => c,
                     p_row       => 3,
                     p_sheet     => l,
                     p_value     => py,
                     p_alignment => al_center,
                     p_fillId    => fillid,
                     p_fontId    => sty_font_hdr,
                     p_borderId  => as_xlsx.get_border());
        as_xlsx.cell(p_col       => c + 1,
                     p_row       => 3,
                     p_sheet     => l,
                     p_value     => cy,
                     p_alignment => al_center,
                     p_fillId    => fillid,
                     p_fontId    => sty_font_hdr,
                     p_borderId  => as_xlsx.get_border());
        if not v_blk.with_lfl then
          -- обычный блок (склады): два цикла лет
          as_xlsx.cell(p_col => c + 2, p_row => 3, p_sheet => l, p_value => py,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 3, p_row => 3, p_sheet => l, p_value => cy,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 4, p_row => 3, p_sheet => l, p_value => py,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 5, p_row => 3, p_sheet => l, p_value => cy,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
        else
          -- ИТОГО/РЦ: LFL + второй цикл лет (+НЗ у ИТОГО/РЦ ЕКБ)
          as_xlsx.cell(p_col => c + 2, p_row => 3, p_sheet => l, p_value => c_lbl_lfl,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 5, p_row => 3, p_sheet => l, p_value => c_lbl_lfl,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 3, p_row => 3, p_sheet => l, p_value => py,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 4, p_row => 3, p_sheet => l, p_value => cy,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 8, p_row => 3, p_sheet => l, p_value => c_lbl_lfl,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 6, p_row => 3, p_sheet => l, p_value => py,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 7, p_row => 3, p_sheet => l, p_value => cy,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          if v_blk.nz_off = 9 then
            as_xlsx.cell(p_col => c + 9, p_row => 3, p_sheet => l, p_value => c_lbl_nz,
                         p_alignment => al_center, p_fillId => fillid,
                         p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          end if;

        end if;


        --=== ВТОРАЯ СТРОКА ШАПКИ ===
        if v_blk.wide_hdr then
          -- три группы по 3 колонки: Паллеты / СС паллета без НДС / Сумма отгрузки
          as_xlsx.mergecells(p_tl_col => c,     p_br_col => c + 2, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 1, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 2, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c,     p_row => 2, p_sheet => l, p_value => c_lbl_pallet,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.mergecells(p_tl_col => c + 3, p_br_col => c + 5, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 4, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 5, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 3, p_row => 2, p_sheet => l, p_value => c_lbl_ss,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.mergecells(p_tl_col => c + 6, p_br_col => c + 8, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 7, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 8, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 6, p_row => 2, p_sheet => l, p_value => c_lbl_summ,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          if v_blk.nz_off = 9 then
            as_xlsx.cell(p_col => c + 9, p_row => 2, p_sheet => l, p_value => c_cell_blank,
                         p_alignment => al_center, p_fillId => fillid,
                         p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border()); -- загрушка над НЗ
          end if;

        else
          -- узкая шапка складов: группы по 2 колонки
          as_xlsx.mergecells(p_tl_col => c,     p_br_col => c + 1, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 1, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c,     p_row => 2, p_sheet => l, p_value => c_lbl_pallet,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.mergecells(p_tl_col => c + 2, p_br_col => c + 3, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 3, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 2, p_row => 2, p_sheet => l, p_value => c_lbl_ss,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.mergecells(p_tl_col => c + 4, p_br_col => c + 5, p_br_row => 2, p_tl_row => 2, p_sheet => l);
          as_xlsx.cell(p_col => c + 5, p_row => 2, p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 4, p_row => 2, p_sheet => l, p_value => c_lbl_summ,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 6, p_row => 2, p_sheet => l, p_value => c_cell_blank,
                       p_alignment => al_center, p_fillId => fillid,
                       p_fontId => sty_font_hdr, p_borderId => as_xlsx.get_border());

        end if;

      end loop;

      close c_facilities;
    end if;


    --=== ЦИКЛ ПО МЕСЯЦАМ ===
    for i in 1 .. i_max loop
      continue;
      select count(*) into sdv_row
        from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
         and to_number(to_char(tran_date, 'yyyy')) = cy
         and tran_date <= max_date;

      sdv_row_global := sdv_row_global + sdv_row;

      -- раскраска строк дней (выходные)
      for r in c_hdr_rows + 1 + c_month_h * (i - 1) + sdv_row_global - sdv_row - err
              .. c_hdr_rows + 1 + c_month_h * (i - 1) + sdv_row_global - 1 - err loop
        select to_number(to_char(tran_date, 'd')) into dw
          from t_dates
         where tran_date = trunc(tran_date, 'mm') + r - c_hdr_rows - 1
                                 - c_month_h * (i - 1) - (sdv_row_global - sdv_row) + err
           and to_number(to_char(tran_date, 'mm')) = i
           and to_number(to_char(tran_date, 'yyyy')) = cy;

        if dw in (c_day_sat, c_day_sun) then
          as_xlsx.set_row(p_sheet => l, p_row => r,
                          p_fillId => fill_weekend, p_fontId => sty_font_daily,
                          p_alignment => al_right_mid, p_borderId => as_xlsx.get_border());
        else
          as_xlsx.set_row(p_sheet => l, p_row => r,
                          p_fontId => sty_font_body,
                          p_alignment => al_right_mid, p_borderId => as_xlsx.get_border());

        end if;

      end loop;


      select tran_date into m_date
        from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
         and to_number(to_char(tran_date, 'yyyy')) = cy
         and rownum = 1;


      -- склейка столбцов года и месяца на высоту блока
      as_xlsx.mergecells(p_tl_col => 1, p_br_col => 1,
                         p_tl_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 - err,
                         p_br_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 4 - err,
                         p_sheet => l);
      as_xlsx.mergecells(p_tl_col => 2, p_br_col => 2,
                         p_tl_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 - err,
                         p_br_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 4 - err,
                         p_sheet => l);
                         
     /*                    
      as_xlsx.mergecells(p_tl_col => 1,p_br_col => 1,p_br_row => d + sdv_row_global + 4 * (i - 1) + 4 - err,p_tl_row => d + sdv_row_global + 4 * (i - 1) + 1 - err,p_sheet  => l);
      as_xlsx.mergecells(p_tl_col => 2,p_br_col => 2,p_br_row => d + sdv_row_global + 4 * (i - 1) + 4 - err,p_tl_row => d + sdv_row_global + 4 * (i - 1) + 1 - err,p_sheet  => l);
*/

      -- год
      as_xlsx.cell(p_col => 1,
                   p_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 - err,
                   p_sheet => l, p_value => m_date,
                   p_fillId => fill_white, p_fontId => sty_font_body,
                   p_numFmtId => sty_numfmt_year, p_borderId => as_xlsx.get_border(),
                   p_alignment => al_center);          
                   
      for k in 1 .. 3 loop
        as_xlsx.cell(p_col => 1, p_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 + k - err,
                    p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
      end loop;

      -- месяц
      as_xlsx.cell(p_col => 2,
                   p_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 - err,
                   p_sheet => l, p_value => m_date,
                   p_fillId => fill_b7dee8, p_fontId => sty_font_month,
                   p_numFmtId => sty_numfmt_month, p_borderId => as_xlsx.get_border(),
                   p_alignment => al_rotate_mid);

      for k in 1 .. 3 loop
        as_xlsx.cell(p_col => 2, p_row => c_hdr_rows + sdv_row_global + c_month_h * (i - 1) + 1 + k - err,
                    p_sheet => l, p_value => c_cell_blank, p_borderId => as_xlsx.get_border());
      end loop;


      -- ДАТЫ
      as_xlsx.query2sheet(p_sql =>
        'select to_char(tran_date, ''Dy'') d, rownum dd, tran_date
           from dwh.t_dates
          where to_number(to_char(tran_date,''mm''))=' || i || '
            and to_number(to_char(tran_date,''yyyy''))=' || cy || '
            and tran_date <= to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'')',
        p_sheet => l, p_column_headers => FALSE,
        p_sdvig_row => c_hdr_rows + c_month_h * (i - 1) + (sdv_row_global - sdv_row) - err);

      if l = 1 then
        --===== ИТОГО =====
        base_sql := f_sql_base(c_blocks(1), i, cy, max_date, null);
        write_block(p_sheet     => l,
                    p_base_sql  => base_sql,
                    p_lfl       => TRUE,
                    p_total     => TRUE,
                    p_first_row => c_hdr_rows + c_month_h * (i - 1) + (sdv_row_global - sdv_row),
                    p_stat_row  => c_hdr_rows + sdv_row_global + c_month_h * (i - 1),
                    p_col       => c_blocks(1).col_data);
        -- подписи строк статистики + их оформление
        as_xlsx.set_row(p_sheet => l, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 1,
                        p_fillId => fill_min, p_fontId => sty_font_stat, p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col => 3, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 1,
                     p_sheet => l, p_value => c_lbl_min, p_fillId => fill_min);
        as_xlsx.set_row(p_sheet => l, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 2,
                        p_fillId => fill_avg, p_fontId => sty_font_stat, p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col => 3, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 2,
                     p_sheet => l, p_value => c_lbl_avg, p_fillId => fill_avg);
        as_xlsx.set_row(p_sheet => l, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 3,
                        p_fillId => fill_max, p_fontId => sty_font_stat, p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col => 3, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 3,
                     p_sheet => l, p_value => c_lbl_max, p_fillId => fill_max);
        as_xlsx.set_row(p_sheet => l, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 4,
                        p_fillId => fill_sum, p_fontId => sty_font_stat, p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col => 3, p_row => c_hdr_rows + sdv_row_global + c_month_h*(i-1) + 4,
                     p_sheet => l, p_value => c_lbl_sum, p_fillId => fill_sum);

        --===== остальные блоки (РЦ ЕКБ + склады + РЦ) =====
        open c_facilities;
        loop
          fetch c_facilities into cur;
          exit when c_facilities%NOTFOUND;
          v_blk := c_blocks(cur.id);
          -- колонка данных: склады ЕКБ - по формуле (7,8 в оригинале не выводились),
          -- остальные - из конфигурации; NULL => блок данных нет
          if v_blk.col_lfl then
            if cur.id <= c_last_wh_id - 2 then -- склады с данными: id 3..6
              v_col := f_wh_col_data(cur.id);
            else
              v_col := null;                   -- id 7,8: как sdv_cols=-1 в оригинале

            end if;

          else
            v_col := v_blk.col_data;

          end if;

          if v_col is not null then
            base_sql := f_sql_base(v_blk, i, cy, max_date, cur.t_name);
            write_block(p_sheet     => l,
                        p_base_sql  => base_sql,
                        p_lfl       => v_blk.with_lfl,
                        p_total     => FALSE,
                        p_first_row => c_hdr_rows + c_month_h * (i - 1) + (sdv_row_global - sdv_row),
                        p_stat_row  => c_hdr_rows + sdv_row_global + c_month_h * (i - 1),
                        p_col       => v_col);
          end if;

        end loop;

        close c_facilities;
      end if; -- l = 1

    end loop; -- месяцы
    sdv_row_global := 0;
  end loop; -- листы

  as_xlsx.save(c_out_dir, c_out_file);
/*
exception
  when others then
    if c_facilities%IsOPEN then
      close c_facilities;
    end if;
    dwh.cognos_utl.log_err(err_code => sqlcode,
                           err_msg  => sqlerrm,
                           proc     => c_proc);
    raise;*/

end create_sklod_report_new;

--<< PROG-4861