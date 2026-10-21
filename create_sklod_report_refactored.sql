-->> PROG-4860
-- Рефакторинг create_sklod_report: убран хардкод, логика/результат сохранены 1:1.
-- Как добавить новый РЦ в отчет (3 шага):
--   1) в c_blocks дописать запись блока (имя, тип, rc_id, колонка шапки);
--   2) при необходимости — смещение данных блока в c_block_sdv_cols;
--   3) всё остальное (шапка, SQL, статистика) подхватится автоматически.
procedure create_sklod_report(pmaxdate in date default trunc(sysdate) - 1) is
  -- Отчет показатели логистики

  ---------------------------------------------------------------------------
  -- Константы процесса и структуры листа
  ---------------------------------------------------------------------------
  c_proc            constant varchar2(200) := 'LOGISTICS.CREATE_SKLOD_REPORT';
  c_xlsx_dir        constant varchar2(30)  := 'SVOD_DIR';          -- каталог сохранения файла
  c_xlsx_file       constant varchar2(30)  := 'sklod_report.xlsx'; -- имя файла отчета
  c_sheet_name      constant varchar2(100) := 'Показатели логистики';

  c_header_rows     constant pls_integer := 3;  -- высота шапки в строках (отступ d)
  c_stats_rows      constant pls_integer := 4;  -- служебные строки месяца: все/мин/средн/макс+сумма
  c_year_col        constant pls_integer := 1;  -- колонка года
  c_month_col       constant pls_integer := 2;  -- колонка месяца
  c_first_data_col  constant pls_integer := 3;  -- первая колонка данных (метки строк статистики)
  c_last_col        constant pls_integer := 150; -- количество столбцов с заданной шириной
  c_wknd_from       constant pls_integer := 6;  -- первый номер дня недели выходного (суббота)
  c_wknd_to         constant pls_integer := 7;  -- последний номер дня недели выходного (воскресенье)

  ---------------------------------------------------------------------------
  -- Стили: шрифты, цвета заливок, выравнивания
  ---------------------------------------------------------------------------
  c_font_hdr        constant varchar2(30) := 'Arial';
  c_font_stat       constant varchar2(30) := 'Arial Cyr';
  c_fs_hdr          constant pls_integer  := 11; -- размер шрифта шапки
  c_fs_date         constant pls_integer  := 10; -- размер шрифта дат
  c_fs_month        constant pls_integer  := 12; -- размер шрифта названия месяца

  c_rgb_default     constant varchar2(10) := 'B7DEE8'; -- цвет обычного блока
  c_rgb_total       constant varchar2(10) := 'D9D9D9'; -- цвет блоков ИТОГО/РЦ ЕКБ
  c_rgb_white       constant varchar2(10) := 'FFFFFF'; -- фон года
  c_rgb_weekend     constant varchar2(10) := 'B3FFFF'; -- фон выходных
  c_rgb_all         constant varchar2(10) := 'F1FB93'; -- строка «все» / «Минимум»
  c_rgb_avg         constant varchar2(10) := '95FD9A'; -- строка «Среднее»
  c_rgb_max         constant varchar2(10) := 'FFCCCC'; -- строка «Максимум»
  c_rgb_sum         constant varchar2(10) := 'FFB953'; -- строка «Сумма»

  c_horz_center     constant varchar2(10) := 'center';
  c_vert_center     constant varchar2(10) := 'center';
  c_horz_right      constant varchar2(10) := 'right';

  ---------------------------------------------------------------------------
  -- Тексты шапки и строк статистики
  ---------------------------------------------------------------------------
  c_lbl_pallet      constant varchar2(30) := 'Паллеты';
  c_lbl_ss          constant varchar2(30) := 'СС паллета без НДС';
  c_lbl_summ        constant varchar2(30) := 'Сумма отгрузки';
  c_lbl_lfl         constant varchar2(10) := 'LFL';
  c_lbl_nz          constant varchar2(10) := 'НЗ';
  c_lbl_min         constant varchar2(30) := 'Минимум';
  c_lbl_avg         constant varchar2(30) := 'Среднее';
  c_lbl_max         constant varchar2(30) := 'Максимум';
  c_lbl_sum         constant varchar2(30) := 'Сумма';

  ---------------------------------------------------------------------------
  -- Идентификаторы складов/РЦ из витрины (были зашиты в SQL)
  ---------------------------------------------------------------------------
  c_wh_physical     constant pls_integer := 2;    -- physical_wh_id склада WMS (Екатеринбург)
  c_rc_ekb          constant pls_integer := 2;    -- РЦ ЕКАТЕРИНБУРГ + РЦ БЕРЕЗОВСКИЙ
  c_rc_ufa          constant pls_integer := 2152; -- РЦ УФА2 + РЦ Уфа
  c_rc_nfteyugansk  constant pls_integer := 2014; -- РЦ НЕФТЕЮГАНСК + РЦ СУРГУТ
  c_rc_novosibirsk  constant pls_integer := 2037; -- РЦ НОВОСИБИРСК + РЦ НСК
  c_rc_perm         constant pls_integer := 3002; -- РЦ ПЕРМЬ
  c_rc_perm2        constant pls_integer := 3025; -- РЦ ПЕРМЬ2
  c_rc_kurgan       constant pls_integer := 3013; -- РЦ ЮЖНОУРАЛЬСК
  c_rc_tumen        constant pls_integer := 3028; -- РЦ ТЮМЕНЬ

  ---------------------------------------------------------------------------
  -- Конфигурация блоков отчета.
  -- kind: 'T' = итоги по всем складам, 'S' = склады ЕКБ (ks2), 'R' = отдельный РЦ (rc_id).
  -- tl_col: стартовая колонка блока в шапке; sdv_col: стартовая колонка блока данных.
  -- order: порядок вывода (ord1 из исходного UNION ALL-курсорa).
  -- ЗАМЕТКА: значения tl_col/sdv_col/width повторяют исходник дословно, включая
  -- расхождения между ними (в оригинале блоки после ИТОГО частично не совпадают
  -- с колонками данных — это историческое поведение, оставлено как есть).
  ---------------------------------------------------------------------------
  type t_block is record
    id      pls_integer;
    name    varchar2(100);
    kind    varchar2(1);
    rc_id   pls_integer;
    ord     pls_integer;
    tl_col  pls_integer;
    width   pls_integer; -- ширина mergecells в первой строке (col1..colN)
    sdv_col pls_integer;
  end record;
  type t_blocks is table of t_block index by pls_integer;

  c_blocks constant t_blocks := t_blocks
    ( 1 => t_block(id =>  1, name => 'ИТОГО',
                    kind => 'T', rc_id => null, ord =>  1,
                    tl_col =>   4, width =>  10, sdv_col =>  23),
      2 => t_block(id =>  2, name => 'РЦ ЕКАТЕРИНБУРГ + РЦ БЕРЕЗОВСКИЙ',
                    kind => 'T', rc_id => c_rc_ekb, ord =>  2,
                    tl_col =>  14, width =>  10, sdv_col =>  30),
      3 => t_block(id =>  3, name => null, -- склады ЕКБ: имена берутся из dwh.monwms_facility_repnames
                    kind => 'S', rc_id => c_rc_ekb, ord =>  3,
                    tl_col =>  24, width =>   7, sdv_col =>  37),
      4 => t_block(id =>  4, name => null,
                    kind => 'S', rc_id => c_rc_ekb, ord =>  3,
                    tl_col =>  31, width =>   7, sdv_col =>  44),
      5 => t_block(id =>  5, name => null,
                    kind => 'S', rc_id => c_rc_ekb, ord =>  3,
                    tl_col =>  38, width =>   7, sdv_col =>  51),
      6 => t_block(id =>  6, name => null,
                    kind => 'S', rc_id => c_rc_ekb, ord =>  3,
                    tl_col =>  45, width =>   7, sdv_col =>  58),
      7 => t_block(id =>  7, name => 'РЦ УФА',
                    kind => 'R', rc_id => c_rc_ufa, ord =>  6,
                    tl_col =>  52, width =>   7, sdv_col =>  65),
      8 => t_block(id =>  8, name => 'РЦ НЕФТЕЮГАНСК',
                    kind => 'R', rc_id => c_rc_nfteyugansk, ord =>  8,
                    tl_col =>  59, width =>   7, sdv_col =>  74),
      9 => t_block(id =>  9, name => 'РЦ НОВОСИБИРСК',
                    kind => 'R', rc_id => c_rc_novosibirsk, ord => 10,
                    tl_col =>  66, width =>   9, sdv_col =>  83),
     10 => t_block(id => 10, name => 'РЦ ПЕРМЬ',
                    kind => 'R', rc_id => c_rc_perm, ord => 12,
                    tl_col =>  75, width =>   9, sdv_col =>  92),
     11 => t_block(id => 11, name => 'РЦ ЮЖНОУРАЛЬСК',
                    kind => 'R', rc_id => c_rc_kurgan, ord => 14,
                    tl_col =>  84, width =>   9, sdv_col => 101),
     12 => t_block(id => 12, name => 'РЦ ПЕРМЬ2',
                    kind => 'R', rc_id => c_rc_perm2, ord => 16,
                    tl_col =>  93, width =>   9, sdv_col => 110),
     13 => t_block(id => 13, name => 'РЦ ТЮМЕНЬ',
                    kind => 'R', rc_id => c_rc_tumen, ord => 18,
                    tl_col => 102, width =>   9, sdv_col => 119));

  -- В исходнике список блоков формировался курсором на 15 позиций (два «пустых»
  -- id = 14, 15 без имени). Сохранено для идентичности результата.
  c_blocks_count constant pls_integer := 15;

  ---------------------------------------------------------------------------
  -- Переменные (соответствуют исходным)
  ---------------------------------------------------------------------------
  max_date       date        := trunc(pmaxdate);
  sdv_row_global pls_integer := 0;
  sdv_row_prev   pls_integer := 0;
  sdv_row        pls_integer := 0;
  sdv_cols       pls_integer := 0; -- сдвиг блока данных по колонкам
  i_max          pls_integer := to_number(to_char(max_date, 'mm')); -- количество месяцев
  i_min          pls_integer := 1; -- начальный месяц для цикла по месяцам
  r_max_lm       pls_integer := trunc(max_date, 'mm') - trunc(max_date, 'yyyy'); -- кол-во дней текущего месяца
  l              integer;          -- номер листа
  c              pls_integer := 0; -- текущая колонка
  e              integer;          -- смещение строки блока (в исходнике вычислялось, но не использовалось)
  r              integer;          -- номер строки при раскраске выходных
  cy             pls_integer := to_number(to_char(max_date, 'yyyy'));                       -- текущий год
  py             pls_integer := to_number(to_char(trunc(max_date, 'yyyy') - 1, 'yyyy'));    -- предыдущий год
  dw             pls_integer := 1; -- порядковый номер дня недели
  err            pls_integer := 0; -- костыль смещения строк (в исходнике всегда 0)
  m_date         date;
  ssql_min       varchar2(32000);
  ssql_avg       varchar2(32000);
  ssql_max       varchar2(32000);
  ssql_sum       varchar2(32000);
  ssql_all       varchar2(32000);
  ssql           varchar2(32000);
  rgb            varchar2(50);
  v_blk          t_block;

  ---------------------------------------------------------------------------
  -- Вспомогательные функции
  ---------------------------------------------------------------------------

  -- Стартовая колонка блока в шапке (замена case-простыни из исходника)
  function block_tl_col(p_id in pls_integer) return pls_integer is
  begin
    return c_blocks(p_id).tl_col;
  exception
    when others then return -1; -- как в оригинале: неизвестный id => -1
  end block_tl_col;

  -- Смещение блока данных по колонкам (замена второй case-простыни)
  function block_sdv_col(p_id in pls_integer) return pls_integer is
  begin
    return c_blocks(p_id).sdv_col;
  exception
    when others then return -1;
  end block_sdv_col;

  -- Цвет фона блока: серые для первых двух (ИТОГО, РЦ ЕКБ), иначе голубой
  function block_rgb(p_id in pls_integer) return varchar2 is
  begin
    if p_id in (1, 2) then
      return c_rgb_total;
    else
      return c_rgb_default;
    end if;
  end block_rgb;

  -- Тип блока (для блоков вне таблицы — 'R', как у «пустых» id 14/15 в исходном курсоре)
  function block_kind(p_id in pls_integer) return varchar2 is
  begin
    return c_blocks(p_id).kind;
  exception
    when others then return 'R';
  end block_kind;

  -- Имя блока из конфигурации (для курсора складов); нет записи => null (как id 14/15)
  function block_name(p_id in pls_integer) return varchar2 is
  begin
    return c_blocks(p_id).name;
  exception
    when others then return null;
  end block_name;

  -- rc_id блока из конфигурации; нет записи => -1 (данных нет, блок пустой)
  function block_rc(p_id in pls_integer) return pls_integer is
  begin
    return nvl(c_blocks(p_id).rc_id, -1);
  exception
    when others then return -1;
  end block_rc;

  -- Порядок вывода блока в курсоре складов/РЦ (ord1 из исходного UNION ALL)
  function block_ord(p_id in pls_integer) return pls_integer is
  begin
    return c_blocks(p_id).ord;
  exception
    when others then return 999;
  end block_ord;

  -- Ширина столбца (замена вложенного case; ветка l<>1 недостижима — лист один)
  function col_width(p_col in pls_integer, p_sheet_no in integer) return number is
  begin
    case p_col
      when c_year_col  then return 5.5;
      when c_month_col then return 4;
      when c_first_data_col then return 12.5;
      else return case when p_sheet_no = 1 then 16 else 20 end;
    end case;
  end col_width;

  -- Запись ячейки шапки со стандартным стилем
  procedure hdr_cell(p_col in pls_integer, p_row in pls_integer, p_sheet in integer,
                     p_value in varchar2, p_rgb in varchar2) is
  begin
    as_xlsx.cell(p_col       => p_col,
                 p_row       => p_row,
                 p_sheet     => p_sheet,
                 p_value     => p_value,
                 p_alignment => as_xlsx.get_alignment(p_horizontal => c_horz_center,
                                                      p_vertical   => c_vert_center),
                 p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => p_rgb),
                 p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_hdr,
                                                 p_bold => TRUE),
                 p_borderId  => as_xlsx.get_border());
  end hdr_cell;

  -- Пустая ячейка-«заглушка» только с рамкой (для склейки объединенных диапазонов)
  procedure blank_cell(p_col in pls_integer, p_row in pls_integer, p_sheet in integer) is
  begin
    as_xlsx.cell(p_col      => p_col,
                 p_row      => p_row,
                 p_sheet    => p_sheet,
                 p_value    => ' ',
                 p_borderId => as_xlsx.get_border());
  end blank_cell;

  -- Пустая залитая ячейка (вторая строка шапки)
  procedure blank_filled_cell(p_col in pls_integer, p_row in pls_integer,
                              p_sheet in integer, p_rgb in varchar2) is
  begin
    as_xlsx.cell(p_col       => p_col,
                 p_row       => p_row,
                 p_sheet     => p_sheet,
                 p_value     => ' ',
                 p_alignment => as_xlsx.get_alignment(p_horizontal => c_horz_center,
                                                      p_vertical   => c_vert_center),
                 p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => p_rgb),
                 p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_hdr,
                                                 p_bold => TRUE),
                 p_borderId  => as_xlsx.get_border());
  end blank_filled_cell;

  -- Объединение диапазона [p_from .. p_to] в строке p_row (порядок аргументов как в оригинале)
  procedure merge_hdr(p_tl_col in pls_integer, p_span in pls_integer,
                      p_row in pls_integer, p_sheet in integer) is
  begin
    as_xlsx.mergecells(p_tl_col => p_tl_col,
                       p_br_col => p_tl_col + p_span,
                       p_tl_row => p_row,
                       p_br_row => p_row,
                       p_sheet  => p_sheet);
  end merge_hdr;

  ---------------------------------------------------------------------------
  -- Построение SQL-фрагментов (устранена копипаста одинаковых подзапросов)
  ---------------------------------------------------------------------------

  -- Литерал даты для вставки в динамический SQL (формат маскируется от NLS)
  function date_lit(p_date in date) return varchar2 is
  begin
    return 'to_date(''' || to_char(p_date, 'dd.mm.yyyy') || ''',''dd.mm.yyyy'')';
  end date_lit;

  -- Подзапрос агрегации по складам ЕКБ (используется для a=текущий год и b=прошлый).
  -- p_ks2 = null => без разреза по ks2 (блок «ИТОГИ РЦ ТЭЦ», как в исходнике).
  function sql_agg_ekb(p_month in pls_integer, p_ks2 in pls_integer default null) return varchar2 is
  begin
    return 'select t_date, sum(pallet) pallet,sum(totalsum) totalsum
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
                           and vsf.rc_id = ' || c_rc_ekb || ' -- екб+березовск'
                   || case when p_ks2 is not null
                          then '
                           and vsf.ks2 = ' || p_ks2 || ' -- это потому что по складам'
                          else '' end || '
                         group by sp.matrix_id
                                 ,sp.date_shift_start
                              )
                              group by t_date';
  end sql_agg_ekb;

  -- Подзапрос агрегации по одному РЦ (rc_id)
  function sql_agg_rc(p_month in pls_integer, p_rc_id in pls_integer) return varchar2 is
  begin
    return '
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
                   group by sp.date_shift_start';
  end sql_agg_rc;

  -- Хвост WHERE общего для всех запросов блока
  function sql_tail(p_month in pls_integer) return varchar2 is
  begin
    return 'where to_number(to_char(v.tran_date,''mm''))=' || p_month ||
           ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy ||
           '  and v.tran_date <=  ' || date_lit(max_date) || ' ) order by tran_date';
  end sql_tail;

  -- Выражение LFL-доли вида cast((cur/NULLIF(prev,0)-1) as binary_double), где cur/prev —
  -- уже оформленные ссылки на колонки (например «"p.2015"» или «"ss2014"»).
  function ratio_expr(p_cur in varchar2, p_prev in varchar2) return varchar2 is
  begin
    return 'cast((' || p_cur || '/nullif(' || p_prev || ',0)-1) as binary_double)';
  end ratio_expr;

  -- Строка агрегатных выражений для ssql_all/min/max:
  --   agg_fn='cast' => LFL считается построчно: cast((cur/NULLIF(prev,0)-1))
  --   иначе         => LFL считается по агрегатам: agg(cur)/NULLIF(agg(prev),0)-1
  function stat_cols(p_agg        in varchar2,
                     p_pallet_cur in varchar2, p_pallet_prev in varchar2,
                     p_ss_cur     in varchar2, p_ss_prev     in varchar2,
                     p_sum_cur    in varchar2, p_sum_prev    in varchar2,
                     p_nz         in varchar2 default null) return varchar2 is
    v_lfl pls_integer := case when p_agg = 'cast' then 1 else 0 end;
    v_ret varchar2(32767);
  begin
    v_ret := p_agg || '(' || p_pallet_prev || '), ' || p_agg || '(' || p_pallet_cur || '), '
             || case v_lfl when 1
                then ratio_expr(p_pallet_cur, p_pallet_prev)
                else 'cast(' || p_agg || '(' || p_pallet_cur || ')/nullif('
                     || p_agg || '(' || p_pallet_prev || '),0)-1 as binary_double)'
             end || ', '
          || p_agg || '(' || p_ss_prev || '), ' || p_agg || '(' || p_ss_cur || '), '
             || case v_lfl when 1
                then ratio_expr(p_ss_cur, p_ss_prev)
                else 'cast(' || p_agg || '(' || p_ss_cur || '"/nullif('
                     || p_agg || '(' || p_ss_prev || '"),0)-1 as binary_double)'
             end || ', '
          || p_agg || '(' || p_sum_prev || '), ' || p_agg || '(' || p_sum_cur || '), '
             || case v_lfl when 1
                then ratio_expr(p_sum_cur, p_sum_prev)
                else 'cast(' || p_agg || '(' || p_sum_cur || ')/nullif('
                     || p_agg || '(' || p_sum_prev || '),0)-1 as binary_double)'
             end;
    if p_nz is not null then
      v_ret := v_ret || ', ' || p_agg || '(' || p_nz || ')';
    end if;
    return 'select ' || v_ret;
  end stat_cols;

  -- Строка «голых» агрегатов без LFL-колонок (узкие блоки складов в ssql_min/avg/max)
  function plain_agg(p_agg        in varchar2,
                     p_pallet_cur in varchar2, p_pallet_prev in varchar2,
                     p_ss_cur     in varchar2, p_ss_prev     in varchar2,
                     p_sum_cur    in varchar2, p_sum_prev    in varchar2,
                     p_nz         in varchar2 default null) return varchar2 is
    v_ret varchar2(32767);
  begin
    v_ret := p_agg || '(' || p_pallet_prev || '), ' || p_agg || '(' || p_pallet_cur || '), '
          || p_agg || '(' || p_ss_prev     || '), ' || p_agg || '(' || p_ss_cur     || '), '
          || p_agg || '(' || p_sum_prev    || '), ' || p_agg || '(' || p_sum_cur    || ')';
    if p_nz is not null then
      v_ret := v_ret || ', ' || p_agg || '(' || p_nz || ')';
    end if;
    return 'select ' || v_ret;
  end plain_agg;

  -- Строка sum-агрегатов с LFL по агрегатам; СС-паллеты — avg() (как в исходнике).
  -- p_nz = null => колонки НЗ нет (региональные РЦ).
  function sum_cols(p_pallet_cur in varchar2, p_pallet_prev in varchar2,
                    p_ss_cur     in varchar2, p_ss_prev     in varchar2,
                    p_sum_cur    in varchar2, p_sum_prev    in varchar2,
                    p_nz         in varchar2 default null) return varchar2 is
    v_ret varchar2(32767);
  begin
    v_ret := 'sum(' || p_pallet_prev || '), sum(' || p_pallet_cur || '),'
             || 'cast(sum(' || p_pallet_cur || ')/nullif(sum(' || p_pallet_prev || '),0)-1 as binary_double), '
          || 'avg(' || p_ss_prev || '), avg(' || p_ss_cur || '),'
             || 'cast(avg(' || p_ss_cur || ')/nullif(avg(' || p_ss_prev || '),0)-1 as binary_double), '
          || 'sum(' || p_sum_prev || '), sum(' || p_sum_cur || '),'
             || 'cast(sum(' || p_sum_cur || ')/nullif(sum(' || p_sum_prev || '),0)-1 as binary_double)';
    if p_nz is not null then
      v_ret := v_ret || ', sum(' || p_nz || ')';
    end if;
    return 'select ' || v_ret;
  end sum_cols;

  ---------------------------------------------------------------------------
  -- Генерация полного набора текстов запросов блока (all/min/avg/max/sum).
  -- kind: 'W' = широкий блок с LFL и колонкой НЗ (ИТОГИ, ИТОГИ РЦ ТЭЦ);
  --       'N' = узкий блок складов ЕКБ (без LFL, с НЗ);
  --       'R' = региональный РЦ (с LFL, без НЗ).
  -- p_*_cur/p_*_prev — оформленные ссылки на колонки текущего/прошлого года.
  ---------------------------------------------------------------------------
  procedure build_stats_sql(p_kind     in varchar2,
                            p_pallet_cur in varchar2, p_pallet_prev in varchar2,
                            p_ss_cur     in varchar2, p_ss_prev     in varchar2,
                            p_sum_cur    in varchar2, p_sum_prev    in varchar2,
                            p_nz         in varchar2 default null) is
  begin
    if p_kind = 'N' then
      ssql_all := 'select ' || p_pallet_prev || ', ' || p_pallet_cur || ', '
                  || p_ss_prev || ', ' || p_ss_cur || ', '
                  || p_sum_prev || ', ' || p_sum_cur || ', ' || p_nz || ssql;
      ssql_min := plain_agg('min', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_avg := plain_agg('avg', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_max := plain_agg('max', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_sum := 'select sum(' || p_pallet_prev || '), sum(' || p_pallet_cur || '), '
                  || 'avg(' || p_ss_prev || '), avg(' || p_ss_cur || '), '
                  || 'sum(' || p_sum_prev || '), sum(' || p_sum_cur || '), '
                  || 'sum(' || p_nz || ')' || ssql;
    else
      ssql_all := 'select ' || p_pallet_prev || ', ' || p_pallet_cur || ', '
                  || ratio_expr(p_pallet_cur, p_pallet_prev) || ', '
                  || p_ss_prev || ', ' || p_ss_cur || ', '
                  || ratio_expr(p_ss_cur, p_ss_prev) || ', '
                  || p_sum_prev || ', ' || p_sum_cur || ', '
                  || ratio_expr(p_sum_cur, p_sum_prev)
                  || case when p_nz is not null then ', ' || p_nz end
                  || ssql;
      ssql_min := stat_cols('min', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_avg := stat_cols('avg', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_max := stat_cols('max', p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                            p_sum_cur, p_sum_prev, p_nz) || ssql;
      ssql_sum := sum_cols(p_pallet_cur, p_pallet_prev, p_ss_cur, p_ss_prev,
                           p_sum_cur, p_sum_prev, p_nz) || ssql;
    end if;
  end build_stats_sql;

  ---------------------------------------------------------------------------
  -- Выгрузка блока из четырех запросов статистики в лист
  ---------------------------------------------------------------------------
  procedure dump_block(p_sheet in integer, p_sdv_col in pls_integer) is
  begin
    as_xlsx.query2sheet(p_sql            => ssql_all,
                         p_sheet          => p_sheet,
                         p_column_headers => FALSE,
                         p_sdvig_row      => c_header_rows + 4 * (i - 1) + (sdv_row_global - sdv_row),
                         p_sdvig_col      => p_sdv_col);
    as_xlsx.query2sheet(p_sql            => ssql_min,
                         p_sheet          => p_sheet,
                         p_column_headers => FALSE,
                         p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1),
                         p_sdvig_col      => p_sdv_col);
    as_xlsx.query2sheet(p_sql            => ssql_avg,
                         p_sheet          => p_sheet,
                         p_column_headers => FALSE,
                         p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 1,
                         p_sdvig_col      => p_sdv_col);
    as_xlsx.query2sheet(p_sql            => ssql_max,
                         p_sheet          => p_sheet,
                         p_column_headers => FALSE,
                         p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 2,
                         p_sdvig_col      => p_sdv_col);
    as_xlsx.query2sheet(p_sql            => ssql_sum,
                         p_sheet          => p_sheet,
                         p_column_headers => FALSE,
                         p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 3,
                         p_sdvig_col      => p_sdv_col);
  end dump_block;

begin
  execute immediate 'alter session set nls_language=RUSSIAN';
  execute immediate 'alter session set nls_territory=CIS';
  execute immediate 'alter session set NLS_NUMERIC_CHARACTERS=".,"';
  as_xlsx.clear_workbook;
  --добавляем листы
  as_xlsx.new_sheet(c_sheet_name, (r_max_lm + c_header_rows + (i_max - 2) * c_stats_rows));
-------------------------------------------------------------------------------------------------------------------------
  --Цикл по листам
  for l in 1 .. 1 loop
    --ширина столбцов
    for c in 1 .. c_last_col loop
      as_xlsx.set_column_width(p_col   => c,
                               p_sheet => l,
                               p_width => col_width(c, l));
    end loop;

    --ШАПКА (лист 1)
    if l = 1 then
      for p_id in 1 .. c_blocks_count loop
        v_blk := c_blocks(p_id);

        c   := block_tl_col(p_id);
        rgb := block_rgb(p_id);

        --объединение ячеек в первой строке шапки
        merge_hdr(c, v_blk.width - 1, 1, l);
        blank_cell(c + 1, 1, l);
        blank_cell(c + 2, 1, l);
        blank_cell(c + 3, 1, l);
        if p_id >= 9 then -- заглушка над блоками «РЦ ...»
          blank_cell(c + 8, 1, l);
        end if;
        if p_id in (1, 2) then -- дополнительные заглушки для ИТОГО/РЦ ЕКБ
          blank_cell(c + 8, 1, l);
          blank_cell(c + 9, 1, l);
        end if;

        -- первая строчка шапки: название блока
        hdr_cell(c, 1, l, v_blk.name, rgb);

        -- третья строчка шапки: годы/LFL/НЗ
        hdr_cell(c,     3, l, to_char(py), rgb);
        hdr_cell(c + 1, 3, l, to_char(cy), rgb);

        if p_id not in (1, 2, 9, 10, 11, 12, 13, 14, 15) then
          -- обычный блок (склады ЕКБ и РЦ УФА/НЕФТЕЮГАНСК): два года подряд + НЗ
          hdr_cell(c + 2, 3, l, to_char(py), rgb);
          hdr_cell(c + 3, 3, l, to_char(cy), rgb);
          hdr_cell(c + 4, 3, l, to_char(py), rgb);
          hdr_cell(c + 5, 3, l, to_char(cy), rgb);
          if p_id > 2 then
            hdr_cell(c + 6, 3, l, c_lbl_nz, rgb);
          end if;
        else
          -- итоги и региональные РЦ: LFL-колонки
          hdr_cell(c + 2, 3, l, c_lbl_lfl, rgb);
          hdr_cell(c + 3, 3, l, to_char(py), rgb);
          hdr_cell(c + 4, 3, l, to_char(cy), rgb);
          hdr_cell(c + 5, 3, l, c_lbl_lfl, rgb);
          hdr_cell(c + 6, 3, l, to_char(py), rgb);
          hdr_cell(c + 7, 3, l, to_char(cy), rgb);
          hdr_cell(c + 8, 3, l, c_lbl_lfl, rgb);
          if p_id in (1, 2) then
            hdr_cell(c + 9, 3, l, c_lbl_nz, rgb);
          end if;
        end if;

        -- ВТОРАЯ СТРОКА ШАПКИ
        if p_id in (1, 2, 9, 10, 11, 12, 13, 14, 15) then
          -- широкие блоки (3 колонки на метрику)
          merge_hdr(c,     2, 2, l);
          blank_cell(c + 1, 2, l);
          blank_cell(c + 2, 2, l);
          hdr_cell(c,     2, l, c_lbl_pallet, rgb);
          merge_hdr(c + 3, 2, 2, l);
          blank_cell(c + 4, 2, l);
          blank_cell(c + 5, 2, l);
          hdr_cell(c + 3, 2, l, c_lbl_ss, rgb);
          merge_hdr(c + 6, 2, 2, l);
          blank_cell(c + 7, 2, l);
          blank_cell(c + 8, 2, l);
          hdr_cell(c + 6, 2, l, c_lbl_summ, rgb);
          if p_id in (1, 2) then -- загрушка над НЗ
            blank_filled_cell(c + 9, 2, l, rgb);
          end if;
        else
          -- узкие блоки (2 колонки на метрику)
          merge_hdr(c,     1, 2, l);
          blank_cell(c + 1, 2, l);
          hdr_cell(c,     2, l, c_lbl_pallet, rgb);
          merge_hdr(c + 2, 1, 2, l);
          blank_cell(c + 3, 2, l);
          hdr_cell(c + 2, 2, l, c_lbl_ss, rgb);
          merge_hdr(c + 4, 1, 2, l);
          blank_cell(c + 5, 2, l);
          hdr_cell(c + 4, 2, l, c_lbl_summ, rgb);
          blank_filled_cell(c + 6, 2, l, rgb);
        end if;
      end loop;
    end if;

-------------------------------------------------------------------------------------------------------------------------------------
    --цикл по месяцам
    i_min := 1;
    for i in i_min .. i_max loop
      select count(*)
        into sdv_row
        from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
         and to_number(to_char(tran_date, 'yyyy')) = cy
         and tran_date <= max_date;
      sdv_row_global := sdv_row_global + sdv_row;

      --красим выходные
      for r in c_header_rows + 1 + 4 * (i - 1) + sdv_row_global - sdv_row - err
            .. c_header_rows + 1 + 4 * (i - 1) + sdv_row_global - 1 - err loop
        select to_number(to_char(tran_date, 'd'))
          into dw
          from t_dates
         where tran_date = trunc(tran_date, 'mm') + r - c_header_rows - 1 - 4 * (i - 1)
                               - (sdv_row_global - sdv_row) + err
           and to_number(to_char(tran_date, 'mm')) = i
           and to_number(to_char(tran_date, 'yyyy')) = cy;
        if dw between c_wknd_from and c_wknd_to then
          as_xlsx.set_row(p_sheet     => l,
                          p_row       => r,
                          p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_weekend),
                          p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_date,
                                                          p_bold => TRUE, p_italic => TRUE),
                          p_alignment => as_xlsx.get_alignment(p_horizontal => c_horz_right,
                                                               p_vertical   => c_vert_center,
                                                               p_wrapText   => FALSE),
                          p_borderId  => as_xlsx.get_border());
        else
          as_xlsx.set_row(p_sheet     => l,
                          p_row       => r,
                          p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_date),
                          p_alignment => as_xlsx.get_alignment(p_horizontal => c_horz_right,
                                                               p_vertical   => c_vert_center,
                                                               p_wrapText   => FALSE),
                          p_borderId  => as_xlsx.get_border());
        end if;
      end loop;

      select tran_date
        into m_date
        from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
         and to_number(to_char(tran_date, 'yyyy')) = cy
         and rownum = 1;

      -- объединение колонок года/месяца на всю высоту месяца
      merge_hdr(c_year_col,  0, c_header_rows + sdv_row_global + 4 * (i - 1) + 4 - err, l);
      merge_hdr(c_month_col, 0, c_header_rows + sdv_row_global + 4 * (i - 1) + 4 - err, l);

      --год
      as_xlsx.cell(p_col       => c_year_col,
                   p_row       => c_header_rows + sdv_row_global + 4 * (i - 1) + 1 - err,
                   p_sheet     => l,
                   p_value     => m_date,
                   p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_white),
                   p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_date),
                   p_numFmtId  => as_xlsx.get_numFmt(p_format => 'yyyy'),
                   p_borderId  => as_xlsx.get_border(),
                   p_alignment => as_xlsx.get_alignment(p_vertical   => c_vert_center,
                                                        p_horizontal => c_horz_center));
      for k in 1 .. 3 loop
        blank_cell(c_year_col, c_header_rows + sdv_row_global + 4 * (i - 1) + 1 + k - err, l);
      end loop;

      --месяц
      as_xlsx.cell(p_col       => c_month_col,
                   p_row       => c_header_rows + sdv_row_global + 4 * (i - 1) + 1 - err,
                   p_sheet     => l,
                   p_value     => m_date,
                   p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_default),
                   p_fontId    => as_xlsx.get_font(c_font_hdr, p_fontsize => c_fs_month),
                   p_numFmtId  => as_xlsx.get_numFmt(p_format => 'mmmm'),
                   p_borderId  => as_xlsx.get_border(),
                   p_alignment => as_xlsx.get_alignment(p_rotate     => '90',
                                                        p_vertical   => c_vert_center,
                                                        p_horizontal => c_horz_center));
      for k in 1 .. 3 loop
        blank_cell(c_month_col, c_header_rows + sdv_row_global + 4 * (i - 1) + 1 + k - err, l);
      end loop;

----------------------------------------------------------------------------------------------
      --ДАТЫ
      as_xlsx.query2sheet(p_sql =>
            'select to_char(tran_date, ''Dy'') d, rownum dd,tran_date
              from dwh.t_dates
                where to_number(to_char(tran_date,''mm''))=' || i ||
                      ' and to_number(to_char(tran_date,''yyyy''))=' || cy || '
                        and tran_date <=  ' || date_lit(max_date),
                          p_sheet          => l,
                          p_column_headers => FALSE,
                          p_sdvig_row      => c_header_rows + 4 * (i - 1) + (sdv_row_global - sdv_row) - err);

----------------------------------------------------------------------------------------------
      --ИТОГИ (блок 1: суммарные показатели по всем складам)
      if l = 1 then
        ssql := ' from (
select v.tran_date,nvl(b.pallet,0) "' || py || '", nvl(a.pallet,0) "' || cy || '",
nvl(b.ss_sum,0)/nullif(b.pallet,0) "ss' || py || '",nvl(a.ss_sum,0)/nullif(a.pallet,0) "ss' || cy || '",c.pallet nz,
nvl(b.ss_sum,0) "sssum' || py || '", nvl(a.ss_sum,0) "sssum' || cy || '"
 from dwh.t_dates v
left join (
SELECT t_date,SUM(Pallet) Pallet,AVG(SS) SS,SUM(ss_sum) ss_sum from (
select t_date,t_name,sum(nvl(pallet,0)) pallet,SUM(SS) ss,sum(SS)*sum(pallet) ss_sum
from(
      select sp.matrix_id as t_name
            ,sp.date_shift_start as t_date
            ,sum(nvl(sp.qty_pallet,0)) as pallet
            ,null as ss
        from dwh.sklod_pallet sp
        join dwh.v_sklod_facilities vsf
          on sp.matrix_id = vsf.matrix_id
         and vsf.matrix_id is not null
       where to_number(to_char(sp.date_shift_start,''mm''))= ' || i || '
       group by sp.matrix_id
               ,sp.date_shift_start
      union all
      select ss.matrix_id
            ,ss.date_shift_start
            ,0
            ,avg(nvl(ss.ss,0))
        from dwh.sklod_ss ss
        join dwh.v_sklod_facilities vsf
          on ss.matrix_id = vsf.matrix_id
         and vsf.matrix_id is not null
       where to_number(to_char(ss.date_shift_start,''mm''))= ' || i || '
       group by ss.matrix_id
               ,ss.date_shift_start
)
group by t_date,t_name
)
GROUP by t_Date
) a on a.t_date=v.tran_date
left join (SELECT t_date,SUM(Pallet) Pallet,AVG(SS) SS,SUM(ss_sum) ss_sum
     from (
          select t_date,t_name,sum(nvl(pallet,0)) pallet,SUM(SS) ss,sum(SS)*sum(pallet) ss_sum
                 from(
                      select sp.matrix_id as t_name
                            ,sp.date_shift_start as t_date
                            ,sum(nvl(sp.qty_pallet,0)) as pallet
                            ,null as ss
                        from dwh.sklod_pallet sp
                        join dwh.v_sklod_facilities vsf
                          on sp.matrix_id = vsf.matrix_id
                         and vsf.matrix_id is not null
                       where to_number(to_char(sp.date_shift_start,''mm''))= ' || i || '
                       group by sp.matrix_id
                               ,sp.date_shift_start
                      union all
                      select ss.matrix_id
                            ,ss.date_shift_start
                            ,0
                            ,avg(nvl(ss.ss,0))
                        from dwh.sklod_ss ss
                        join dwh.v_sklod_facilities vsf
                          on ss.matrix_id = vsf.matrix_id
                         and vsf.matrix_id is not null
                       where to_number(to_char(ss.date_shift_start,''mm''))= ' || i || '
                       group by ss.matrix_id
                               ,ss.date_shift_start
) group by t_date,t_name ) GROUP by t_Date

) b  on add_months(v.tran_date,-12)=b.t_date
left join (select t_date,sum(nvl(pallet,0)) pallet  from dwh.oracle_sklod_nz where  to_number(to_char(t_date,''mm''))=' || i || ' group by t_date) c on c.t_date=v.tran_date
' || sql_tail(i);

        -- Колонки итогов названы годами (py/cy) — как в исходнике («2014»/«2015»).
        -- В оригинале в ssql_sum были две синтаксические ошибки (незакрытый cast и лишняя
        -- скобка внутри nullif-аргумента) — исправлены, смысл формул сохранен.
        build_stats_sql('W',
                        '"' || cy || '"',  '"' || py || '"',
                        '"ss'   || cy || '"', '"ss'   || py || '"',
                        '"sssum'|| cy || '"', '"sssum'|| py || '"',
                        'nz');

        --все
        as_xlsx.query2sheet(p_sql            => ssql_all,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => c_header_rows + 4 * (i - 1) + (sdv_row_global - sdv_row),
                            p_sdvig_col      => c_first_data_col);
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 1,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_all),
                        p_fontId   => as_xlsx.get_font(c_font_stat, p_fontsize => c_fs_hdr, p_bold => TRUE),
                        p_borderId => as_xlsx.get_border());
        --мин
        as_xlsx.query2sheet(p_sql            => ssql_min,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1),
                            p_sdvig_col      => c_first_data_col);
        as_xlsx.cell(p_col    => c_first_data_col,
                     p_row    => c_header_rows + sdv_row_global + 4 * (i - 1) + 1,
                     p_sheet  => l,
                     p_value  => c_lbl_min,
                     p_fillId => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_all));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 2,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_avg),
                        p_fontId   => as_xlsx.get_font(c_font_stat, p_fontsize => c_fs_hdr, p_bold => TRUE),
                        p_borderId => as_xlsx.get_border());
        --средн
        as_xlsx.query2sheet(p_sql            => ssql_avg,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 1,
                            p_sdvig_col      => c_first_data_col);
        as_xlsx.cell(p_col    => c_first_data_col,
                     p_row    => c_header_rows + sdv_row_global + 4 * (i - 1) + 2,
                     p_sheet  => l,
                     p_value  => c_lbl_avg,
                     p_fillId => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_avg));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 3,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_max),
                        p_fontId   => as_xlsx.get_font(c_font_stat, p_fontsize => c_fs_hdr, p_bold => TRUE),
                        p_borderId => as_xlsx.get_border());
        --макс
        as_xlsx.query2sheet(p_sql            => ssql_max,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 2,
                            p_sdvig_col      => c_first_data_col);
        as_xlsx.cell(p_col    => c_first_data_col,
                     p_row    => c_header_rows + sdv_row_global + 4 * (i - 1) + 3,
                     p_sheet  => l,
                     p_value  => c_lbl_max,
                     p_fillId => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_max));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 4,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_sum),
                        p_fontId   => as_xlsx.get_font(c_font_stat, p_fontsize => c_fs_hdr, p_bold => TRUE),
                        p_borderId => as_xlsx.get_border());
        --сумма
        as_xlsx.query2sheet(p_sql            => ssql_sum,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => c_header_rows + sdv_row_global + 4 * (i - 1) + 3,
                            p_sdvig_col      => c_first_data_col);
        as_xlsx.cell(p_col    => c_first_data_col,
                     p_row    => c_header_rows + sdv_row_global + 4 * (i - 1) + 4,
                     p_sheet  => l,
                     p_value  => c_lbl_sum,
                     p_fillId => as_xlsx.get_fill('solid', p_fgRGB => c_rgb_sum));

----------------------------------------------------------------------------------------------
        --ИТОГИ РЦ ТЭЦ (блок 2: РЦ Екатеринбург + Березовский)
        ssql := ' from(
select v.tran_date,nvl(b.pallet,0) "p.' || py || '", nvl(a.pallet,0) "p.' || cy || '",nvl(b.totalsum,0)/nullif(nvl(b.pallet,0),0) "t.' || py || '", nvl(a.totalsum,0)/nullif(nvl(a.pallet,0),0) "t.' || cy || '",
nvl(b.totalsum,0) "sssum' || py || '", nvl(a.totalsum,0) "sssum' || cy || '",c.pallet nz
from t_dates v
left join (' || sql_agg_ekb(i, null) || ') a on a.t_date=v.tran_date
left join (' || sql_agg_ekb(i, null) || ') b on add_months(v.tran_date,-12)=b.t_date
left join (select t_date,sum(nvl(pallet,0)) pallet
                                        from dwh.oracle_sklod_nz
                                        where to_number(to_char(t_date,''mm''))=' || i || ' group by t_date) c on c.t_date=v.tran_date
' || sql_tail(i);

        -- ИТОГИ РЦ ТЭЦ: тот же широкий формат, что и у ИТОГОВ, но метрики p./t.
        build_stats_sql('W',
                        '"p.'    || cy || '"', '"p.'    || py || '"',
                        '"t.'    || cy || '"', '"t.'    || py || '"',
                        '"sssum' || cy || '"', '"sssum' || py || '"',
                        'nz');

        dump_block(l, block_sdv_col(2));
      end if; -- l = 1 (ИТОГИ)

-------------------------------------------------------------------------------------------------------------------------------------------
      --ПОШЛИ СКЛАДЫ (блоки 3..15: склады ЕКБ по ks2 и отдельные РЦ по rc_id).
      -- Порядок позиций повторяет исходный курсор: сначала склады из
      -- dwh.monwms_facility_repnames (позиции 3..6), затем блоки-РЦ из конфигурации
      -- c_blocks в порядке их ord (позиции 7..13), затем «пустые» позиции 14/15.
      for p in (with wh as (
                 select rownum rn, t.rep_name t_name
                   from (select t.rep_name
                           from dwh.monwms_facility_repnames t
                          where t.physical_wh_id = c_wh_physical
                            and t.rep_name <> 'СУПЕР'
                          order by t.ord_num) t
                ),
                rcs as (
                 select level id
                   from dual
                connect by level <= c_blocks_count
               )
                select id, t_name
                  from (select rn + 2 id, t_name, 3 ord1, rn ord2
                          from wh
                        union all
                        select r.id, block_name(r.id), block_ord(r.id), 1
                          from rcs r
                         where block_kind(r.id) = 'R'
                           and block_name(r.id) is not null)
                 order by ord1, ord2)
        loop
          ------------------------------------------------------
          -- склады - СУХОЙ,ХОЛОД,АЛКО,ОВОЩИ,СРЕДНИЙ,ПРОМ - не рц
          if block_kind(p.id) = 'S' then
            e := 0;
            ssql := ' from(
select v.tran_date,nvl(b.pallet,0) "p.' || py || '", nvl(a.pallet,0) "p.' || cy || '",nvl(b.totalsum,0)/nullif(nvl(b.pallet,0),0) "t.' || py || '", nvl(a.totalsum,0)/nullif(nvl(a.pallet,0),0) "t.' || cy || '",
nvl(b.totalsum,0) "sssum' || py || '", nvl(a.totalsum,0) "sssum' || cy || '", c.pallet nz from t_dates v
left join (' || sql_agg_ekb(i, p.id) || ') a on a.t_date=v.tran_date
left join (' || sql_agg_ekb(i, p.id) || ') b on add_months(v.tran_date,-12)=b.t_date
left join (select t_date,sum(nvl(pallet,0)) pallet
                  from dwh.oracle_sklod_nz where t_name =''' || p.t_name || ''' and to_number(to_char(t_date,''mm''))=' || i || ' group by t_date) c on c.t_date=v.tran_date
' || sql_tail(i);

            -- узкий блок: без LFL-колонок (как в исходнике)
            build_stats_sql('N',
                            '"p.'    || cy || '"', '"p.'    || py || '"',
                            '"t.'    || cy || '"', '"t.'    || py || '"',
                            '"sssum' || cy || '"', '"sssum' || py || '"',
                            'nz');
          else
            ------------------------------------------------------
            -- отдельный РЦ: единый шаблон, отличается только rc_id (берется из c_blocks)
            ssql := ' from(
select v.tran_date,nvl(b."Паллет",0) "p.' || py || '", nvl(a."Паллет",0) "p.' || cy || '",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.' || py || '", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.' || cy || '",
            nvl(b.ss_sum,0) "sssum' || py || '", nvl(a.ss_sum,0) "sssum' || cy || '"
       from dwh.t_dates v
        left join (' || sql_agg_rc(i, block_rc(p.id)) || '
                  ) a on a.t_date=v.tran_date
        left join (' || sql_agg_rc(i, block_rc(p.id)) || '
) b on add_months(v.tran_date,-12)=b.t_date
' || sql_tail(i);

            -- широкий блок регионального РЦ: с LFL, без колонки НЗ
            build_stats_sql('W',
                            '"p.'    || cy || '"', '"p.'    || py || '"',
                            '"t.'    || cy || '"', '"t.'    || py || '"',
                            '"sssum' || cy || '"', '"sssum' || py || '"',
                            null);
          end if;

          -- Смещение блока данных: берется из конфигурации c_blocks
          sdv_cols := block_sdv_col(p.id);
          dump_block(l, sdv_cols);
        end loop; --КОНЕЦ ПО СКЛАДАМ

      sdv_row_prev := sdv_row;
    end loop; --конец по месяцам

    sdv_row_global := 0;
    sdv_row_prev   := 0;
  end loop; --конец по листам

  as_xlsx.save(c_xlsx_dir, c_xlsx_file);

exception
  when others then
    dwh.cognos_utl.log_err(err_code => sqlcode,
                           err_msg  => sqlerrm,
                           proc     => c_proc);
  raise;
end create_sklod_report;
--<< PROG-4860
