-->> PROG-4860
procedure create_sklod_report(pmaxdate in date default trunc(sysdate)-1) is
  --Отчет показатели логистики
  c_proc         constant varchar2(200) := 'LOGISTICS.CREATE_SKLOD_REPORT';
  max_date       date:= trunc(pmaxdate);--trunc(sysdate)-1; -- максимальная дата отчета.
  sdv_row_global pls_integer:=0;
  sdv_row_prev   pls_integer:=0;
  sdv_row        pls_integer:=0;
  sdv_cols       pls_integer:=0; -- сдвиг блока данных по колонкам
  i_max          pls_integer:=to_number(to_char(max_date, 'mm')); -- количество месяцев
  i_min          pls_integer:=1;   -- начальный месяц для цикла по месяцам
  r_max_lm       pls_integer:=trunc(max_date, 'mm') - trunc(max_date, 'yyyy');   --количество строк текущего месяца
  d              pls_integer :=3;   --отступ шапка в строках
  l              integer;
  w              integer;
  c              pls_integer:=0;
  e              integer;
  r              integer;
  cy             pls_integer:= to_number(to_char(max_date, 'yyyy')); -- текущий год
  py             pls_integer:= to_number(to_char(trunc(max_date, 'yyyy') - 1, 'yyyy')); -- предыдущий год
  dw             pls_integer:=1;   -- порядковый номер дня недели
  err            pls_integer:=0;   --12; --костыль, чтобы третий лист с апреля начинался
  m_date         date;
  ssql_min       varchar2(32000);
  ssql_avg       varchar2(32000);          ssql_max       varchar2(32000);
  ssql_sum       varchar2(32000);          ssql_all       varchar2(32000);
  ssql           varchar2(32000);          rgb            varchar2(50);
begin
  execute immediate 'alter session set nls_language=RUSSIAN';
  execute immediate 'alter session set nls_territory=CIS';
  execute immediate 'alter session set NLS_NUMERIC_CHARACTERS=".,"';
  as_xlsx.clear_workbook;
  --добавляем листы
  as_xlsx.new_sheet('Показатели логистики', (r_max_lm + d + (i_max - 2) * 4));
-------------------------------------------------------------------------------------------------------------------------
  --Цикл по листам
  FOR L IN 1 .. 1 LOOP
    --ширина столбцов
    for c in 1 .. 150 loop
      case
        when (c = 1) then w := 5.5;
        when (c = 2) then w := 4;
        when (c = 3) then w := 12.5;
        else
          case
            when l = 1 then w := 16;
            else            w := 20;
          end case;
      end case;
      as_xlsx.set_column_width(p_col => c, p_sheet => l, p_width => w);
    end loop;
  --ШАПКА
    if l = 1 then
      --1-ый лист
      for p in (
        select rownum id,t_name from
              (select t_name from (
            select 3 ord1, rownum ord2, t_name, matrix_id from (select t.rep_name as t_name, t.matrix_id from dwh.monwms_facility_repnames t
                                                                       where t.physical_wh_id = 2  and t.rep_name <> 'СУПЕР' order by t.ord_num)
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
            ) order by ord1,ord2)
        )
      loop
       c:= case when p.id = 1 then 4
                when p.id = 2 then 14
                when p.id = 3 then 24
                when p.id = 4 then 31
                when p.id = 5 then 38
                when p.id = 6 then 45
                when p.id = 7 then 52
                when p.id = 8 then 59

                when p.id = 9 then 66
                when p.id = 10 then 75
                when p.id = 11 then 84
                when p.id = 12 then 93
                when p.id = 13 then 102
                when p.id = 14 then 111
                when p.id = 15 then 120
                  else -1 end;

        rgb := 'B7DEE8';
        case
          when p.id = 1 then   rgb := 'D9D9D9';
          when p.id = 2 then   rgb := 'D9D9D9';
              else                 c := c + 0/*4*/;
        end case;
        --объединение ячеек в первой строке шапки
        as_xlsx.mergecells(p_tl_col => c,p_br_col => (c + 4 + case when p.id in (1, 2) then 5
                                                                   when p.id in (9, 10, 11, 12,13, 14
                                                                                ,15
                                                                                ) then  4 else 2 end),
                           p_br_row => 1, p_tl_row => 1, p_sheet  => l);
        as_xlsx.cell(p_col=> c + 1,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col=> c + 2,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());
        as_xlsx.cell(p_col=> c + 3,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());

        if p.id in (9, 10, 11, 12, 13, 14 
                   ,15 
                   ) then
          as_xlsx.cell(p_col=> c + 8,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());
        end if;
        --костыль для РЦ
        if p.id in (1, 2) then
          as_xlsx.cell(p_col=> c + 8,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col=> c + 9,p_row=> 1,p_sheet=> l,p_value=> ' ',p_borderId => as_xlsx.get_border());
        end if;
        --первая строчка шапки
        as_xlsx.cell(p_col       => c,p_row       => 1,p_sheet     => l,p_value     => p.t_name,
                     p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                     p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                     p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                     p_borderId  => as_xlsx.get_border());
        --третья строчка шапки (куча костылей,быдло-коддинг)
        as_xlsx.cell(p_col       => c,p_row       => 3,p_sheet     => l,p_value     => py,
                     p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                     p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                     p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold => TRUE),
                     p_borderId  => as_xlsx.get_border());
        as_xlsx.cell(p_col       => c + 1,p_row       => 3,p_sheet     => l,p_value     => cy,
                     p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                     p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                     p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                     p_borderId  => as_xlsx.get_border());
  if p.id not in  (1, 2, 9, 10, 11, 12, 13, 14
                  ,15) then
          --костыль на АБЛ и РЦ и регионы
          as_xlsx.cell(p_col => c + 2,p_row => 3,p_sheet => l,p_value => py,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 3,p_row       => 3,p_sheet     => l,p_value     => cy,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col => c + 4,p_row => 3,p_sheet => l,p_value => py,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 5,p_row       => 3,p_sheet     => l,p_value     => cy,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());

        else
          as_xlsx.cell(p_col       => c + 2,p_row       => 3,p_sheet     => l,p_value     => 'LFL',
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
        end if;
        --
        if p.id > 2 and p.id not in (9, 10, 11, 12, 13, 14
                                    ,15
                                    ) then
