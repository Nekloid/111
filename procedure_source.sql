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
-- костыль на АБЛ
            as_xlsx.cell(p_col       => c + 6,
                         p_row       => 3,
                         p_sheet     => l,
                         p_value     => 'НЗ',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
        end if;
        if p.id in (1, 2, 9, 10, 11, 12, 13, 14 
                   ,15
                   ) /*Итоги + РЦ*/ then
          --костыль на РЦ
          if p.id in (1, 2) then
            as_xlsx.cell(p_col       => c + 9,p_row       => 3,p_sheet     => l,p_value     => 'НЗ',p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
          end if;
          as_xlsx.cell(p_col       => c + 5,p_row       => 3,p_sheet     => l,p_value     => 'LFL',
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE), p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col=> c + 3,p_row=> 3,p_sheet=> l,p_value=> py,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 4,p_row       => 3,p_sheet     => l,p_value     => cy,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
                       ----
          as_xlsx.cell(p_col       => c + 8,p_row       => 3,p_sheet     => l,p_value     => 'LFL',
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE), p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col=> c + 6,p_row=> 3,p_sheet=> l,p_value=> py,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 7,p_row       => 3,p_sheet     => l,p_value     => cy,
                       p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid', p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 11, p_bold     => TRUE),
                       p_borderId  => as_xlsx.get_border());
        end if;
----------------------------------------------------------------------------
        --ВТОРАЯ СТРОКА ШАПКИ
        if p.id in (1, 2, 9, 10, 11, 12, 13, 14  -- Итоги и РЦ
                   ,15
                   ) then 
          as_xlsx.mergecells(p_tl_col => c,p_br_col => c + 2,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
          as_xlsx.cell(p_col      => c + 1,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col      => c + 2,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c,p_row       => 2,p_sheet     => l,p_value     => 'Паллеты',p_alignment => as_xlsx.get_alignment(p_horizontal => 'center', p_vertical   => 'center'),
                       p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                       p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE), p_borderId  => as_xlsx.get_border());
          as_xlsx.mergecells(p_tl_col => c + 3,p_br_col => c + 5,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
          as_xlsx.cell(p_col      => c + 4,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col      => c + 5,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 3,p_row       => 2,p_sheet     => l,p_value     => 'СС паллета без НДС',p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                        p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                        p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11, p_bold     => TRUE),
                        p_borderId  => as_xlsx.get_border());
          /*Сумма в шапке*/
          as_xlsx.mergecells(p_tl_col => c + 6,p_br_col => c + 8,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
          as_xlsx.cell(p_col      => c + 7,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col      => c + 8,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
          as_xlsx.cell(p_col       => c + 6,p_row       => 2,p_sheet     => l,p_value     => 'Сумма отгрузки',p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                        p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                        p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11, p_bold     => TRUE),
                        p_borderId  => as_xlsx.get_border());

   if p.id in (1, 2) then -- Итоги и РЦ Екатеринбург (загрушка над НЗ)
            as_xlsx.cell(p_col=> c + 9,p_row=> 2,p_sheet=> l,p_value=> ' ',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
          end if;
        else

            as_xlsx.mergecells(p_tl_col => c,p_br_col => c + 1,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
            as_xlsx.cell(p_col      => c + 1,p_row      => 2,p_sheet    => l,p_value    => ' ', p_borderId => as_xlsx.get_border());
            as_xlsx.cell(p_col       => c,p_row       => 2,p_sheet     => l,p_value     => 'Паллеты',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
            as_xlsx.mergecells(p_tl_col => c + 2,p_br_col => c + 3,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
            as_xlsx.cell(p_col      => c + 3,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
            as_xlsx.cell(p_col       => c + 2,p_row       => 2,p_sheet     => l,p_value     => 'СС паллета без НДС',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
            as_xlsx.mergecells(p_tl_col => c + 4,p_br_col => c + 5,p_br_row => 2,p_tl_row => 2,p_sheet  => l);
            as_xlsx.cell(p_col      => c + 5,p_row      => 2,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
            as_xlsx.cell(p_col       => c + 4,p_row       => 2,p_sheet     => l,p_value     => 'Сумма отгрузки',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
            as_xlsx.cell(p_col       => c + 6,p_row       => 2,p_sheet     => l,p_value     => ' ',
                         p_alignment => as_xlsx.get_alignment(p_horizontal => 'center',p_vertical   => 'center'),
                         p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => rgb),
                         p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 11,p_bold     => TRUE),
                         p_borderId  => as_xlsx.get_border());
        end if;
      end loop;

   end if;

-------------------------------------------------------------------------------------------------------------------------------------
    --цикл по месяцам
    i_min := 1;
    for i in i_min .. i_max loop
      select count(*) into sdv_row from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
             and to_number(to_char(tran_date, 'yyyy')) = cy
             and tran_date <= max_date;
      sdv_row_global := sdv_row_global + sdv_row;
      --красим выходные
      for r in d + 1 + 4 * (i - 1) + sdv_row_global - sdv_row - err .. d + 1 + 4 *(i - 1) + sdv_row_global - 1 - err
        loop
        select to_number(to_char(tran_date, 'd')) into dw from t_dates
         where tran_date = trunc(tran_date, 'mm') + r - d - 1 - 4 * (i - 1) - (sdv_row_global - sdv_row) + err
           and to_number(to_char(tran_date, 'mm')) = i
           and to_number(to_char(tran_date, 'yyyy')) = cy;
        if (dw in (6, 7)) then
          as_xlsx.set_row(p_sheet     => l,
                          p_row       => r,
                          p_fillId    => as_xlsx.get_fill('solid',p_fgRGB => 'B3FFFF'),
                          p_fontId    => as_xlsx.get_font('Arial',p_fontsize => 10,p_bold     => TRUE,p_italic   => TRUE),
                          p_alignment => as_xlsx.get_alignment(p_horizontal => 'right', p_vertical   => 'center', p_wrapText   => FALSE),
                          p_borderId  => as_xlsx.get_border());
        else
          as_xlsx.set_row(p_sheet     => l,
                          p_row       => r,
                          p_fontId    => as_xlsx.get_font('Arial',
                                                          p_fontsize => 10),
                          p_alignment => as_xlsx.get_alignment(p_horizontal => 'right',p_vertical   => 'center',p_wrapText   => FALSE),
                          p_borderId  => as_xlsx.get_border());
        end if;
      end loop;
      --
      select tran_date
        into m_date
        from dwh.t_dates
       where to_number(to_char(tran_date, 'mm')) = i
         and to_number(to_char(tran_date, 'yyyy')) = cy
         and rownum = 1;
      as_xlsx.mergecells(p_tl_col => 1,p_br_col => 1,p_br_row => d + sdv_row_global + 4 * (i - 1) + 4 - err,p_tl_row => d + sdv_row_global + 4 * (i - 1) + 1 - err,p_sheet  => l);
      as_xlsx.mergecells(p_tl_col => 2,p_br_col => 2,p_br_row => d + sdv_row_global + 4 * (i - 1) + 4 - err,p_tl_row => d + sdv_row_global + 4 * (i - 1) + 1 - err,p_sheet  => l);
      --год
      as_xlsx.cell(p_col       => 1,
                   p_row       => d + sdv_row_global + 4 * (i - 1) + 1 - err,
                   p_sheet     => l,
                   p_value     => m_date,
                   p_fillId    => as_xlsx.get_fill('solid',
                                                   p_fgRGB => 'FFFFFF'),
                   p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 10),
                   p_numFmtId  => as_xlsx.get_numFmt(p_format => 'yyyy'),
                   p_borderId  => as_xlsx.get_border(),
                   p_alignment => as_xlsx.get_alignment(p_vertical   => 'center',
                                                        p_horizontal => 'center'));
      as_xlsx.cell(p_col      => 1,p_row      => d + sdv_row_global + 4 * (i - 1) + 1 + 1 - err,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
      as_xlsx.cell(p_col      => 1,p_row      => d + sdv_row_global + 4 * (i - 1) + 1 + 2 - err,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
      as_xlsx.cell(p_col      => 1,p_row      => d + sdv_row_global + 4 * (i - 1) + 1 + 3 - err,p_sheet    => l,p_value    => ' ',p_borderId => as_xlsx.get_border());
      --месяц
      as_xlsx.cell(p_col       => 2,
                   p_row       => d + sdv_row_global + 4 * (i - 1) + 1 - err,
                   p_sheet     => l,
                   p_value     => m_date,
                   p_fillId    => as_xlsx.get_fill('solid',
                                                   p_fgRGB => 'B7DEE8'),
                   p_fontId    => as_xlsx.get_font('Arial', p_fontsize => 12),
                   p_numFmtId  => as_xlsx.get_numFmt(p_format => 'mmmm'),
                   p_borderId  => as_xlsx.get_border(),
                   p_alignment => as_xlsx.get_alignment(p_rotate     => '90',
                                                        p_vertical   => 'center',
                                                        p_horizontal => 'center'));
      as_xlsx.cell(p_col => 2, p_row => d + sdv_row_global + 4 * (i - 1) + 1 + 1 - err, p_sheet    => l, p_value    => ' ', p_borderId => as_xlsx.get_border());
      as_xlsx.cell(p_col => 2, p_row => d + sdv_row_global + 4 * (i - 1) + 1 + 2 - err, p_sheet    => l, p_value    => ' ', p_borderId => as_xlsx.get_border());
      as_xlsx.cell(p_col => 2, p_row => d + sdv_row_global + 4 * (i - 1) + 1 + 3 - err, p_sheet    => l, p_value    => ' ', p_borderId => as_xlsx.get_border());
----------------------------------------------------------------------------------------------
 --ДАТЫ
as_xlsx.query2sheet(p_sql =>
      'select to_char(tran_date, ''Dy'') d, rownum dd,tran_date
              from dwh.t_dates
                where to_number(to_char(tran_date,''mm''))=' || i ||
                      ' and to_number(to_char(tran_date,''yyyy''))=' || cy || '
                        and tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'')',
                          p_sheet          => l,
                          p_column_headers => FALSE,
                          p_sdvig_row      => d + 4 * (i - 1) + (sdv_row_global - sdv_row) - err);
----------------------------------------------------------------------------------------------
 --ИТОГИ
if l = 1 then
ssql:= ' from (
select v.tran_date,nvl(b.pallet,0) "2014", nvl(a.pallet,0) "2015",
nvl(b.ss_sum,0)/nullif(b.pallet,0) "ss2014",nvl(a.ss_sum,0)/nullif(a.pallet,0) "ss2015",c.pallet nz,
nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
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
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))=' || cy || '
and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') )
order by tran_date';

        ssql_all := 'select "2014","2015",cast(("2015"/nullif("2014",0)-1) as binary_double),"ss2014","ss2015",cast(("ss2015"/nullif("ss2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double),nz' || ssql;
        ssql_min := 'select min("2014"),min("2015"),cast(min(("2015"/nullif("2014",0)-1)) as binary_double),min("ss2014"),min("ss2015"),cast(min("ss2015"/nullif("ss2014",0)-1) as binary_double),min("sssum2014"),min("sssum2015"),cast(min("sssum2015"/nullif("sssum2014",0)-1) as binary_double),min(nz)' || ssql;
        ssql_avg := 'select avg("2014"),avg("2015"),cast(avg("2015")/nullif(avg("2014"),0)-1 as binary_double),avg("ss2014"),avg("ss2015"),cast(avg("ss2015"/nullif("ss2014",0)-1) as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015"/nullif("sssum2014",0)-1) as binary_double),avg(nz)' || ssql;
        ssql_max := 'select max("2014"),max("2015"),cast(max(("2015"/nullif("2014",0)-1)) as binary_double),max("ss2014"),max("ss2015"),cast(max("ss2015"/nullif("ss2014",0)-1) as binary_double),max("sssum2014"),max("sssum2015"),cast(max("sssum2015"/nullif("sssum2014",0)-1) as binary_double),max(nz)' || ssql;
        ssql_sum := 'select sum("2014"),sum("2015"),cast(sum("2015")/nullif(sum("2014"),0)-1) as binary_double),avg("ss2014"),avg("ss2015"),cast(avg("ss2015"/nullif("ss2014",0)-1) as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014",0)-1) as binary_double),sum(nz)' || ssql;
        --все
        as_xlsx.query2sheet(p_sql=> ssql_all, p_sheet => l, p_column_headers => FALSE,
                            p_sdvig_row      => d + 4 * (i - 1) + (sdv_row_global - sdv_row),
                            p_sdvig_col      => 3);
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => d + sdv_row_global + 4 * (i - 1) + 1,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => 'F1FB93'),
                        p_fontId   => as_xlsx.get_font('Arial Cyr',  p_fontsize => 11, p_bold     => TRUE),
                        p_borderId => as_xlsx.get_border());
        --мин
        as_xlsx.query2sheet(p_sql            => ssql_min,  p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => d + sdv_row_global + 4 * (i - 1),
                            p_sdvig_col      => 3);
        as_xlsx.cell(p_col => 3,p_row => d + sdv_row_global + 4 * (i - 1) + 1, p_sheet  => l, p_value  => 'Минимум', p_fillId => as_xlsx.get_fill('solid', p_fgRGB => 'F1FB93'));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => d + sdv_row_global + 4 * (i - 1) + 2,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => '95FD9A'),
                        p_fontId   => as_xlsx.get_font('Arial Cyr', p_fontsize => 11, p_bold     => TRUE),
                        p_borderId => as_xlsx.get_border());
        --средн
        as_xlsx.query2sheet(p_sql            => ssql_avg,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => d + sdv_row_global + 4 * (i - 1) + 1,
                            p_sdvig_col      => 3);
        as_xlsx.cell(p_col    => 3,p_row    => d + sdv_row_global + 4 * (i - 1) + 2,p_sheet  => l,p_value  => 'Среднее',p_fillId => as_xlsx.get_fill('solid', p_fgRGB => '95FD9A'));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => d + sdv_row_global + 4 * (i - 1) + 3,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => 'FFCCCC'),
                        p_fontId   => as_xlsx.get_font('Arial Cyr', p_fontsize => 11, p_bold     => TRUE),
                        p_borderId => as_xlsx.get_border());
        --макс
        as_xlsx.query2sheet(p_sql            => ssql_max,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => d + sdv_row_global + 4 * (i - 1) + 2,
                            p_sdvig_col      => 3);
        as_xlsx.cell(p_col    => 3,p_row    => d + sdv_row_global + 4 * (i - 1) + 3,p_sheet  => l,p_value  => 'Максимум',p_fillId => as_xlsx.get_fill('solid', p_fgRGB => 'FFCCCC'));
        as_xlsx.set_row(p_sheet    => l,
                        p_row      => d + sdv_row_global + 4 * (i - 1) + 4,
                        p_fillId   => as_xlsx.get_fill('solid', p_fgRGB => 'FFB953'),
                        p_fontId   => as_xlsx.get_font('Arial Cyr', p_fontsize => 11, p_bold     => TRUE),
                        p_borderId => as_xlsx.get_border());
        --сумма
        as_xlsx.query2sheet(p_sql            => ssql_sum,
                            p_sheet          => l,
                            p_column_headers => FALSE,
                            p_sdvig_row      => d + sdv_row_global + 4 * (i - 1) + 3,
                            p_sdvig_col      => 3);
        as_xlsx.cell(p_col    => 3, p_row    => d + sdv_row_global + 4 * (i - 1) + 4,p_sheet  => l, p_value  => 'Сумма', p_fillId => as_xlsx.get_fill('solid', p_fgRGB => 'FFB953'));

----------------------------------------------------------------------------------------------
  --ИТОГИ РЦ ТЭЦ
ssql := ' from(
select v.tran_date,nvl(b.pallet,0) "p.2014", nvl(a.pallet,0) "p.2015",nvl(b.totalsum,0)/nullif(nvl(b.pallet,0),0) "t.2014", nvl(a.totalsum,0)/nullif(nvl(a.pallet,0),0) "t.2015",
nvl(b.totalsum,0) "sssum2014", nvl(a.totalsum,0) "sssum2015",c.pallet nz
from t_dates v
left join (select t_date, sum(pallet) pallet,sum(totalsum) totalsum from (      
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
       where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
         and vsf.rc_id = 2  -- екб+березовск
       group by sp.matrix_id
               ,sp.date_shift_start
                                                                                    ) group by t_date) a on a.t_date=v.tran_date
left join (select t_date, sum(pallet) pallet,sum(totalsum) totalsum
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
                                 where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                                   and vsf.rc_id = 2  -- екб+березовск
                                 group by sp.matrix_id
                                         ,sp.date_shift_start                      
                                                 ) group by t_date) b                        on add_months(v.tran_date,-12)=b.t_date

        left join (select t_date,sum(nvl(pallet,0)) pallet
                                        from dwh.oracle_sklod_nz
                                        where to_number(to_char(t_date,''mm''))=' || i || ' group by t_date) c                          on c.t_date=v.tran_date
 where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))=' || cy || ' and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';

        ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double),nz' || ssql;
        ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double),min(nz)' || ssql;
        ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double),avg(nz)' || ssql;
        ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double),max(nz)' || ssql;
        ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double),sum(nz)' || ssql;
        --все
        as_xlsx.query2sheet(p_sql => ssql_all, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + 4 * (i - 1) + (sdv_row_global - sdv_row), p_sdvig_col      => 13);
        --мин, средн, макс, сумма:
        as_xlsx.query2sheet(p_sql => ssql_min, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1), p_sdvig_col => 13);
        as_xlsx.query2sheet(p_sql => ssql_avg, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 1, p_sdvig_col => 13);
        as_xlsx.query2sheet(p_sql => ssql_max, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 2, p_sdvig_col => 13);
        as_xlsx.query2sheet(p_sql => ssql_sum, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 3, p_sdvig_col => 13);
        --mix
-------------------------------------------------------------------------------------------------------------------------------------------
 --ПОШЛИ СКЛАДЫ
 if 1=1 then

        FOR P IN (
                 select rownum id,t_name from
                          (select t_name from (
                        select 3 ord1, rownum ord2, t_name, matrix_id from (select t.rep_name as t_name, t.matrix_id from dwh.monwms_facility_repnames t
                               where t.physical_wh_id = 2  and t.rep_name <> 'СУПЕР' order by t.ord_num)
                        --union all
                        --select 4,1,'РЦ БЕРЕЗОВСКИЙ',null from dual     -- исключили отдельное отображение березовского.
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
    LOOP
------------------------------------------------------
--склады - СУХОЙ,ХОЛОД,АЛКО,ОВОЩИ,СРЕДНИЙ,ПРОМ - не рц
      CASE
        --
        WHEN (P.ID < 7) THEN
              e:= 0;
              ssql:= ' from(
select v.tran_date,nvl(b.pallet,0) "p.2014", nvl(a.pallet,0) "p.2015",nvl(b.totalsum,0)/nullif(nvl(b.pallet,0),0) "t.2014", nvl(a.totalsum,0)/nullif(nvl(a.pallet,0),0) "t.2015",
nvl(b.totalsum,0) "sssum2014", nvl(a.totalsum,0) "sssum2015", c.pallet nz from t_dates v
left join (select t_date, sum(pallet) pallet,sum(totalsum) totalsum
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
                         where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                           and vsf.rc_id = 2 -- екб+березовск
                           and vsf.ks2 = ' || p.id || ' -- это потому что по складам
                         group by sp.matrix_id
                                 ,sp.date_shift_start
                              )
                              group by t_date) a on a.t_date=v.tran_date
left join (select t_date, sum(pallet) pallet,sum(totalsum) totalsum
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
                         where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                           and vsf.rc_id = 2 -- екб+березовск
                           and vsf.ks2 = ' || p.id || ' -- это потому что по складам
                         group by sp.matrix_id
                                 ,sp.date_shift_start
                              )
                              group by t_date) b on add_months(v.tran_date,-12)=b.t_date
left join (select t_date,sum(nvl(pallet,0)) pallet
                  from dwh.oracle_sklod_nz where t_name =''' || p.t_name || ''' and to_number(to_char(t_date,''mm''))=' || i || ' group by t_date) c on c.t_date=v.tran_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))=' || cy || ' and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
---
              ssql_all := 'select "p.2014","p.2015","t.2014","t.2015","sssum2014","sssum2015",nz' || ssql;
              ssql_min := 'select min("p.2014"),min("p.2015"),min("t.2014"),min("t.2015"),min("sssum2014"),min("sssum2015"),min(nz)' || ssql;
              ssql_avg := 'select avg("p.2014"),avg("p.2015"),avg("t.2014"),avg("t.2015"),avg("sssum2014"),avg("sssum2015"),avg(nz)' || ssql;
              ssql_max := 'select max("p.2014"),max("p.2015"),max("t.2014"),max("t.2015"),max("sssum2014"),max("sssum2015"),max(nz)' || ssql;
              ssql_sum := 'select sum("p.2014"),sum("p.2015"),avg("t.2014"),avg("t.2015"),sum("sssum2014"),sum("sssum2015"),sum(nz)' || ssql;
------------------------------------------------------
       WHEN P.ID = 7 THEN /*РЦ УФА2 + РЦ Уфа*/
              ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2152
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2152
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
              ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
              ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
              ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
------------------------------------------------------
       WHEN P.ID = 8 THEN /*РЦ НЕФТЕЮГАНСК + РЦ СУРГУТ*/
              e:= 1;
              ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2014
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2014
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
              ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
              ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
              ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
------------------------------------------------------
       WHEN P.ID = 9 THEN /* РЦ НОВОСИБИРСК + РЦ НСК*/
              e:= 2;
              ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2037
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 2037
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
              ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
              ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
              ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
              ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
------------------------------------------------------
       WHEN P.ID = 10 THEN /* РЦ ПЕРМЬ*/
                      ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3002
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3002
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
                      ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
                      ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
                      ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
       WHEN P.ID = 11 THEN /* РЦ ЮЖНОУРАЛЬСК*/
                      ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3013
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3013
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
                      ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
                      ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
                      ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
       WHEN P.ID = 12 THEN /* РЦ ПЕРМЬ2*/
                      ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3025
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3025
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
                      ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
                      ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
                      ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
       WHEN P.ID = 13 THEN /* РЦ ТЮМЕНЬ*/
                      ssql:= ' from(
select v.tran_date,nvl(b."Паллет",0) "p.2014", nvl(a."Паллет",0) "p.2015",nvl(b.ss_sum/nullif(b."Паллет",0),0) "t.2014", nvl(a.ss_sum/nullif(a."Паллет",0),0) "t.2015",
            nvl(b.ss_sum,0) "sssum2014", nvl(a.ss_sum,0) "sssum2015"
       from dwh.t_dates v
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3028
                   group by sp.date_shift_start
                  ) a on a.t_date=v.tran_date
        left join (
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
                   where to_number(to_char(sp.date_shift_start,''mm''))=' || i || '
                     and vsf.rc_id = 3028
                   group by sp.date_shift_start
) b on add_months(v.tran_date,-12)=b.t_date
where to_number(to_char(v.tran_date,''mm''))=' || i || ' and to_number(to_char(v.tran_date,''yyyy''))= ' || cy || '  and v.tran_date <=  to_date('''|| to_char(max_date,'dd.mm.yyyy') || ''',''dd.mm.yyyy'') ) order by tran_date';
                      ssql_all := 'select "p.2014","p.2015",cast(("p.2015"/nullif("p.2014",0)-1) as binary_double),"t.2014","t.2015",cast(("t.2015"/nullif("t.2014",0)-1) as binary_double),"sssum2014","sssum2015",cast(("sssum2015"/nullif("sssum2014",0)-1) as binary_double)' || ssql;
                      ssql_min := 'select min("p.2014"),min("p.2015"),cast(min(("p.2015"/nullif("p.2014",0)-1)) as binary_double),min("t.2014"),min("t.2015"),cast(min(("t.2015"/nullif("t.2014",0)-1)) as binary_double),min("sssum2014"),min("sssum2015"),cast(min(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_avg := 'select avg("p.2014"),avg("p.2015"),cast(avg("p.2015")/nullif(avg("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),avg("sssum2014"),avg("sssum2015"),cast(avg("sssum2015")/nullif(avg("sssum2014"),0)-1 as binary_double)' || ssql;
                      ssql_max := 'select max("p.2014"),max("p.2015"),cast(max(("p.2015"/nullif("p.2014",0)-1)) as binary_double),max("t.2014"),max("t.2015"),cast(max(("t.2015"/nullif("t.2014",0)-1)) as binary_double),max("sssum2014"),max("sssum2015"),cast(max(("sssum2015"/nullif("sssum2014",0)-1)) as binary_double)' || ssql;
                      ssql_sum := 'select sum("p.2014"),sum("p.2015"),cast(sum("p.2015")/nullif(sum("p.2014"),0)-1 as binary_double),avg("t.2014"),avg("t.2015"),cast(avg("t.2015")/nullif(avg("t.2014"),0)-1 as binary_double),sum("sssum2014"),sum("sssum2015"),cast(sum("sssum2015")/nullif(sum("sssum2014"),0)-1 as binary_double)' || ssql;
   end case;
          --Смещение блока данных:
          sdv_cols:= case  when p.id = 1 then 23
                when p.id = 2 then 30
                when p.id = 3 then 37
                when p.id = 4 then 44
                when p.id = 5 then 51
                when p.id = 6 then 58
                when p.id = 7 then 65
                when p.id = 8 then 74
                when p.id = 9 then 83
                when p.id = 10 then 92
                when p.id = 11 then 101
                when p.id = 12 then 110
                when p.id = 13 then 119
                  else -1 end;
          as_xlsx.query2sheet(p_sql => ssql_all, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + 4 * (i - 1) + (sdv_row_global - sdv_row), p_sdvig_col => sdv_cols);
          as_xlsx.query2sheet(p_sql => ssql_min, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1), p_sdvig_col => sdv_cols);
          as_xlsx.query2sheet(p_sql => ssql_avg, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 1, p_sdvig_col => sdv_cols);
          as_xlsx.query2sheet(p_sql => ssql_max, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 2, p_sdvig_col => sdv_cols);
          as_xlsx.query2sheet(p_sql => ssql_sum, p_sheet => l, p_column_headers => FALSE, p_sdvig_row => d + sdv_row_global + 4 * (i - 1) + 3, p_sdvig_col => sdv_cols);
        END LOOP; --КОНЕЦ ПО СКЛАДАМ
 end if;
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        sdv_row_prev := sdv_row;
    end if;
    end loop; --конец по месяцам
    sdv_row_global := 0;
    sdv_row_prev   := 0;
  end loop; --конец по листам
   as_xlsx.save('SVOD_DIR', 'sklod_report.xlsx');

exception
  when others then
    dwh.cognos_utl.log_err(err_code  => sqlcode
                          ,err_msg   => sqlerrm
                          ,proc      => c_proc
                          );
  raise;
end create_sklod_report;
--<< PROG-4860
