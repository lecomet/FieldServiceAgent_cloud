WITH params AS (
  SELECT
    ?{month_id} AS month_id,
    ?{acct_day} AS acct_day,
    date_format(date_parse(?{acct_day}, '%Y%m%d'), '%Y-%m-%d') AS biz_date,
    date_format(date_add('day', 1, date_parse(?{acct_day}, '%Y%m%d')), '%Y%m%d') AS point_p_acct_day,
    date_format(date_add('month', -3, date_parse(?{acct_day}, '%Y%m%d')), '%Y-%m-%d') AS intern_cutoff_date,
    date_format(date_add('month', -1, date_parse(?{month_id}, '%Y%m')), '%Y%m') AS last_month_id
),
city_dim AS (
  SELECT
    area_no,
    MAX(area_name) AS area_name,
    city_no,
    MAX(city_name) AS city_name
  FROM dim.dim_city_no
  GROUP BY area_no, city_no
),
zw_business_staff AS (
  SELECT DISTINCT
    zw.staff_id,
    zw.staff_num,
    zw.month_id,
    zw.attr_value,
    zw.area_no,
    zw.city_no
  FROM dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
  CROSS JOIN params pms
  LEFT JOIN dwd.dwd_d_mss_hr_emp_employee_ext_month e
    ON zw.emp_id = e.emp_id
    AND zw.month_id = e."month"
    AND e.p_acct_day = pms.acct_day
  WHERE zw.p_acct_day = pms.acct_day
    AND zw.is_off = '1'
    AND zw.month_id = pms.month_id
    AND e.zw_type IS NOT NULL
),
biz_county AS (
  SELECT
    s.area_no,
    COALESCE(MAX(NULLIF(s.area_name, '')), MAX(cd.area_name)) AS area_name,
    s.city_no,
    COALESCE(MAX(NULLIF(s.city_name, '')), MAX(cd.city_name)) AS city_name,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.yd ELSE 0 END) AS yd_zw_total,
    SUM(s.yd) AS yd_total,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.kd ELSE 0 END) AS zdkd_zw_total,
    SUM(s.kd) AS zdkd_total,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.qwzn_fttr ELSE 0 END) AS fttr_h_zw_total,
    SUM(s.qwzn_fttr) AS fttr_h_total,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.fttrb ELSE 0 END) AS fttr_b_zw_total,
    SUM(s.fttrb) AS fttr_b_total,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.rhzk151 ELSE 0 END) AS rhzk151_zw_total,
    SUM(s.rhzk151) AS rhzk151_total,
    SUM(CASE WHEN z.staff_id IS NOT NULL THEN s.tyzp ELSE 0 END) AS tyzp_zw_total,
    SUM(s.tyzp) AS tyzp_total
  FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge s
  CROSS JOIN params pms
  LEFT JOIN zw_business_staff z
    ON s.dvlp_staff_id = z.staff_id
    AND s.new_month = z.month_id
    AND z.attr_value IN ('10', '20', '50', '72', '73', '200')
    AND s.is_zw = 1
  LEFT JOIN city_dim cd
    ON s.area_no = cd.area_no
    AND s.city_no = cd.city_no
  WHERE s.p_acct_day = pms.acct_day
    AND s.new_month = pms.month_id
    AND s.new_date = pms.biz_date
  GROUP BY s.area_no, s.city_no
),
point_county AS (
  SELECT
    z.area_no,
    z.city_no,
    SUM(ROUND(COALESCE(TRY_CAST(pt.fz_point AS DOUBLE), 0), 0)) AS fz_point,
    SUM(ROUND(COALESCE(TRY_CAST(pt.yy_point AS DOUBLE), 0), 0)) AS yy_point
  FROM dwd.dwd_d_mss_hr_dwv_d_hrt_zw_point_zw_merge pt
  CROSS JOIN params pms
  LEFT JOIN zw_business_staff z
    ON pt.staff_num = z.staff_num
    AND z.attr_value IN ('10', '20', '50', '72', '73', '200')
  WHERE pt.acct_day = pms.acct_day
    AND pt.p_acct_day = pms.point_p_acct_day
    AND pt.attr_value IS NOT NULL
    AND z.staff_num IS NOT NULL
    AND (
      COALESCE(TRY_CAST(pt.fz_point AS DECIMAL(18, 2)), 0) != 0
      OR COALESCE(TRY_CAST(pt.yy_point AS DECIMAL(18, 2)), 0) != 0
    )
  GROUP BY z.area_no, z.city_no
),
people_county AS (
  SELECT
    area_no,
    city_no,
    SUM(total_people_all) AS total_people_all,
    SUM(total_people_eff) AS total_people_eff
  FROM (
    SELECT
      zw.area_no AS area_no,
      zw.city_no AS city_no,
      COUNT(DISTINCT zw.staff_num) AS total_people_all,
      COUNT(DISTINCT CASE WHEN zw.attr_value != '20' THEN zw.staff_num END) AS total_people_eff
    FROM dwd.dwd_d_mss_hr_emp_employee_ext_month emp
    CROSS JOIN params pms
    LEFT JOIN dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
      ON zw.emp_id = emp.emp_id
      AND emp."month" = zw.month_id
      AND zw.p_acct_day = pms.acct_day
    LEFT JOIN (
      SELECT DISTINCT wb0.code
      FROM dwd.dwd_d_mss_hr_wb_ps_msg_zw wb0
      CROSS JOIN params pms
      WHERE wb0.jobname IN (
        '智能云服务交付工程师',
        '智能云服务交付工程师#',
        '智慧家庭工程师',
        '云网标品交付维护能力建设',
        '云网标品交付维护组织'
      )
        AND wb0.month_id = pms.last_month_id
        AND wb0.p_acct_day = pms.acct_day
    ) wb
      ON zw.clerkcode = wb.code
    WHERE emp.p_acct_day = pms.acct_day
      AND emp.ext1 = '1'
      AND emp.is_confirm = 1
      AND zw.is_off = '1'
      AND zw.attr_value IN ('10', '20', '50', '72', '73')
      AND zw.month_id = pms.month_id
      AND wb.code IS NOT NULL
    GROUP BY zw.area_no, zw.city_no

    UNION ALL

    SELECT
      emp.ext7 AS area_no,
      emp.ext8 AS city_no,
      COUNT(DISTINCT emp.ext18) AS total_people_all,
      COUNT(DISTINCT emp.ext18) AS total_people_eff
    FROM dwd.dwd_d_mss_hr_emp_employee_ext_month emp
    CROSS JOIN params pms
    WHERE emp.p_acct_day = pms.acct_day
      AND emp.ext1 = '2'
      AND emp."month" = pms.month_id
      AND emp.intern_start_time <= pms.intern_cutoff_date
    GROUP BY emp.ext7, emp.ext8
  ) t
  GROUP BY area_no, city_no
)
SELECT
  b.area_name AS "地市",
  b.area_no AS "地市编码",
  b.city_name AS "区县",
  b.city_no AS "区县编码",
  b.yd_total AS "移动-全渠道(量)",
  b.yd_zw_total AS "移动-装维(量)",
  COALESCE(ROUND(CAST(b.yd_zw_total AS DOUBLE) / NULLIF(b.yd_total, 0), 3), 0) AS "移动-装维占比",
  b.zdkd_total AS "终端宽带-全渠道(量)",
  b.zdkd_zw_total AS "终端宽带-装维(量)",
  COALESCE(ROUND(CAST(b.zdkd_zw_total AS DOUBLE) / NULLIF(b.zdkd_total, 0), 3), 0) AS "终端宽带-装维占比",
  b.fttr_h_total AS "FTTR-全渠道(量)",
  b.fttr_h_zw_total AS "FTTR-装维(量)",
  COALESCE(ROUND(CAST(b.fttr_h_zw_total AS DOUBLE) / NULLIF(b.fttr_h_total, 0), 3), 0) AS "FTTR-装维占比",
  b.fttr_b_total AS "FTTR-B-全渠道(量)",
  b.fttr_b_zw_total AS "FTTR-B-装维(量)",
  COALESCE(ROUND(CAST(b.fttr_b_zw_total AS DOUBLE) / NULLIF(b.fttr_b_total, 0), 3), 0) AS "FTTR-B-装维占比",
  b.fttr_h_total + b.fttr_b_total AS "FTTR-H/B-全渠道(量)",
  b.fttr_h_zw_total + b.fttr_b_zw_total AS "FTTR-H/B-装维(量)",
  COALESCE(
    ROUND(
      CAST(b.fttr_h_zw_total + b.fttr_b_zw_total AS DOUBLE)
      / NULLIF(b.fttr_h_total + b.fttr_b_total, 0),
      3
    ),
    0
  ) AS "FTTR-H/B-装维占比",
  b.rhzk151_total AS "151-全渠道(量)",
  b.rhzk151_zw_total AS "151-装维(量)",
  COALESCE(ROUND(CAST(b.rhzk151_zw_total AS DOUBLE) / NULLIF(b.rhzk151_total, 0), 3), 0) AS "151-装维占比",
  b.tyzp_total AS "天翼智屏-全渠道(量)",
  b.tyzp_zw_total AS "天翼智屏-装维(量)",
  COALESCE(ROUND(CAST(b.tyzp_zw_total AS DOUBLE) / NULLIF(b.tyzp_total, 0), 3), 0) AS "天翼智屏-装维占比",
  COALESCE(p.fz_point, 0) AS "发展积分",
  COALESCE(p.yy_point, 0) AS "运营积分",
  COALESCE(p.fz_point, 0) + COALESCE(p.yy_point, 0) AS "价值积分",
  COALESCE(ps.total_people_all, 0) AS "总人数含班长",
  COALESCE(ps.total_people_eff, 0) AS "总人数不含班长",
  ROUND(
    CAST(COALESCE(p.fz_point, 0) + COALESCE(p.yy_point, 0) AS DOUBLE)
    / NULLIF(COALESCE(ps.total_people_eff, 0), 0),
    0
  ) AS "人均价值积分",
  ROUND(
    CAST(b.fttr_h_total + b.fttr_b_total AS DOUBLE)
    / NULLIF(COALESCE(ps.total_people_eff, 0), 0),
    2
  ) AS "人均FTTR-H/B"
FROM biz_county b
LEFT JOIN point_county p
  ON b.area_no = p.area_no
  AND b.city_no = p.city_no
LEFT JOIN people_county ps
  ON b.area_no = ps.area_no
  AND b.city_no = ps.city_no
ORDER BY b.area_no, b.city_no;
