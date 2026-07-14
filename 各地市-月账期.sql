WITH params AS (
  SELECT
    ?{month_id} AS month_id,
    ?{acct_day} AS acct_day,
    date_format(
      date_add('month', -1, date_parse(?{month_id}, '%Y%m')),
      '%Y%m'
    ) AS last_month_id
),
partitions AS (
  SELECT
    (
      SELECT MAX(p_acct_day)
      FROM dwd.dwd_d_mss_hr_dim_staff_zw10_zw
      WHERE month_id = (SELECT month_id FROM params)
    ) AS zw_p_acct_day,
    (
      SELECT MAX(p_acct_day)
      FROM dwd.dwd_d_mss_hr_emp_employee_ext_month
      WHERE "month" = (SELECT month_id FROM params)
    ) AS e_p_acct_day,
    (
      SELECT MAX(p_acct_day)
      FROM dwd.dwd_d_mss_hr_wb_ps_msg_zw
      WHERE month_id = (SELECT last_month_id FROM params)
    ) AS wb_p_acct_day,
    (
      SELECT MAX(p_acct_day)
      FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge
      WHERE new_month = (SELECT month_id FROM params)
    ) AS smr_p_acct_day,
    (
      SELECT MAX(p_acct_day)
      FROM dwd.dwd_d_mss_hr_dwv_d_hrt_zw_point_zw_merge
      WHERE acct_day = (SELECT acct_day FROM params)
    ) AS point_p_acct_day
),
zw_staff AS (
  SELECT DISTINCT
    zw.staff_num
  FROM dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
  CROSS JOIN params pms
  CROSS JOIN partitions pts
  LEFT JOIN dwd.dwd_d_mss_hr_wb_ps_msg_zw wb
    ON zw.clerkcode = wb.code
    AND wb.month_id = pms.last_month_id
    AND wb.p_acct_day = pts.wb_p_acct_day
  LEFT JOIN dwd.dwd_d_mss_hr_emp_employee_ext_month e
    ON zw.emp_id = e.emp_id
    AND zw.month_id = e."month"
    AND e.p_acct_day = pts.e_p_acct_day
  WHERE zw.p_acct_day = pts.zw_p_acct_day
    AND zw.is_off = '1'
    AND wb.is_off = '1'
    AND zw.month_id = pms.month_id
    AND e.ext1 = '1'
    AND e.is_confirm = 1

  UNION

  SELECT DISTINCT
    COALESCE(zw.staff_num, e.ext18) AS staff_num
  FROM dwd.dwd_d_mss_hr_emp_employee_ext_month e
  CROSS JOIN params pms
  CROSS JOIN partitions pts
  LEFT JOIN dwd.dwd_d_mss_hr_wb_ps_msg_zw wb
    ON e.emp_id = wb.emp_id
    AND wb.month_id = pms.last_month_id
    AND wb.p_acct_day = pts.wb_p_acct_day
  LEFT JOIN dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
    ON e.ext18 = zw.clerkcode
    AND zw.month_id = e."month"
    AND zw.p_acct_day = pts.zw_p_acct_day
  WHERE e.p_acct_day = pts.e_p_acct_day
    AND e.ext1 = '2'
    AND e."month" = pms.month_id
    AND wb.is_off = '1'
),
biz_area AS (
  SELECT
    s.area_no,
    MAX(s.area_name) AS area_name,

    SUM(CASE WHEN z.staff_num IS NOT NULL THEN s.kd ELSE 0 END) AS zdkd_zw_total,
    SUM(s.kd) AS zdkd_total,

    SUM(CASE WHEN z.staff_num IS NOT NULL THEN s.yd ELSE 0 END) AS yd_zw_total,
    SUM(s.yd) AS yd_total,

    SUM(CASE WHEN z.staff_num IS NOT NULL THEN s.fttr_yf ELSE 0 END) AS fttr_h_zw_total,
    SUM(s.fttr_yf) AS fttr_h_total,

    SUM(CASE WHEN z.staff_num IS NOT NULL THEN s.fttrb ELSE 0 END) AS fttr_b_zw_total,
    SUM(s.fttrb) AS fttr_b_total,

    SUM(CASE WHEN z.staff_num IS NOT NULL THEN s.rhzk151 ELSE 0 END) AS rhzk151_zw_total,
    SUM(s.rhzk151) AS rhzk151_total
  FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge s
  CROSS JOIN params pms
  CROSS JOIN partitions pts
  LEFT JOIN zw_staff z
    ON s.dvlp_staff_num = z.staff_num
  WHERE s.p_acct_day = pts.smr_p_acct_day
    AND s.new_month = pms.month_id
  GROUP BY s.area_no
),
point_area AS (
  SELECT
    CAST(pt.area_no AS VARCHAR) AS area_no,
    SUM(COALESCE(TRY_CAST(pt.fz_point AS DECIMAL(18, 2)), 0)) AS fz_point,
    SUM(COALESCE(TRY_CAST(pt.yy_point AS DECIMAL(18, 2)), 0)) AS yy_point
  FROM dwd.dwd_d_mss_hr_dwv_d_hrt_zw_point_zw_merge pt
  CROSS JOIN params pms
  CROSS JOIN partitions pts
  JOIN zw_staff z
    ON pt.staff_num = z.staff_num
  WHERE pt.p_acct_day = pts.point_p_acct_day
    AND pt.acct_day = pms.acct_day
  GROUP BY CAST(pt.area_no AS VARCHAR)
)
SELECT
  b.area_no AS "地市编码",
  b.area_name AS "地市",

  b.zdkd_zw_total AS "宽带装维发展量",
  b.zdkd_total AS "宽带全量发展量",
  COALESCE(ROUND(CAST(b.zdkd_zw_total AS DOUBLE) / NULLIF(b.zdkd_total, 0), 3), 0) AS "宽带装维占比",

  b.yd_zw_total AS "移动装维发展量",
  b.yd_total AS "移动全量发展量",
  COALESCE(ROUND(CAST(b.yd_zw_total AS DOUBLE) / NULLIF(b.yd_total, 0), 3), 0) AS "移动装维占比",

  b.fttr_h_zw_total AS "FTTR-H装维发展量",
  b.fttr_h_total AS "FTTR-H全量发展量",
  COALESCE(ROUND(CAST(b.fttr_h_zw_total AS DOUBLE) / NULLIF(b.fttr_h_total, 0), 3), 0) AS "FTTR-H装维占比",

  b.fttr_b_zw_total AS "FTTR-B装维发展量",
  b.fttr_b_total AS "FTTR-B全量发展量",
  COALESCE(ROUND(CAST(b.fttr_b_zw_total AS DOUBLE) / NULLIF(b.fttr_b_total, 0), 3), 0) AS "FTTR-B装维占比",

  b.fttr_h_zw_total + b.fttr_b_zw_total AS "FTTR-H/B装维发展量",
  b.fttr_h_total + b.fttr_b_total AS "FTTR-H/B全量发展量",
  COALESCE(
    ROUND(
      CAST(b.fttr_h_zw_total + b.fttr_b_zw_total AS DOUBLE)
      / NULLIF(b.fttr_h_total + b.fttr_b_total, 0),
      3
    ),
    0
  ) AS "FTTR-H/B装维占比",

  b.rhzk151_zw_total AS "151装维发展量",
  b.rhzk151_total AS "151全量发展量",
  COALESCE(ROUND(CAST(b.rhzk151_zw_total AS DOUBLE) / NULLIF(b.rhzk151_total, 0), 3), 0) AS "151装维占比",

  COALESCE(p.fz_point, 0) AS "发展积分",
  COALESCE(p.yy_point, 0) AS "运营积分",
  COALESCE(p.fz_point, 0) + COALESCE(p.yy_point, 0) AS "原始积分"
FROM biz_area b
LEFT JOIN point_area p
  ON b.area_no = p.area_no
ORDER BY b.area_no;
