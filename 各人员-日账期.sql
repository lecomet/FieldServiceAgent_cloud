WITH params AS (
  SELECT
    ?{month_id} AS month_id,
    ?{acct_day} AS acct_day,
    date_format(date_parse(?{acct_day}, '%Y%m%d'), '%Y-%m-%d') AS biz_date,
    date_format(date_add('day', 1, date_parse(?{acct_day}, '%Y%m%d')), '%Y%m%d') AS point_p_acct_day,
    date_format(date_add('month', -1, date_parse(?{month_id}, '%Y%m')), '%Y%m') AS last_month_id
),
person_base AS (
  SELECT DISTINCT
    zw.staff_id,
    zw.emp_id,
    zw.staff_name,
    zw.staff_num,
    zw.attr_value,
    wb.jobname AS job_name,
    e.zw_type,
    e.responsible_area,
    c.area_name AS e_area_name,
    c.city_name AS e_city_name
  FROM dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
  CROSS JOIN params pms
  LEFT JOIN dwd.dwd_d_mss_hr_wb_ps_msg_zw wb
    ON zw.clerkcode = wb.code
    AND wb.month_id = pms.last_month_id
    AND wb.p_acct_day = pms.acct_day
  LEFT JOIN dwd.dwd_d_mss_hr_emp_employee_ext_month e
    ON zw.emp_id = e.emp_id
    AND zw.month_id = e."month"
    AND e.p_acct_day = pms.acct_day
  LEFT JOIN dim.dim_city_no c
    ON e.ext7 = c.area_no
    AND e.ext8 = c.city_no
  WHERE zw.p_acct_day = pms.acct_day
    AND zw.month_id = pms.month_id
    AND zw.is_off = '1'
    AND e.ext1 = '1'
    AND e.is_confirm = 1
    AND e.zw_type IN ('0', '1')
    AND zw.attr_value IN ('10', '20', '50', '72', '73', '200')
),
smr_agg AS (
  SELECT
    s.dvlp_staff_id,
    MAX(s.area_name) AS area_name,
    MAX(s.city_name) AS city_name,
    SUM(COALESCE(TRY_CAST(s.yd AS DECIMAL(18, 2)), 0)) AS yd,
    SUM(COALESCE(TRY_CAST(s.kd AS DECIMAL(18, 2)), 0)) AS kd,
    SUM(COALESCE(TRY_CAST(s.qwzn_fttr AS DECIMAL(18, 2)), 0)) AS qwzn_fttr,
    SUM(COALESCE(TRY_CAST(s.wifi_zd AS DECIMAL(18, 2)), 0)) AS wifi_zd,
    SUM(COALESCE(TRY_CAST(s.fttrb AS DECIMAL(18, 2)), 0)) AS fttrb,
    SUM(COALESCE(TRY_CAST(s.rhzk151 AS DECIMAL(18, 2)), 0)) AS rhzk151,
    SUM(COALESCE(TRY_CAST(s.tyzp AS DECIMAL(18, 2)), 0)) AS tyzp,
    SUM(COALESCE(TRY_CAST(s.sqznzw AS DECIMAL(18, 2)), 0)) AS sqznzw
  FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge s
  CROSS JOIN params pms
  WHERE s.p_acct_day = pms.acct_day
    AND s.new_month = pms.month_id
    AND s.new_date = pms.biz_date
    AND s.is_zw = 1
  GROUP BY s.dvlp_staff_id
),
point_agg AS (
  SELECT
    pt.staff_num,
    MAX(pt.area_name) AS area_name,
    MAX(pt.city_name) AS city_name,
    SUM(COALESCE(TRY_CAST(pt.fz_point AS DECIMAL(18, 2)), 0)) AS fz_point,
    SUM(COALESCE(TRY_CAST(pt.yy_point AS DECIMAL(18, 2)), 0)) AS yy_point
  FROM dwd.dwd_d_mss_hr_dwv_d_hrt_zw_point_zw_merge pt
  CROSS JOIN params pms
  WHERE pt.p_acct_day = pms.point_p_acct_day
    AND pt.acct_day = pms.acct_day
    AND pt.attr_value IS NOT NULL
  GROUP BY pt.staff_num
)
SELECT
  COALESCE(NULLIF(s.area_name, ''), NULLIF(pt.area_name, ''), NULLIF(p.e_area_name, '')) AS "地市",
  COALESCE(NULLIF(s.city_name, ''), NULLIF(pt.city_name, ''), NULLIF(p.e_city_name, '')) AS "区县",
  p.staff_name AS "装维姓名",
  p.staff_num AS "工号",
  p.job_name AS "岗位名称",
  CASE
    WHEN p.attr_value = '10' THEN '一线装维'
    WHEN p.attr_value = '20' THEN '装维班长'
    WHEN p.attr_value = '50' THEN '装维经营承包'
    WHEN p.attr_value = '72' THEN '智家工程师-家庭DICT'
    WHEN p.attr_value = '73' THEN '智家工程师-校园'
    WHEN p.attr_value = '200' THEN '装维门店'
    ELSE p.attr_value
  END AS "用工属性",
  CASE
    WHEN p.zw_type = '1' THEN '新装维'
    WHEN p.zw_type = '0' THEN '老装维'
    ELSE p.zw_type
  END AS "新老装维",
  p.responsible_area AS "维护区域",
  COALESCE(s.yd, 0) AS "移动",
  COALESCE(s.kd, 0) AS "终端宽带",
  COALESCE(s.qwzn_fttr, 0) AS "FTTR-H",
  COALESCE(s.fttrb, 0) AS "FTTR-B",
  COALESCE(s.qwzn_fttr, 0) + COALESCE(s.fttrb, 0) AS "FTTR-H/B",
  COALESCE(s.rhzk151, 0) AS "151",
  COALESCE(s.tyzp, 0) AS "天翼智屏",
  COALESCE(pt.fz_point, 0) AS "发展积分",
  COALESCE(pt.yy_point, 0) AS "运营积分",
  COALESCE(pt.fz_point, 0) + COALESCE(pt.yy_point, 0) AS "原始积分",
  (
    COALESCE(s.qwzn_fttr, 0) * 60
    + COALESCE(s.fttrb, 0) * 90
    + COALESCE(s.wifi_zd, 0) * 15
    + COALESCE(s.sqznzw, 0) * 22.5
  ) AS "加载积分",
  COALESCE(pt.fz_point, 0)
  + (
    COALESCE(s.qwzn_fttr, 0) * 60
    + COALESCE(s.fttrb, 0) * 90
    + COALESCE(s.wifi_zd, 0) * 15
    + COALESCE(s.sqznzw, 0) * 22.5
  ) AS "评价积分"
FROM person_base p
LEFT JOIN smr_agg s
  ON p.staff_id = s.dvlp_staff_id
LEFT JOIN point_agg pt
  ON p.staff_num = pt.staff_num
ORDER BY "地市", "区县", "工号"
