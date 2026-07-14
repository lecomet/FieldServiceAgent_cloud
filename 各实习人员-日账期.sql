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
    e.emp_id,
    COALESCE(NULLIF(zw.staff_name, ''), NULLIF(e.ext9, '')) AS staff_name,
    COALESCE(NULLIF(zw.staff_num, ''), NULLIF(e.ext18, '')) AS staff_num,
    COALESCE(NULLIF(zw.attr_value, ''), NULLIF(e.ext21, '')) AS attr_value,
    wb.jobname AS job_name,
    e.zw_type,
    e.responsible_area,
    e.intern_start_time,
    CASE
      WHEN regexp_like(e.ext7, '^[0-9]+$') THEN c.area_name
      ELSE e.ext7
    END AS e_area_name,
    CASE
      WHEN regexp_like(e.ext8, '^[0-9]+$') THEN c.city_name
      ELSE e.ext8
    END AS e_city_name
  FROM dwd.dwd_d_mss_hr_emp_employee_ext_month e
  CROSS JOIN params pms
  LEFT JOIN dwd.dwd_d_mss_hr_wb_ps_msg_zw wb
    ON e.emp_id = wb.emp_id
    AND wb.month_id = pms.last_month_id
    AND wb.p_acct_day = pms.acct_day
  LEFT JOIN dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
    ON e.ext18 = zw.clerkcode
    AND zw.month_id = e."month"
    AND zw.p_acct_day = pms.acct_day
  LEFT JOIN dim.dim_city_no c
    ON e.ext7 = c.area_no
    AND e.ext8 = c.city_no
  WHERE e.p_acct_day = pms.acct_day
    AND e."month" = pms.month_id
    AND e.ext1 = '2'
    AND e.zw_type = '1'
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
    WHEN p.attr_value = '30' THEN '客调'
    WHEN p.attr_value = '40' THEN '代理商装维'
    WHEN p.attr_value = '50' THEN '装维经营承包'
    WHEN p.attr_value = '60' THEN '客户工程师'
    WHEN p.attr_value = '70' THEN '客户经理'
    WHEN p.attr_value = '72' THEN '智家工程师-家庭DICT'
    WHEN p.attr_value = '73' THEN '智家工程师-校园'
    WHEN p.attr_value = '80' THEN '其他人员'
    WHEN p.attr_value = '200' THEN '装维门店'
    WHEN p.attr_value IS NULL OR p.attr_value = '' THEN '实习人员'
    ELSE p.attr_value
  END AS "用工属性",
  CASE
    WHEN p.zw_type = '1' THEN '新装维'
    WHEN p.zw_type = '0' THEN '老装维'
    ELSE p.zw_type
  END AS "新老装维",
  p.responsible_area AS "维护区域",
  p.intern_start_time AS "实习开始时间",
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
