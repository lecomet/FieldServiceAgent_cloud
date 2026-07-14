WITH params AS (
  SELECT
    ?{month_id} AS month_id,
    ?{acct_day} AS acct_day,
    ?{staff_name} AS staff_name_filter,
    date_format(date_parse(?{month_id}, '%Y%m'), '%Y-%m-01') AS month_start_date,
    date_format(date_parse(?{acct_day}, '%Y%m%d'), '%Y-%m-%d') AS end_date,
    date_format(date_parse(?{month_id}, '%Y%m'), '%Y%m01') AS acct_day_start,
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
    e.to_sys_date,
    c.area_name AS e_area_name,
    c.city_name AS e_city_name
  FROM dwd.dwd_d_mss_hr_dim_staff_zw10_zw zw
  CROSS JOIN params pms
  LEFT JOIN (
    SELECT DISTINCT
      wb0.code,
      wb0.jobname
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
    AND zw.attr_value IN ('10', '20', '50', '72', '73')
    AND wb.code IS NOT NULL
),
smr_agg AS (
  SELECT
    s.dvlp_staff_id,
    MAX(s.area_name) AS area_name,
    MAX(s.city_name) AS city_name,
    SUM(COALESCE(TRY_CAST(s.yd AS DECIMAL(18, 2)), 0)) AS yd,
    SUM(COALESCE(TRY_CAST(s.kd AS DECIMAL(18, 2)), 0)) AS kd,
    SUM(COALESCE(TRY_CAST(s.wifi_fw AS DECIMAL(18, 2)), 0)) AS wifi_fw,
    SUM(COALESCE(TRY_CAST(s.wifi_zd AS DECIMAL(18, 2)), 0)) AS wifi_zd,
    SUM(COALESCE(TRY_CAST(s.itv AS DECIMAL(18, 2)), 0)) AS itv,
    SUM(COALESCE(TRY_CAST(s.tysl AS DECIMAL(18, 2)), 0)) AS tysl,
    SUM(COALESCE(TRY_CAST(s.qwzn AS DECIMAL(18, 2)), 0)) AS qwzn,
    SUM(COALESCE(TRY_CAST(s.qwzn_fttr AS DECIMAL(18, 2)), 0)) AS qwzn_fttr,
    SUM(COALESCE(TRY_CAST(s.fttr_yf AS DECIMAL(18, 2)), 0)) AS fttr_yf,
    SUM(COALESCE(TRY_CAST(s.fttrb AS DECIMAL(18, 2)), 0)) AS fttrb,
    SUM(COALESCE(TRY_CAST(s.sqznzw AS DECIMAL(18, 2)), 0)) AS sqznzw,
    SUM(COALESCE(TRY_CAST(s.rhzk151 AS DECIMAL(18, 2)), 0)) AS rhzk151,
    SUM(COALESCE(TRY_CAST(s.tyzp AS DECIMAL(18, 2)), 0)) AS tyzp
  FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge s
  CROSS JOIN params pms
  WHERE s.p_acct_day = pms.acct_day
    AND s.new_month = pms.month_id
    AND s.new_date >= pms.month_start_date
    AND s.new_date <= pms.end_date
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
  WHERE pt.p_acct_day = pms.acct_day
    AND pt.acct_day >= pms.acct_day_start
    AND pt.acct_day <= pms.acct_day
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
    WHEN p.attr_value IS NULL OR p.attr_value = '' THEN '一线装维'
    ELSE p.attr_value
  END AS "用工属性",
  CASE
    WHEN p.zw_type = '1' THEN '新装维'
    WHEN p.zw_type = '0' THEN '老装维'
    ELSE p.zw_type
  END AS "新老装维",
  p.responsible_area AS "维护区域",
  p.to_sys_date AS "入职时间",
  COALESCE(s.yd, 0) AS "移动",
  COALESCE(s.kd, 0) AS "终端宽带",
  COALESCE(s.wifi_fw, 0) AS "全屋WiFi服务",
  COALESCE(s.wifi_zd, 0) AS "全屋WiFi终端",
  COALESCE(s.itv, 0) AS "天翼高清",
  COALESCE(s.tysl, 0) AS "天翼看家",
  COALESCE(s.qwzn, 0) AS "全屋智能",
  COALESCE(s.qwzn_fttr, 0) AS "FTTR-H",
  COALESCE(s.fttr_yf, 0) AS "FTTR月付",
  COALESCE(s.fttrb, 0) AS "FTTR-B",
  COALESCE(s.qwzn_fttr, 0) + COALESCE(s.fttrb, 0) AS "FTTR-H/B",
  COALESCE(s.sqznzw, 0) AS "社区智能组网",
  COALESCE(s.rhzk151, 0) AS "151",
  COALESCE(s.tyzp, 0) AS "天翼智屏",
  (
    COALESCE(s.yd, 0)
    + COALESCE(s.kd, 0)
    + COALESCE(s.wifi_fw, 0)
    + COALESCE(s.wifi_zd, 0)
    + COALESCE(s.itv, 0)
    + COALESCE(s.tysl, 0)
    + COALESCE(s.qwzn, 0)
    + COALESCE(s.qwzn_fttr, 0)
    + COALESCE(s.fttr_yf, 0)
    + COALESCE(s.sqznzw, 0)
    + COALESCE(s.fttrb, 0)
    + COALESCE(s.rhzk151, 0)
    + COALESCE(s.tyzp, 0)
  ) AS "全业务量",
  ROUND(
    COALESCE(s.kd, 0) * 0.5
    + COALESCE(s.wifi_zd, 0) * 0.25
    + COALESCE(s.qwzn_fttr, 0)
    + COALESCE(s.sqznzw, 0) * 0.25 * 1.5
    + COALESCE(s.fttrb, 0) * 1.5,
    1
  ) AS "折算量",
  (
    COALESCE(s.qwzn_fttr, 0) * 60
    + COALESCE(s.fttrb, 0) * 90
    + COALESCE(s.wifi_zd, 0) * 15
    + COALESCE(s.sqznzw, 0) * 22.5
  ) AS "加载积分",
  COALESCE(pt.yy_point, 0) AS "运营积分",
  COALESCE(pt.fz_point, 0) AS "发展积分",
  COALESCE(pt.fz_point, 0) + COALESCE(pt.yy_point, 0) AS "原始积分",
  COALESCE(pt.fz_point, 0)
  + (
    COALESCE(s.qwzn_fttr, 0) * 60
    + COALESCE(s.fttrb, 0) * 90
    + COALESCE(s.wifi_zd, 0) * 15
    + COALESCE(s.sqznzw, 0) * 22.5
  ) AS "评价积分"
FROM person_base p
CROSS JOIN params pms
LEFT JOIN smr_agg s
  ON p.staff_id = s.dvlp_staff_id
LEFT JOIN point_agg pt
  ON p.staff_num = pt.staff_num
WHERE pms.staff_name_filter IN ('全部', 'ALL', 'all')
  OR p.staff_name = pms.staff_name_filter
ORDER BY "地市", "区县", "工号"
