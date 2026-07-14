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
person_base AS (
  SELECT DISTINCT
    e.emp_id,
    COALESCE(zw.staff_name, e.ext9) AS staff_name,
    COALESCE(zw.staff_num, e.ext18) AS staff_num,
    COALESCE(zw.attr_value, e.ext21) AS attr_value,
    wb.jobname AS job_name,
    e.zw_type,
    e.responsible_area,
    e.intern_start_time,
    CASE
      WHEN regexp_like(e.ext7, '^[0-9]+$') THEN c_code.area_name
      ELSE e.ext7
    END AS e_area_name,
    CASE
      WHEN regexp_like(e.ext8, '^[0-9]+$') THEN c_code.city_name
      ELSE e.ext8
    END AS e_city_name
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
  LEFT JOIN dim.dim_city_no c_code
    ON e.ext7 = c_code.area_no
    AND e.ext8 = c_code.city_no
  WHERE e.p_acct_day = pts.e_p_acct_day
    AND e.ext1 = '2'
    AND e."month" = pms.month_id
    AND wb.is_off = '1'
),
smr_agg AS (
  SELECT
    dvlp_staff_num,
    MAX(area_name) AS area_name,
    MAX(city_name) AS city_name,
    SUM(yd) AS yd,
    SUM(kd) AS kd,
    SUM(qwzn_fttr) AS qwzn_fttr,
    SUM(wifi_zd) AS wifi_zd,
    SUM(fttr_yf) AS fttr_yf,
    SUM(fttrb) AS fttrb,
    SUM(rhzk151) AS rhzk151,
    SUM(tyzp) AS tyzp,
    SUM(sqznzw) AS sqznzw
  FROM dwd.dwd_d_mss_hr_dwv_d_smr_user_list_result_merge
  CROSS JOIN params pms
  CROSS JOIN partitions pts
  WHERE p_acct_day = pms.acct_day
    AND new_month = pms.month_id
  GROUP BY dvlp_staff_num
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
  CROSS JOIN partitions pts
  WHERE pt.p_acct_day = pts.point_p_acct_day
    AND pt.acct_day = pms.acct_day
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
  p.intern_start_time AS "实习开始时间",
  COALESCE(s.yd, 0) AS "移动",
  COALESCE(s.kd, 0) AS "终端宽带",
  COALESCE(s.fttr_yf, 0) AS "FTTR-H",
  COALESCE(s.fttrb, 0) AS "FTTR-B",
  COALESCE(s.fttr_yf, 0) + COALESCE(s.fttrb, 0) AS "FTTR-H/B",
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
  ON p.staff_num = s.dvlp_staff_num
LEFT JOIN point_agg pt
  ON p.staff_num = pt.staff_num;
