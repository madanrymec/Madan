-- Index-only tuning for 50 lakh+ row tables (no schema changes, no generated columns).
-- Run SHOW INDEX FROM <table>; first and skip any index that already exists.
-- After creating, run ANALYZE TABLE <table>; and check with EXPLAIN that the new index shows in `key`.

-- billing_data: filters on fin_year, Billing_Date, region/office
CREATE INDEX idx_bd_fy_date   ON billing_data (fin_year, Billing_Date);
CREATE INDEX idx_bd_date      ON billing_data (Billing_Date);
CREATE INDEX idx_bd_fy_unit   ON billing_data (fin_year, Unit);
CREATE INDEX idx_bd_region    ON billing_data (Sales_Region_Name);
CREATE INDEX idx_bd_office    ON billing_data (Sales_office);

-- billing_target / billing_target1
CREATE INDEX idx_bt_fy  ON billing_target  (fin_year);
CREATE INDEX idx_bt1_fy ON billing_target1 (fin_year);

-- order_data: filters on fin_year, Created_Date
CREATE INDEX idx_od_fy_date ON order_data (fin_year, Created_Date);
CREATE INDEX idx_od_date    ON order_data (Created_Date);

-- order_target / order_target1
CREATE INDEX idx_ot_fy  ON order_target  (fin_year);
CREATE INDEX idx_ot1_fy ON order_target1 (fin_year);

-- collections_data: Posting_Date range filters (now index-friendly after the DATE() rewrite)
CREATE INDEX idx_cd_posting ON collections_data (Posting_Date);

-- pending_order / stock tables: As-on-date lookups
CREATE INDEX idx_po_asondate ON pending_order (As_on_Date);
CREATE INDEX idx_sd_asondate ON stock_data (AS_ON_Date);
CREATE INDEX idx_sa_asondate ON stock_analysis (As_on_date);
CREATE INDEX idx_cs_asondate ON collections_set (AS_ON_Date);

-- purchase_registry
CREATE INDEX idx_pr_fy_post ON purchase_registry (fin_Year, Post_Date);
CREATE INDEX idx_pr_post    ON purchase_registry (Post_Date);

-- erm_data (MIS reports)
CREATE INDEX idx_erm_fy_acct_post ON erm_data (Fiscal_Year, Account_Number, Posting_Date);
