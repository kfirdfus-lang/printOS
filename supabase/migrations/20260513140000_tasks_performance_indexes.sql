CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status);
CREATE INDEX IF NOT EXISTS idx_tasks_created_at ON tasks(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_tasks_sales_agent_created ON tasks(sales_agent, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_tasks_active_recent ON tasks(created_at DESC)
WHERE status != 'ארכיון';

COMMENT ON INDEX idx_tasks_active_recent IS 'אינדקס חלקי - רק הזמנות פעילות, לדשבורד התפעולי';
