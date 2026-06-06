-- Nautilus Enterprise / B4 Engineering & Diving
-- PostgreSQL schema draft

CREATE TABLE companies (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    name TEXT NOT NULL,
    subscription_plan TEXT DEFAULT 'enterprise',
    financial_year_start_month INT DEFAULT 11,
    created_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE users (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    full_name TEXT NOT NULL,
    email TEXT NOT NULL,
    password_hash TEXT,
    role TEXT NOT NULL CHECK (role IN ('super_admin','company_admin','ops_manager','finance','field_user','client_viewer')),
    active BOOLEAN DEFAULT true,
    created_at TIMESTAMPTZ DEFAULT now(),
    UNIQUE(company_id, email)
);

CREATE TABLE divisions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    name TEXT NOT NULL,
    active BOOLEAN DEFAULT true,
    UNIQUE(company_id, name)
);

CREATE TABLE clients (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    name TEXT NOT NULL,
    contact_person TEXT,
    phone TEXT,
    email TEXT,
    address TEXT,
    vat_number TEXT,
    active BOOLEAN DEFAULT true,
    UNIQUE(company_id, name)
);

CREATE TABLE vessels (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    name TEXT NOT NULL,
    registration TEXT,
    client_id UUID REFERENCES clients(id),
    division_id UUID REFERENCES divisions(id),
    active BOOLEAN DEFAULT true,
    UNIQUE(company_id, name)
);

CREATE TABLE jobs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    job_number TEXT NOT NULL,
    job_date DATE NOT NULL,
    reporting_month TEXT,
    reporting_year INT,
    financial_year TEXT,
    client_id UUID REFERENCES clients(id),
    vessel_id UUID REFERENCES vessels(id),
    division_id UUID REFERENCES divisions(id),
    ops_manager_id UUID REFERENCES users(id),
    description TEXT,
    completion_date DATE,
    quote_number TEXT,
    po_number TEXT,
    report_number TEXT,
    invoice_number TEXT,
    start_time TIME,
    end_time TIME,
    total_hours NUMERIC(12,2) DEFAULT 0,
    revenue_incl_vat NUMERIC(14,2) DEFAULT 0,
    revenue_excl_vat NUMERIC(14,2) DEFAULT 0,
    labour_cost NUMERIC(14,2) DEFAULT 0,
    equipment_cost NUMERIC(14,2) DEFAULT 0,
    workshop_cost NUMERIC(14,2) DEFAULT 0,
    other_cost NUMERIC(14,2) DEFAULT 0,
    total_cost NUMERIC(14,2) GENERATED ALWAYS AS (labour_cost + equipment_cost + workshop_cost + other_cost) STORED,
    gross_profit NUMERIC(14,2) GENERATED ALWAYS AS (revenue_excl_vat - (labour_cost + equipment_cost + workshop_cost + other_cost)) STORED,
    status TEXT DEFAULT 'draft' CHECK (status IN ('draft','quoted','po_received','in_progress','completed','invoiced','paid','cancelled')),
    created_by UUID REFERENCES users(id),
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now(),
    UNIQUE(company_id, job_number)
);

CREATE TABLE employees (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    employee_number TEXT,
    full_name TEXT NOT NULL,
    position TEXT,
    hourly_rate NUMERIC(12,2) DEFAULT 0,
    monthly_salary NUMERIC(14,2) DEFAULT 0,
    active BOOLEAN DEFAULT true
);

CREATE TABLE equipment (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    name TEXT NOT NULL,
    serial_number TEXT,
    hourly_rate NUMERIC(12,2) DEFAULT 0,
    daily_rate NUMERIC(12,2) DEFAULT 0,
    active BOOLEAN DEFAULT true
);

CREATE TABLE job_labour (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    job_id UUID NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
    employee_id UUID REFERENCES employees(id),
    hours NUMERIC(12,2) NOT NULL DEFAULT 0,
    rate NUMERIC(12,2) NOT NULL DEFAULT 0,
    cost NUMERIC(14,2) GENERATED ALWAYS AS (hours * rate) STORED
);

CREATE TABLE job_equipment (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    job_id UUID NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
    equipment_id UUID REFERENCES equipment(id),
    hours_used NUMERIC(12,2) DEFAULT 0,
    days_used NUMERIC(12,2) DEFAULT 0,
    rate NUMERIC(12,2) DEFAULT 0,
    cost NUMERIC(14,2) DEFAULT 0
);

CREATE TABLE expenses (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    job_id UUID REFERENCES jobs(id),
    category TEXT NOT NULL,
    description TEXT,
    amount NUMERIC(14,2) NOT NULL DEFAULT 0,
    expense_date DATE NOT NULL,
    created_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE import_batches (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    uploaded_by UUID REFERENCES users(id),
    file_name TEXT NOT NULL,
    status TEXT DEFAULT 'uploaded' CHECK (status IN ('uploaded','validated','approved','committed','failed')),
    total_rows INT DEFAULT 0,
    valid_rows INT DEFAULT 0,
    error_rows INT DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT now()
);

CREATE TABLE import_staging_jobs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    batch_id UUID NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    row_number INT NOT NULL,
    raw_data JSONB NOT NULL,
    validation_status TEXT DEFAULT 'pending',
    validation_errors JSONB DEFAULT '[]'::jsonb,
    mapped_job JSONB
);

CREATE TABLE audit_log (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id UUID NOT NULL REFERENCES companies(id),
    user_id UUID REFERENCES users(id),
    entity_type TEXT NOT NULL,
    entity_id UUID,
    action TEXT NOT NULL,
    before_data JSONB,
    after_data JSONB,
    created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX idx_jobs_company_date ON jobs(company_id, job_date);
CREATE INDEX idx_jobs_company_status ON jobs(company_id, status);
CREATE INDEX idx_jobs_division ON jobs(division_id);
CREATE INDEX idx_jobs_ops_manager ON jobs(ops_manager_id);
