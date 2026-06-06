#!/usr/bin/env python3
"""
Load all sheets from the Excel file into PostgreSQL.
Set DATABASE_URL env var or use Railway's auto-injected env vars (PGHOST, PGPORT, etc).

Usage:
    python3 load_data.py
"""
import os
import sys
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text, inspect


def get_db_url():
    """Build PostgreSQL connection URL from env vars."""
    # Railway provides: PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE
    # Or use DATABASE_URL if set
    if os.getenv("DATABASE_URL"):
        return os.getenv("DATABASE_URL")
    
    host = os.getenv("PGHOST")
    port = os.getenv("PGPORT", "5432")
    user = os.getenv("PGUSER")
    password = os.getenv("PGPASSWORD")
    db = os.getenv("PGDATABASE")
    
    if not all([host, user, password, db]):
        raise ValueError(
            "Set DATABASE_URL or (PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE)"
        )
    
    return f"postgresql://{user}:{password}@{host}:{port}/{db}"


def load_excel_to_db():
    """Load all sheets from Excel to PostgreSQL."""
    # Try multiple paths
    possible_paths = [
        Path("../Bases de datos (1)/Mineria autlán historicos (1).xlsx"),
        Path("../Bases de datos (1)/Mineria autlán historicos.xlsx"),
        Path("../Bases de datos/Mineria autlán historicos.xlsx"),
        Path("/Users/gonzalomoal/Desktop/PYT/Bases de datos (1)/Mineria autlán historicos (1).xlsx"),
    ]
    
    excel_file = None
    for path in possible_paths:
        if path.exists():
            excel_file = path
            break
    
    if excel_file is None:
        print(f"Error: Excel file not found in any of these locations:")
        for p in possible_paths:
            print(f"  - {p}")
        sys.exit(1)
    
    db_url = get_db_url()
    engine = create_engine(db_url, echo=False)
    
    print(f"Connecting to database: {db_url.split('@')[-1]}")
    
    try:
        # Test connection
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("✓ Database connection successful")
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        sys.exit(1)
    
    # Load all sheets
    xls = pd.ExcelFile(excel_file)
    print(f"\nLoading {len(xls.sheet_names)} sheets from {excel_file.name}:")
    
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Sanitize table name (lowercase, replace spaces/special chars with underscore)
            table_name = (
                sheet_name.lower()
                .replace(" ", "_")
                .replace(".", "_")
                .replace("-", "_")
                .replace("ó", "o")
                .replace("á", "a")
                .replace("é", "e")
                .replace("í", "i")
                .replace("ú", "u")
            )
            
            # Drop if exists and recreate
            df.to_sql(table_name, engine, if_exists="replace", index=False)
            print(f"  ✓ {sheet_name:30s} → {table_name:30s} ({len(df)} rows)")
        except Exception as e:
            print(f"  ✗ {sheet_name:30s} failed: {e}")
    
    print("\n✓ All sheets loaded successfully")


if __name__ == "__main__":
    load_excel_to_db()
