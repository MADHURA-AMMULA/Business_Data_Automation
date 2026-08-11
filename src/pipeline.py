from pathlib import Path
import pandas as pd
import logging




BASE_DIR = Path(__file__).resolve().parent.parent

RAW_FILE = BASE_DIR / "data" / "raw" / "Sample - Superstore.csv"
PROCESSED_DIR = BASE_DIR / "data" / "processed"
OUTPUT_DIR = BASE_DIR / "output"

CLEAN_FILE = PROCESSED_DIR / "superstore_clean.csv"
EXCEL_FILE = PROCESSED_DIR / "superstore_analysis.xlsx"
LOG_FILE = OUTPUT_DIR / "pipeline.log"




PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)




logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger(__name__)




def load_data():
    logger.info("Loading raw data...")

    df = pd.read_csv(
        RAW_FILE,
        encoding="latin1"
    )

    logger.info(f"Loaded {len(df)} rows and {len(df.columns)} columns.")

    return df




def clean_data(df):

    logger.info("Starting data cleaning...")

    df = df.copy()


    before = len(df)

    df = df.drop_duplicates()

    removed = before - len(df)

    logger.info(f"Removed {removed} duplicate rows.")


    df["Order Date"] = pd.to_datetime(
        df["Order Date"],
        errors="coerce"
    )

    df["Ship Date"] = pd.to_datetime(
        df["Ship Date"],
        errors="coerce"
    )

    
    numeric_columns = [
        "Sales",
        "Quantity",
        "Discount",
        "Profit"
    ]

    for column in numeric_columns:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

    
    text_columns = [
        "Order ID",
        "Ship Mode",
        "Customer ID",
        "Customer Name",
        "Segment",
        "Country",
        "City",
        "State",
        "Region",
        "Product ID",
        "Category",
        "Sub-Category",
        "Product Name"
    ]

    for column in text_columns:
        df[column] = df[column].astype("string").str.strip()


    df["Shipping Days"] = (
        df["Ship Date"] - df["Order Date"]
    ).dt.days

    
    df["Profit Margin"] = (
        df["Profit"] / df["Sales"]
    )

    
    df["Profit Margin"] = (
        df["Profit Margin"]
        .replace([float("inf"), -float("inf")], 0)
        .fillna(0)
    )


    df["Order Year"] = df["Order Date"].dt.year
    df["Order Month"] = df["Order Date"].dt.month
    df["Order Month Name"] = df["Order Date"].dt.month_name()


    df["Profit Margin"] = df["Profit Margin"].round(4)

    logger.info("Data cleaning completed.")

    return df




def validate_data(df):

    logger.info("Validating processed data...")

    required_columns = [
        "Order ID",
        "Order Date",
        "Sales",
        "Quantity",
        "Discount",
        "Profit"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Missing required columns: {missing_columns}"
        )

    if df["Order ID"].isna().any():
        raise ValueError("Order ID contains missing values.")

    if df["Sales"].isna().any():
        raise ValueError("Sales contains missing values.")

    if df["Profit"].isna().any():
        raise ValueError("Profit contains missing values.")

    if (df["Quantity"] <= 0).any():
        raise ValueError("Invalid quantity detected.")

    logger.info("Validation successful.")




def calculate_kpis(df):

    logger.info("Calculating KPIs...")

    total_sales = df["Sales"].sum()

    total_profit = df["Profit"].sum()

    total_quantity = df["Quantity"].sum()

    total_orders = df["Order ID"].nunique()

    total_customers = df["Customer ID"].nunique()

    average_order_value = (
        total_sales / total_orders
    )

    profit_margin = (
        total_profit / total_sales
    )

    kpis = {
        "Total Sales": total_sales,
        "Total Profit": total_profit,
        "Total Quantity": total_quantity,
        "Total Orders": total_orders,
        "Total Customers": total_customers,
        "Average Order Value": average_order_value,
        "Profit Margin": profit_margin
    }

    return kpis




def category_analysis(df):

    result = (
        df.groupby("Category")
        .agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("Order ID", "nunique")
        )
        .reset_index()
    )

    result["Profit Margin"] = (
        result["Profit"] / result["Sales"]
    )

    return result.sort_values(
        "Sales",
        ascending=False
    )




def region_analysis(df):

    result = (
        df.groupby("Region")
        .agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("Order ID", "nunique")
        )
        .reset_index()
    )

    result["Profit Margin"] = (
        result["Profit"] / result["Sales"]
    )

    return result.sort_values(
        "Sales",
        ascending=False
    )




def top_products(df):

    result = (
        df.groupby(
            ["Product ID", "Product Name"]
        )
        .agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Quantity=("Quantity", "sum")
        )
        .reset_index()
    )

    return result.sort_values(
        "Sales",
        ascending=False
    ).head(10)



def monthly_analysis(df):

    result = (
        df.groupby(
            ["Order Year", "Order Month"]
        )
        .agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Quantity=("Quantity", "sum")
        )
        .reset_index()
    )

    return result.sort_values(
        ["Order Year", "Order Month"]
    )



def save_processed_data(df):

    logger.info("Saving processed CSV...")

    df.to_csv(
        CLEAN_FILE,
        index=False
    )

    logger.info(
        f"Processed data saved to {CLEAN_FILE}"
    )




def export_excel(
    df,
    kpis,
    category_df,
    region_df,
    product_df,
    monthly_df
):

    logger.info("Creating Excel report...")

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl"
    ) as writer:


        kpi_df = pd.DataFrame(
            list(kpis.items()),
            columns=["KPI", "Value"]
        )

        kpi_df.to_excel(
            writer,
            sheet_name="KPIs",
            index=False
        )

      
        category_df.to_excel(
            writer,
            sheet_name="Category Analysis",
            index=False
        )

      
        region_df.to_excel(
            writer,
            sheet_name="Region Analysis",
            index=False
        )


        product_df.to_excel(
            writer,
            sheet_name="Top Products",
            index=False
        )

        
        monthly_df.to_excel(
            writer,
            sheet_name="Monthly Analysis",
            index=False
        )

        
        df.to_excel(
            writer,
            sheet_name="Clean Data",
            index=False
        )

    logger.info(
        f"Excel report saved to {EXCEL_FILE}"
    )



def main():

    logger.info("=" * 60)
    logger.info("SUPERSTORE AUTOMATION PIPELINE STARTED")
    logger.info("=" * 60)

    
    df = load_data()

   
    df = clean_data(df)

   
    validate_data(df)

   
    kpis = calculate_kpis(df)

    
    category_df = category_analysis(df)

    region_df = region_analysis(df)

    product_df = top_products(df)

    monthly_df = monthly_analysis(df)

   
    save_processed_data(df)

   
    export_excel(
        df,
        kpis,
        category_df,
        region_df,
        product_df,
        monthly_df
    )

    logger.info("Pipeline completed successfully.")

    print("\nPipeline completed successfully!")
    print(f"Rows processed: {len(df):,}")
    print(f"Total Sales: ${kpis['Total Sales']:,.2f}")
    print(f"Total Profit: ${kpis['Total Profit']:,.2f}")
    print(f"Total Orders: {kpis['Total Orders']:,}")
    print(f"Total Customers: {kpis['Total Customers']:,}")
    print(f"Excel output: {EXCEL_FILE}")


if __name__ == "__main__":
    main()