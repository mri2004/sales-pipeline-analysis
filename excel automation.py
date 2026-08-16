import mysql.connector
import pandas as pd

# connecting to MySql

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="password",
    database="sales_analytics"
)

print("Connected to MySQL!")


queries = {

    # Query 1
    "Pipeline Analysis": """
        SELECT
            deal_stage,
            COUNT(opportunity_id) AS number_of_deals,
            ROUND(AVG(close_value), 2) AS average_deal_value,
            ROUND(SUM(close_value), 2) AS total_deal_value
        FROM sales_pipeline_clean
        GROUP BY deal_stage
        ORDER BY total_deal_value DESC;
    """,

    # Query 2
    "Overall Win Rate": """
        SELECT
            ROUND(
                100.0 * SUM(deal_stage = 'Won') / COUNT(*),
                2
            ) AS win_rate
        FROM sales_pipeline_clean;
    """,

    # Query 3
    "Agent Revenue": """
        SELECT
            sales_agent,
            ROUND(SUM(close_value), 2) AS total_revenue
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY sales_agent
        ORDER BY total_revenue DESC;
    """,

    # Query 4
    "Product Revenue": """
        SELECT
            product,
            ROUND(SUM(close_value), 2) AS total_revenue
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY product
        ORDER BY total_revenue DESC;
    """,

    # Query 5
    "Top 10 Accounts": """
        SELECT
            account,
            ROUND(SUM(close_value), 2) AS total_revenue
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY account
        ORDER BY total_revenue DESC
        LIMIT 10;
    """,

    # Query 6
    "Avg Closing Time": """
        SELECT
            ROUND(
                AVG(DATEDIFF(close_date, engage_date)),
                2
            ) AS avg_days_to_close
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won';
    """,

    # Query 7
    "Agent Avg Deal": """
        SELECT
            sales_agent,
            ROUND(AVG(close_value), 2) AS avg_deal_value
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY sales_agent
        ORDER BY avg_deal_value DESC;
    """,

    # Query 8
    "Product Win Rate": """
        SELECT
            product,
            ROUND(
                100.0 * SUM(deal_stage = 'Won') / COUNT(*),
                2
            ) AS win_rate
        FROM sales_pipeline_clean
        GROUP BY product
        ORDER BY win_rate DESC;
    """,

    # Query 9
    "Account Opportunities": """
        SELECT
            account,
            COUNT(opportunity_id) AS total_opportunities
        FROM sales_pipeline_clean
        GROUP BY account
        ORDER BY total_opportunities DESC;
    """,

    # Query 10
    "Sector Revenue": """
        SELECT
            a.sector,
            ROUND(SUM(sp.close_value), 2) AS total_revenue
        FROM accounts AS a
        INNER JOIN sales_pipeline_clean AS sp
            ON a.account = sp.account
        WHERE sp.deal_stage = 'Won'
        GROUP BY a.sector
        ORDER BY total_revenue DESC;
    """,

    # Query 11
    "Office Revenue": """
        SELECT
            st.regional_office,
            ROUND(SUM(sp.close_value), 2) AS total_revenue
        FROM sales_teams AS st
        INNER JOIN sales_pipeline_clean AS sp
            ON st.sales_agent = sp.sales_agent
        WHERE sp.deal_stage = 'Won'
        GROUP BY st.regional_office
        ORDER BY total_revenue DESC;
    """,

    # Query 12
    "Agent Performance": """
        SELECT
            sales_agent,
            COUNT(opportunity_id) AS total_opportunities,
            SUM(deal_stage = 'Won') AS won_opportunities,
            ROUND(
                100.0 * SUM(deal_stage = 'Won') / COUNT(*),
                2
            ) AS win_rate
        FROM sales_pipeline_clean
        GROUP BY sales_agent
        ORDER BY win_rate DESC;
    """,

    # Query 13
    "Monthly Revenue": """
        SELECT
            YEAR(close_date) AS year,
            MONTH(close_date) AS month,
            ROUND(SUM(close_value), 2) AS total_revenue
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY YEAR(close_date), MONTH(close_date)
        ORDER BY YEAR(close_date), MONTH(close_date);
    """,

    # Query 14
    "Agent Closing Time": """
        SELECT
            sales_agent,
            ROUND(
                AVG(DATEDIFF(close_date, engage_date)),
                2
            ) AS avg_days_to_close
        FROM sales_pipeline_clean
        WHERE deal_stage = 'Won'
        GROUP BY sales_agent
        ORDER BY avg_days_to_close ASC;
    """
}


# creating excel file

output_file = "B2B_Sales_Automated_Report.xlsx"

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    for sheet_name, query in queries.items():

        df = pd.read_sql(query, conn)

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

        print(f"Created sheet: {sheet_name}")


conn.close()

print("Excel report created successfully!")
