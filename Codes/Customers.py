from typing import Optional, List
from pydantic import BaseModel, Field
from llama_index.llms.gemini import Gemini
from llama_index.core.output_parsers import PydanticOutputParser
import os
import pandas as pd


# Define the Pydantic Model for Company Information
class CompanyInfo(BaseModel):
    foundation_year: Optional[int] = Field(
        description="The year the company was founded",
        default=None
    )
    annual_revenue: Optional[str] = Field(
        description="Average annual revenue in INR",
        default=None
    )
    quarterly_revenue: Optional[str] = Field(
        description="Average quarterly revenue in INR",
        default=None
    )
    sector: Optional[str] = Field(
        description="Primary industry sector of the company",
        default=None
    )

    class Config:
        # Allow extra fields to be more flexible
        extra = 'allow'


def get_company_info(company_names: List[str]) -> List[dict]:
    """
    Fetch general information about multiple companies, including estimated revenue and other details.
    Returns a list of dictionaries, each containing information for one company.
    """
    # Set up the Google API Key
    os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"

    # Initialize Gemini LLM
    llm = Gemini()

    # Create a Pydantic output parser
    class CompanyInfoList(BaseModel):
        companies: List[CompanyInfo]

    output_parser = PydanticOutputParser(output_cls=CompanyInfoList)

    # Adjusted prompt to request information for all companies in the list
    company_list_text = "\n".join([f"- {name}" for name in company_names])
    prompt_text = (
        f"Provide general, estimated information about the following companies:\n\n"
        f"{company_list_text}\n\n"
        "You MUST respond in the following JSON format:\n"
        f"{output_parser.get_format_string()}\n\n"
        "Please include:\n"
        "- Approximate year of foundation\n"
        "- Average annual revenue range (e.g., 100000-200000)\n"
        "- Average quarterly revenue range (e.g., 259079-345726)\n"
        "- Primary industry sector\n"
        "Do NOT provide real-time or constantly changing financial data. Use approximate or typical values."
    )

    try:
        # Generate the response
        response = llm.complete(prompt_text)

        # Parse the response
        parsed_output = output_parser.parse(response.text)

        # Convert list of `CompanyInfo` objects to list of dictionaries
        return [company.dict() for company in parsed_output.companies]

    except Exception as e:
        print(f"Error retrieving information for companies: {e}")
        print(f"Raw response: {response.text}")
        return []  # Return an empty list on failure


def customers(customer: List[str], company_name: str, industry_name: str):
    """
    Process a list of customers, fetch their information in one API call, and save it to an Excel file.
    """
    if not customer:
        print('No customers found')
        return

    # Fetch company information in a batch
    company_info_list = get_company_info(customer)

    if not company_info_list:
        print("Failed to retrieve company information.")
        return

    # Prepare data for DataFrame
    data = {
        'Company Name': company_name,
        'Industry Name': industry_name,
        'Customer': customer,
        'Foundation Year': [info.get("foundation_year", "N/A") for info in company_info_list],
        'Annual Revenue': [info.get("annual_revenue", "N/A") for info in company_info_list],
        'Quarterly Revenue': [info.get("quarterly_revenue", "N/A") for info in company_info_list],
        'Sector': [info.get("sector", "N/A") for info in company_info_list],
    }
    df = pd.DataFrame(data)

    excel_path = '../ExcelFiles/pdfData.xlsx'
    try:
        # Load existing data from the worksheet
        with pd.ExcelFile(excel_path, engine='openpyxl') as excel_file:
            if 'Company Customers' in excel_file.sheet_names:
                # Read existing data
                customer_data_existing = pd.read_excel(excel_path, sheet_name='Company Customers', engine='openpyxl')
                # Combine existing data with new data
                customer_combined = pd.concat([customer_data_existing, df], ignore_index=True)
            else:
                # If the worksheet doesn't exist, initialize combined data with new data
                customer_combined = df

        # Write the combined data back to the same worksheet
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            customer_combined.to_excel(writer, sheet_name='Company Customers', index=False)

    except FileNotFoundError:
        # If the file itself doesn't exist, create it and write the data
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Company Customers', index=False)

    except Exception as e:
        print(f"An error occurred: {e}")

