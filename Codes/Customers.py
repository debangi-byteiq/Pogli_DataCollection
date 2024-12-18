from typing import Optional
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
        description="Average annual revenue in USD",
        default=None
    )
    quarterly_revenue: Optional[str] = Field(
        description="Average quarterly revenue in USD",
        default=None
    )
    sector: Optional[str] = Field(
        description="Primary industry sector of the company",
        default=None
    )

    class Config:
        # Allow extra fields to be more flexible
        extra = 'allow'


def get_company_info(company_name: str) -> CompanyInfo:
    # Set up the Google API Key
    os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"

    # Initialize Gemini LLM
    llm = Gemini()

    # Create a Pydantic output parser
    output_parser = PydanticOutputParser(output_cls=CompanyInfo)

    # Create a detailed prompt template
    prompt_text = (
        f"Provide comprehensive information about {company_name}:\n\n"
        "You MUST respond in the following JSON format:\n"
        f"{output_parser.get_format_string()}\n\n"
        "Please ensure all information is as accurate and current as possible.\n"
        "Include details about the company's:\n"
        "- Exact year of foundation\n"
        "- Average annual revenue of the company\n"
        "- Average quarterly revenue of the company\n"
        "- Primary industry sector\n"
    )

    try:
        # Generate the response directly
        response = llm.complete(prompt_text)

        # Parse the response
        parsed_output = output_parser.parse(response.text)

        return parsed_output

    except Exception as e:
        print(f"Error retrieving information for {company_name}: {e}")
        print(f"Raw response: {response.text}")
        return CompanyInfo()  # Return an empty CompanyInfo object


def customers(customer, company_name, industry_name):
    keys = list()
    foundation_year = list()
    ann_rev = list()
    quart_rev = list()
    sector = list()
    if len(customer) != 0:
        for each in customer:
            company_info = get_company_info(each)
            keys.append(each)
            foundation_year.append(company_info.foundation_year)
            ann_rev.append(company_info.annual_revenue)
            quart_rev.append(company_info.quarterly_revenue)
            sector.append(company_info.sector)
            df = pd.DataFrame({'Company Name': company_name, 'Industry Name': industry_name,'Customer': keys, 'Foundation Year': foundation_year, 'Annual Revenue': ann_rev,
                               'Quarterly Revenue': quart_rev, 'Sector': sector})
            df['company_name'] = company_name
            df['industry_name'] = industry_name
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
                with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    customer_combined.to_excel(writer, sheet_name='Company Customers', index=False)

            except FileNotFoundError:
                # If the file itself doesn't exist, create it and write the data
                with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                    df.to_excel(writer, sheet_name='Company Customers', index=False)

            except Exception as e:
                print(f"An error occurred: {e}")
    else:
        print('No customers found')



