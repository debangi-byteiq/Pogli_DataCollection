import pandas as pd
import fitz
import re
import warnings
from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from typing import Optional, List
from pydantic import Field, BaseModel
import os
from BrandRatings import brand
from Description import products
from Customers import customers

os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()

def annualReport_details(doc_text):
    prompt = PromptTemplate(
        """
        You are an expert assistant for extracting company information from documents in JSON format.
        You extract data and return it in JSON format, according to the provided JSON schema, from the given company document.
        Extract the data carefully by analyzing and understanding the provided document.
        REMEMBER to return extracted data only from the provided context.
        CONTEXT:
        {text}
        """
    )

    class CompanyDetails(BaseModel):
        """ Company details """
        company_name: Optional[str] = Field(description="Company Name")
        industry_name: Optional[str] = Field(description="Industry Name")
        cin: Optional[str] = Field(description="CIN (Corporate Identification Number)")
        registered_office_address: Optional[str] = Field(description="Registered Office Address")
        international_operating_centers: Optional[List[str]] = Field(description="List of International Operating Centers")
        national_operating_centers: Optional[List[str]] = Field(description="List of National Operating Centers")
        corporate_address: Optional[str] = Field(description="Corporate Address")
        stock_exchange_listed: Optional[str] = Field(description="Is the company listed on stock exchange?")
        paid_up_share_capital: Optional[str] = Field(description="Paid-up Share Capital")
        assurance_provider: Optional[str] = Field(description="Assurance Provider")
        assurance_type: Optional[str] = Field(description="Assurance Type")
        csr_applicable: Optional[str] = Field(description="CSR Applicable (Yes/No)")
        turnover: Optional[str] = Field(description="Turnover")
        net_worth: Optional[str] = Field(description="Net Worth")
        phone_number: Optional[str] = Field(description="Phone Number of the company")
        email_id: Optional[str] = Field(description="Email ID of the company")
        founder_name: Optional[str] = Field(description="Founder Name")
        founder_designation: Optional[str] = Field(description="Founder Designation")
        foundation_date: Optional[str] = Field(description="Foundation Date")
        revenue: Optional[str] = Field(description="Revenue")
        number_of_workers: Optional[str] = Field(description="Number of Workers")
        number_of_employees: Optional[str] = Field(description="Number of Employees")
        industry_domain: Optional[str] = Field(description="Industry Domain")
        is_listed_on_bse: Optional[str] = Field(description="Is Listed on BSE or Not (Yes/No)")
        company_type: Optional[str] = Field(description="Company Type (Private/Public)")
        complaints_filed: Optional[str] = Field(description="Complaints Filed")
        complaints_pending: Optional[str] = Field(description="Complaints Pending")
        days_of_account_payable: Optional[str] = Field(description="Days of Account Payable")
        website_link: Optional[str] = Field(description="Website Link")
        country_name: Optional[str] = Field(description="Country Name")
        products: Optional[list] = Field(description="List of Products or services offered by the company")
        customers: Optional[list] = Field(description="List of Customers of the Company")
        operating_sites: Optional[list] = Field(description="List of Operating sites or the locations from which the company operates")

    class Details(BaseModel):
        """Company Details"""
        company_details: Optional[CompanyDetails] = Field(description="Details about the company")

    details = extract_pydantic_data(Details, prompt, ' '.join(doc_text))
    return details


def BRSR_details(doc_text):
    prompt = PromptTemplate(
        """
        You are an expert assistant for extracting company information from documents in JSON format.
        Extract the data carefully by analyzing and understanding the provided document.
        REMEMBER to return extracted data only from the provided context.
        CONTEXT:
        {text}
        """
    )

    class CompanyDetails(BaseModel):
        """ Company details """

        company_name: Optional[str] = Field(description="Company Name")
        industry_name: Optional[str] = Field(description="Industry Name")
        year: Optional[int] = Field(description="Reporting Year")
        esg_type: Optional[str] = Field(description="Type of ESG (Environmental, Social, or Governance)")
        esg_category: Optional[str] = Field(description="Category within ESG Type")
        value: Optional[float] = Field(description="Value reported for the ESG parameter")
        values_in_percentage: Optional[float] = Field(description="Values expressed in percentage")
        yes_or_no: Optional[str] = Field(description="Indicator for Yes or No")
        unit: Optional[str] = Field(description="Unit of Measurement (e.g., Days, Percentage, Kilolitres, Giga Joules)")

    class Details(BaseModel):
        """ Company Details """
        company_details: List[CompanyDetails] = Field(description="Details about the company, including ESG data")

    details = extract_pydantic_data(Details, prompt, ' '.join(doc_text))
    return details


def extract_pydantic_data(model, prompt, text, llm=llm):
    program = LLMTextCompletionProgram.from_defaults(
        output_cls=model,
        llm=llm,
        prompt=prompt,
        verbose=True,
    )
    output = program(text=text)
    details = output.model_dump()
    return details


def extract_clean_text_from_pdf(pdf_path):
    print("Cleaning data from the given pdf")
    doc = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
    clean_text = re.sub(r'\s+', ' ', text).strip()
    return clean_text


def update_existing_excel(annualReport_data, BRSR_data, excel_path, company_name, industry_name):
    """
    This function appends data to an existing Excel sheet named "Company Details".
    If the file doesn't exist, it creates a new file with the data.

    Args:
        annualReport_data (dict): The data from annual reports to append, in the format {key: value}.
        BRSR_data (dict): The data from BRSR Report to append, in the format {key: value}.
        excel_path (str): The path to the existing Excel file.

    Returns:
        None
    """
    print("Appending data to Excel")
    print(BRSR_data)
    try:
        # Try to read the existing file
        annual_report_data_existing = pd.read_excel(excel_path, sheet_name='Company Details', engine='openpyxl')
        brsr_data_existing = pd.read_excel(excel_path, sheet_name='Company ESG', engine='openpyxl')
        # Convert the new data into a DataFrame
        annual_report_new = pd.DataFrame([annualReport_data])
        brsr_new = pd.DataFrame([BRSR_data])
        annual_report_new['Company Name'], brsr_new['Company Name'] = company_name
        annual_report_new['Industry Name'], brsr_new['Industry Name'] = industry_name
        # Append the new data to the existing DataFrame
        annual_report_combined = pd.concat([annual_report_data_existing, annual_report_new], ignore_index=True)
        brsr_combined = pd.concat([brsr_data_existing, brsr_new], ignore_index=True)
        # Save the combined DataFrame back to the same sheet
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            annual_report_combined.to_excel(writer, sheet_name='Company Details', index=False)
            brsr_combined.to_excel(writer, sheet_name='Company ESG', index=False)
    except FileNotFoundError:
        print(f"Excel file not found: {excel_path}. Creating a new file.")
        # If the file doesn't exist, create a new DataFrame with the data
        annual_report_new = pd.DataFrame([annualReport_data])
        brsr_new = pd.DataFrame([BRSR_data])
        # Save the DataFrame to a new Excel file
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            annual_report_new.to_excel(writer, sheet_name='Company Details', index=False)
            brsr_new.to_excel(writer, sheet_name='Company ESG', index=False)
    except Exception as e:
        print(f"An error occurred: {e}")


def main():
    warnings.filterwarnings("ignore")
    company_name = ''
    industry_name = ''
    annual_reports_path = "../AnnualReports/india nippon.pdf"
    brsr_path = "../BRSR/indiaNippon_BRSR.pdf"
    excel_path = "../ExcelFiles/pdfData.xlsx"
    glassdoor_link = 'https://www.glassdoor.co.in/Overview/Working-at-Signpost-India-EI_IE2372115.11,25.htm'
    ambitionBox_link = 'https://www.ambitionbox.com/reviews/signpost-india-reviews'
    justDial_link = 'https://www.justdial.com/jdmart/Mumbai/Signpost-India-Pvt-Ltd-Registered-Office-Near-Santacruz-Airport-Terminal-Vile-Parle-East/022PXX22-XX22-181127114814-B8L6_BZDET/catalogue'
    crisil_link = 'https://www.crisilratings.com/en/home/our-business/ratings/company-factsheet.CTODAL.html'
    ticker_link = 'https://ticker.finology.in/company/SIGNPOST'

    print("Starting Annual Reports Scraper")
    annual_report_text = extract_clean_text_from_pdf(annual_reports_path)
    # Extract company details
    annual_report_details = annualReport_details(annual_report_text)

    print("Starting BRSR Scraper")
    brsr_text = extract_clean_text_from_pdf(brsr_path)
    brsr_details = BRSR_details(brsr_text)

    update_existing_excel(annual_report_details['company_details'], brsr_details['company_details'], excel_path, company_name, industry_name)

    # Scraping brand data
    brand(glassdoor_link, ambitionBox_link, justDial_link, crisil_link, ticker_link, company_name, industry_name)

    # Scraping product data
    print("Collecting product data")
    products(annual_report_details['company_details']['products'], company_name, industry_name)

    # Scraping Cutomers data
    print("Collecting customer data")
    customers(annual_report_details['company_details']['customers'], company_name, industry_name)




if __name__ == "__main__":
    main()
