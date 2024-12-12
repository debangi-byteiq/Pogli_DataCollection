from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from typing import Optional, List
from pydantic import Field, BaseModel
import os
import pandas as pd
import fitz
import re


os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()

def get_company_details(doc_text):
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
        phone_number: Optional[str] = Field(description="Phone Number")
        email_id: Optional[str] = Field(description="Email ID")
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
        products: Optional[str] = Field(description="Products or services offered by the company")
        customers: Optional[str] = Field(description="Customers of the Company")
        operating_sites: Optional[str] = Field(description="List of Operating sites or the locations from which the company operates")

    class Details(BaseModel):
        """Company Details"""
        company_details: Optional[CompanyDetails] = Field(description="Details about the company")

    details = extract_pydantic_data(Details, prompt, ' '.join(doc_text))
    return details


def extract_pydantic_data(model, prompt, text,llm=llm):
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
    # Open the provided PDF file
    doc = fitz.open(pdf_path)

    # Initialize an empty string to hold all text
    text = ""

    # Iterate through each page in the document
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)  # Load the page
        text += page.get_text()  # Extract text from the page

    # Remove extra spaces, line breaks, and non-essential formatting
    clean_text = re.sub(r'\s+', ' ', text).strip()  # Replace multiple spaces and newlines with a single space
    return clean_text


def main():
    company_name = input("Enter Company Name")
    pdf_path = input("Enter PDF Path")

    cleaned_text = extract_clean_text_from_pdf(pdf_path)
    details = get_company_details(cleaned_text)
    print(details)

    # Creating the DataFrame with keys in 'Entity' and values in 'Value'
    keys = list(details['company_details'].keys())
    values = list(details['company_details'].values())
    df = pd.DataFrame({
        'Entity': keys,
        'Value': values
    })
    df = df.transpose()

    # Save the DataFrame to an Excel file
    df.to_excel(f'{company_name}.xlsx', index=False)


if __name__ == "__main__":
    main()
