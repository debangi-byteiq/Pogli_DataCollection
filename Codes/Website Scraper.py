from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from pydantic import BaseModel, Field
import requests
import re
import os
import pandas as pd
from typing import Optional
from bs4 import BeautifulSoup

# Set up LLM (Google Gemini)
os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()

def get_company_details(urls):
    soup = []
    for url in urls:
        # Fetch website content
        response = requests.get(url)
        response.raise_for_status()  # Raise exception for non-200 status codes
        html_content = response.text

        # Parse HTML content using BeautifulSoup
        soup.append(BeautifulSoup(html_content, 'html.parser'))

    # Extract relevant text from the parsed HTML
    # This step will depend on the specific website structure
    # Here's a basic example:
    text_content = '||'.join(list(each.get_text() for each in soup))
    # print(text_content)

    # Clean the text (optional, depending on the website's structure)
    cleaned_text = re.sub(r'\s+', ' ', text_content).strip()

    prompt = PromptTemplate(
        """
        You are an expert assistant for extracting company information from a list of websites in JSON format.
        Extract the data carefully by analyzing and understanding the provided website content including the text on page and image contents.
        REMEMBER to return extracted data only from the provided context.
        CONTEXT:
        {text}
        """
    )

    class CompanyDetails(BaseModel):
        """ Company details """

        company_name: Optional[str] = Field(description="Company Name")
        industry_name: Optional[str] = Field(description="Industry to which the company belongs")
        products: Optional[list] = Field(description="List of Name of Products or their capabilities and services offered by the company")
        customers: Optional[list] = Field(description="List of Name of the Costumer or clients of the given company or who they serve their services to")
        operating_sites: Optional[list] = Field(description="List of the Name of the sites , Location and number of employees in each operating sites")
        key_personnels: Optional[list] = Field(description="List of Key personnels in the company along with their name and designation")

    class Details(BaseModel):
        """ Company Details """
        company_details: list[CompanyDetails] = Field(description="Details about the company, including ESG data")

    details = extract_pydantic_data(Details, prompt, cleaned_text)
    return details


def extract_pydantic_data(model, prompt, text, llm =llm):
    program = LLMTextCompletionProgram.from_defaults(
        output_cls=model,
        llm=llm,
        prompt=prompt,
        verbose=True,
    )
    output = program(text=text)
    details = output.model_dump()
    return details


def main():
    company_name = input("Enter Company Name: ")
    # urls = ['https://www.igarashimotors.com/international-customers.php',
    #         'https://www.igarashimotors.com/domestic-customers.php',
    #         'https://www.igarashimotors.com/index.php',
    #         'https://www.igarashimotors.com/investor-relations.php']

    urls = ['https://signpostindia.com/captura/',
            'https://signpostindia.com/about-us-who-we-are/',
            'https://signpostindia.com/']

    details = get_company_details(urls)

    df = pd.DataFrame(details['company_details'])
    df.to_excel(f'{company_name}.xlsx', index=False)
    print(f"Details saved to {company_name}.xlsx")

if __name__ == "__main__":
    main()
