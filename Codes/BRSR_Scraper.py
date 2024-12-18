from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from pydantic import BaseModel, Field
import fitz
import re
import os
import pandas as pd
from typing import Optional

# Set up LLM (Google Gemini)
os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()

def extract_clean_text_from_pdf(pdf_path):
    '''
    Function to extract clean text from PDF
    :param pdf_path: path to pdf file
    :return: cleaned text
    '''
    doc = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()
    clean_text = re.sub(r'\s+', ' ', text).strip()
    return clean_text


def get_company_details(doc_text):
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
        company_id: Optional[str] = Field(description="Unique Identifier for the Company")
        year: Optional[int] = Field(description="Reporting Year")
        esg_type: Optional[str] = Field(description="Type of ESG (Environmental, Social, or Governance)")
        esg_category: Optional[str] = Field(description="Category within ESG Type")
        value: Optional[float] = Field(description="Value reported for the ESG parameter")
        values_in_percentage: Optional[float] = Field(description="Values expressed in percentage")
        yes_or_no: Optional[str] = Field(description="Indicator for Yes or No")
        unit: Optional[str] = Field(description="Unit of Measurement (e.g., Days, Percentage, Kilolitres, Giga Joules)")

    class Details(BaseModel):
        """ Company Details """
        company_details: list[CompanyDetails] = Field(description="Details about the company, including ESG data")

    details = extract_pydantic_data(Details, prompt, ' '.join(doc_text))
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
    company_name = input("Enter Company Name")
    # Load PDF, extract text, and get details
    pdf_path = input("Enter PDF Path")
    cleaned_text = extract_clean_text_from_pdf(pdf_path)
    details = get_company_details(cleaned_text)
    print(details)

    df = pd.DataFrame([details['company_details']])
    # Save to CSV
    df.to_csv(f'{company_name}.csv', index=False)
    print(f"Details saved to {company_name}.csv")


if __name__ == "__main__":
    main()
