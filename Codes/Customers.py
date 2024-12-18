from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
import os
import re

os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()

def get_company_info(company_name):
    prompt = PromptTemplate(
        f"""
        Provide the following information about the company {company_name}:
        1. Foundation year
        2. Annual revenue
        3. Quarterly revenue
        4. Sector
        """
    )

    program = LLMTextCompletionProgram.from_defaults(
        llm=llm,
        prompt=prompt
    )

    response = program(text="")
    response_text = response.text
    info = response_text
    print(info)

    # # Improved text parsing using regular expressions
    # info = {}
    # pattern = r"(\d+\.\s*)(.*)"
    # for line in response_text.split('\n'):
    #     match = re.match(pattern, line)
    #     if match:
    #         key, value = match.groups()
    #         info[key.strip('.')] = value.strip()

    return info

# Example usage
companies = ["Google", "Microsoft", "Apple"]
for company in companies:
    company_info = get_company_info(company)
    print(f"{company}:")
    print(company_info)
    # print(f"  Foundation Year: {company_info['1']}")
    # print(f"  Annual Revenue: {company_info['2']}")
    # print(f"  Quarterly Revenue: {company_info['3']}")
    # print(f"  Sector: {company_info['4']}")
