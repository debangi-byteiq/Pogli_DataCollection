from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from pydantic import Field, BaseModel
import os
import pandas as pd

os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()


def get_description(products):
    prompt = PromptTemplate(
        """
        You are an expert assistant for giving short descriptions about the terms or products that are provided to you.
        You understand the terms and the basis they are given on and provide short summarised descriptions in 30-40 words.
        REMEMBER to return extracted data only from the provided context.
        CONTEXT:
        {text}
        """
    )

    class Description(BaseModel):
        description: str = Field(description="Short description of the product")

    details = extract_pydantic_data(Description, prompt, ' '.join(products))

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


def products(product_list):
    # product_list = ['Wireless front end', 'Satellite Communication', 'Embedded systems', 'Signal Processing', 'Network management', 'Software development']
    keys = list()
    values = list()
    for each in product_list:
        print(each)
        description = get_description(each)
        print(f"{each}: {description['description']}")
        keys.append(each)
        values.append(description['description'])

    # Creating the DataFrame with keys in 'Entity' and values in 'Value'
    df = pd.DataFrame({'Products': keys, 'Description': values})
    excel_path = '../Excel Files/pdfData.xlsx'
    product_combined = pd.DataFrame()
    try:
        # Attempt to read the existing worksheet
        with pd.ExcelFile(excel_path, engine='openpyxl') as excel_file:
            if 'Company Products' in excel_file.sheet_names:
                product_data_existing = pd.read_excel(excel_path, sheet_name='Company Products', engine='openpyxl')
                # Combine existing and new data
                product_combined = pd.concat([product_data_existing, df], ignore_index=True)
            else:
                # If the sheet doesn't exist, initialize combined data with new data
                product_combined = df

        # Write combined data back to Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
            product_combined.to_excel(writer, sheet_name='Company Products', index=False)

    except FileNotFoundError:
        # If the file itself doesn't exist, create it
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Company Products', index=False)


def main():
    product_list = ['Wireless front end', 'Satellite Communication', 'Embedded systems', 'Signal Processing', 'Network management', 'Software development']
    products(product_list)

if __name__ == '__main__':
    main()