from llama_index.llms.gemini import Gemini
from llama_index.core.prompts import PromptTemplate
from llama_index.core.program import LLMTextCompletionProgram
from pydantic import Field, BaseModel
import os
import pandas as pd

os.environ["GOOGLE_API_KEY"] = "AIzaSyAjP37AbKfS7gHyy72DkQDXckP5FBIRwto"
llm = Gemini()


def get_batch_descriptions(products):
    """
    Fetch short descriptions for a batch of products in one API call and return them as a dictionary.
    """
    prompt = PromptTemplate(
        """
        You are an expert assistant for giving short descriptions about the terms or products that are provided to you.
        You understand the terms and the basis they are given on and provide short summarized descriptions in 30-40 words.
        Ensure each product is clearly identified with a matching description.
        CONTEXT:
        {text}
        """
    )

    class Description(BaseModel):
        descriptions: list[str] = Field(
            description="List of short descriptions for the provided products"
        )

    # Number the products for clarity
    product_context = "\n".join([f"{i+1}. {product}" for i, product in enumerate(products)])

    details = extract_pydantic_data(Description, prompt, product_context)

    # Validate and ensure the return type is correct
    if not isinstance(details, dict) or "descriptions" not in details:
        raise ValueError("Invalid response format from extract_pydantic_data.")

    returned_descriptions = details["descriptions"]

    # Ensure descriptions align with the product list
    if len(products) != len(returned_descriptions):
        print("Warning: Mismatch in the number of descriptions. Using placeholders for missing items.")
        returned_descriptions = [
            desc if i < len(returned_descriptions) else "Description unavailable"
            for i, desc in enumerate(products)
        ]

    # Return descriptions as a dictionary
    return dict(zip(products, returned_descriptions))


def extract_pydantic_data(model, prompt, text, llm=llm):
    """
    Extract structured data from the LLM using Pydantic models.
    """
    program = LLMTextCompletionProgram.from_defaults(
        output_cls=model,
        llm=llm,
        prompt=prompt,
        verbose=True,
    )
    output = program(text=text)
    details = output.model_dump()
    return details


def products(product_list, company_name, industry_name):
    """
    Process a list of products, fetch their descriptions in one API call, and save the data to an Excel file.
    """
    if not product_list:
        print("No products found")
        return

    try:
        # Fetch descriptions for all products in a single call
        product_descriptions = get_batch_descriptions(product_list)

        # Prepare data for DataFrame
        data = {
            "Company Name": company_name,
            "Industry Name": industry_name,
            "Products": list(product_descriptions.keys()),
            "Description": list(product_descriptions.values()),
        }
        df = pd.DataFrame(data)

        # Path to the Excel file
        excel_path = "../ExcelFiles/pdfData.xlsx"

        # Write or append the data to the Excel file
        try:
            with pd.ExcelFile(excel_path, engine="openpyxl") as excel_file:
                if "Company Products" in excel_file.sheet_names:
                    # Read existing data
                    product_data_existing = pd.read_excel(
                        excel_path, sheet_name="Company Products", engine="openpyxl"
                    )
                    # Combine existing data with new data
                    product_combined = pd.concat([product_data_existing, df], ignore_index=True)
                else:
                    # If the worksheet doesn't exist, initialize combined data with new data
                    product_combined = df

            # Write the combined data back to the same worksheet
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                product_combined.to_excel(writer, sheet_name="Company Products", index=False)

        except FileNotFoundError:
            # If the file itself doesn't exist, create it and write the data
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, sheet_name="Company Products", index=False)

        except Exception as e:
            print(f"An error occurred while saving to Excel: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")
