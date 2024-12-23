# import pandas as pd
# import fitz
# import re
# import warnings
# from llama_index.llms.gemini import Gemini
# from llama_index.core.prompts import PromptTemplate
# from llama_index.core.program import LLMTextCompletionProgram
# from typing import Optional, List
# from pydantic import Field, BaseModel
# import os
#
# def extract_clean_text_from_pdf_by_chunks(pdf_path, chunk_size=5):
#     """
#     Extracts text from a PDF file in chunks of pages.
#
#     Args:
#         pdf_path (str): Path to the PDF file.
#         chunk_size (int): Number of pages per chunk.
#
#     Returns:
#         List[str]: List of text chunks from the PDF.
#     """
#     print("Extracting data from the PDF in chunks")
#     doc = fitz.open(pdf_path)
#     chunks = []
#     chunk_text = ""
#     for page_num in range(len(doc)):
#         page = doc.load_page(page_num)
#         chunk_text += page.get_text()
#         if (page_num + 1) % chunk_size == 0 or page_num == len(doc) - 1:
#             chunks.append(re.sub(r'\s+', ' ', chunk_text).strip())
#             chunk_text = ""
#     return chunks
#
#
# def process_pdf_chunks(chunks, process_function):
#     """
#     Processes each chunk of text from the PDF and merges the results.
#
#     Args:
#         chunks (List[str]): List of text chunks to process.
#         process_function (Callable): Function to process a chunk.
#
#     Returns:
#         dict: Merged data from all chunks.
#     """
#     combined_data = {}
#     for i, chunk in enumerate(chunks):
#         print(f"Processing chunk {i + 1} of {len(chunks)}")
#         chunk_data = process_function(chunk)
#         # Merge dictionaries, updating keys or combining lists where necessary
#         for key, value in chunk_data.items():
#             if key not in combined_data:
#                 combined_data[key] = value
#             elif isinstance(value, list):
#                 combined_data[key].extend(value)
#             else:
#                 combined_data[key] = value
#     return combined_data
#
#
# def annualReport_details(doc_text):
#     """
#     Extracts details from a single chunk of the annual report.
#     Modify as per the updated schema or requirements.
#     """
#     prompt = PromptTemplate(
#         """
#         Extract the company details in JSON format based on the following schema:
#         {text}
#         """
#     )
#
#     class CompanyDetails(BaseModel):
#         # Define your fields here
#         pass
#
#     details = extract_pydantic_data(CompanyDetails, prompt, doc_text)
#     return details
#
#
# def query_combined_data(data, query_field):
#     """
#     Queries the combined data to extract specific fields.
#
#     Args:
#         data (dict): Combined data from all chunks.
#         query_field (str): Field to query.
#
#     Returns:
#         Any: Value of the queried field.
#     """
#     return data.get(query_field, None)
#
#
# def main():
#     warnings.filterwarnings("ignore")
#     company_name = 'TVS ELECTRONICS LTD'
#     industry_name = 'Computers Hardware & Equipments'
#     annual_reports_path = "../AnnualReports/IFGL_AnnRep.pdf"
#     excel_path = "../ExcelFiles/pdfData.xlsx"
#
#     print("Starting Annual Reports Scraper")
#     chunks = extract_clean_text_from_pdf_by_chunks(annual_reports_path, chunk_size=5)
#     annual_report_data = process_pdf_chunks(chunks, annualReport_details)
#
#     # Example: Query specific fields
#     turnover = query_combined_data(annual_report_data, 'turnover')
#     print(f"Turnover: {turnover}")
#
#     # Save or further process the data
#     update_existing_excel(annual_report_data, {}, excel_path, company_name, industry_name)
#
# # ########################################################3
# from pydantic import ValidationError
#
#
# def extract_pydantic_data(model, prompt, text, llm=llm, max_chunk_size=1000):
#     """
#     Extract data using Pydantic and LLMTextCompletionProgram with token limitations in mind.
#
#     Args:
#         model: Pydantic model for validation.
#         prompt: Template string for the prompt.
#         text: Input text to process.
#         llm: LLM instance.
#         max_chunk_size: Maximum size of text chunks for processing (default: 1000).
#
#     Returns:
#         Combined details from processed chunks.
#     """
#     from llama_index.llms import LLMTextCompletionProgram
#
#     # Initialize the program
#     program = LLMTextCompletionProgram.from_defaults(
#         output_cls=model,
#         llm=llm,
#         prompt=prompt,
#         verbose=True,
#     )
#
#     details = []
#     # Split text into manageable chunks
#     chunks = [text[i: i + max_chunk_size] for i in range(0, len(text), max_chunk_size)]
#
#     for idx, chunk in enumerate(chunks):
#         try:
#             print(f"Processing chunk {idx + 1}/{len(chunks)}")
#             output = program(text=chunk)  # Run the program for each chunk
#             details.append(output.model_dump())  # Validate and store the processed data
#         except ValidationError as e:
#             print(f"Validation error in chunk {idx + 1}: {e}")
#         except Exception as e:
#             print(f"An error occurred while processing chunk {idx + 1}: {e}")
#
#     # Combine results from all chunks
#     combined_details = {}
#     for detail in details:
#         combined_details.update(detail)  # Combine dictionaries
#
#     return combined_details
# if __name__ == "__main__":
#     main()


from llama_index.core.node_parser import SentenceSplitter
from llama_index.core import Document

text_splitter = SentenceSplitter(chunk_size=100, chunk_overlap=10)

sample_text = '''
Lorem ipsum dolor sit amet, consectetur adipiscing elit. Mauris ultrices diam at feugiat pellentesque. Donec vel tortor ipsum. Nunc malesuada gravida lectus vel convallis. Vestibulum posuere ante dui, sit amet commodo mi tempor vitae. Sed blandit convallis nunc aliquam vulputate. Maecenas cursus dolor at metus egestas mattis. Fusce pretium blandit risus, non finibus arcu rutrum ut. Fusce eu volutpat nunc. Praesent bibendum placerat mauris, at eleifend eros. Fusce commodo arcu eget convallis varius.

Fusce vitae ipsum gravida, placerat purus nec, viverra mi. Suspendisse porta mauris ac semper facilisis. Vestibulum viverra in turpis eu pharetra. Nam pellentesque erat lectus, sollicitudin tempor ligula pharetra et. Aliquam erat volutpat. Curabitur ut urna in orci semper convallis sed nec dolor. Vestibulum id ornare ex, ut vestibulum dolor. Vivamus condimentum turpis quam, in sodales mi pulvinar at. Duis rutrum dignissim diam, ut pellentesque arcu eleifend sit amet. Nulla eu elit vel enim fringilla tempor. Maecenas egestas lacus sit amet scelerisque euismod. Nullam id eros vitae urna bibendum consectetur eu ut tellus. Ut faucibus quis lacus non convallis. Quisque maximus, urna vel pulvinar malesuada, arcu nisi cursus dolor, nec lacinia ex libero vitae odio. Suspendisse egestas sapien enim, quis tincidunt quam finibus eu. Suspendisse vitae neque.
'''

chunks = text_splitter.get_nodes_from_documents(
    [Document(text=sample_text)],
    show_progress=False
)

print(type(chunks))
# print(chunks)
for idx, chunk in enumerate(chunks):
    print(idx,chunk.text, end='\n===========================================================================================================================================================')
