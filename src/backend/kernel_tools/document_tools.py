import inspect
from typing import Annotated, Callable, List

import pandas as pd
from bs4 import BeautifulSoup
from io import StringIO

from semantic_kernel.functions import kernel_function
from models.messages_kernel import AgentType
import json
from typing import get_type_hints


class DocumentTools:
    # Define Document tools (functions)
    formatting_instructions = "Instructions: You help read documents, extract information, and save extracted table inforamtion to an excel file. " \
    agent_name = AgentType.HR.value

    @staticmethod
    @kernel_function(description="Read contents of the data provided, extract tables and return them on HTML format.")
    async def read_document_contents(document: bytes ) -> str:
        #"placeholder for reading document contents and extracting tables"
        html_content= """
        Here are the extracted tables in HTML format: 1. **Sales Numbers for Q1:** <table border='1' style='border-collapse:collapse;width:50%;'> <thead> <tr> <th>Quarter</th> <th>Bicycle Sales</th> <th>Helmet Sales</th> <th>Total Sales</th> </tr> </thead> <tbody> <tr> <td>Q1</td> <td>$150,000</td> <td>$75,000</td> <td>$225,000</td> </tr> <tr> <td>Q2</td> <td>$180,000</td> <td>$90,000</td> <td>$270,000</td> </tr> <tr> <td>Q3</td> <td>$210,000</td> <td>$105,000</td> <td>$315,000</td> </tr> <tr> <td>Q4</td> <td>$240,000</td> <td>$120,000</td> <td>$360,000</td> </tr> </tbody></table>2. **Projected Sales Numbers:**<table border='1' style='border-collapse:collapse;width:50%;'> <thead> <tr> <th>Quarter</th> <th>Projected Bicycle Sales</th> <th>Projected Helmet Sales</th> <th>Projected Total Sales</th> </tr> </thead> <tbody> <tr> <td>Q1</td> <td>$270,000</td> <td>$135,000</td> <td>$405,000</td> </tr> <tr> <td>Q2</td> <td>$300,000</td> <td>$150,000</td> <td>$450,000</td> </tr> <tr> <td>Q3</td> <td>$330,000</td> <td>$165,000</td> <td>$495,000</td> </tr> <tr> <td>Q4</td> <td>$360,000</td> <td>$180,000</td> <td>$540,000</td> </tr> </tbody></table>"
        """
        return (html_content)
            
    @staticmethod
    @kernel_function(description="Save the extracted HTML tables to an Excel file.")
    async def save_to_excel(html_content: str) -> bytes:
            """
        Converts HTML tables in the provided HTML content to an Excel file.
        
        Args:
            html_content (str): The HTML content containing tables.
        
        Returns:
            bytes: The function saves the Excel file as 'output.xlsx'.
    """

        # Parse the HTML content
        soup = BeautifulSoup(html_content, 'html.parser')

        # Find all tables in the HTML
        tables = soup.find_all('table')

        # Create a Pandas Excel writer using XlsxWriter as the engine
        excel_writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

        # Iterate over the tables and convert each to a DataFrame
        for i, table in enumerate(tables):
            # Read the HTML table into a DataFrame
            df = pd.read_html(StringIO(str(table)))[0]
            # Write the DataFrame to a specific sheet in the Excel file
            df.to_excel(excel_writer, sheet_name=f'Table_{i+1}', index=False)

        # Save the Excel file
        excel_writer.close()
        # Get the bytes of the Excel file
        with open('output.xlsx', 'rb') as f:
            excel_bytes = f.read()
        return excel_bytes

    @classmethod
    def get_all_kernel_functions(cls) -> dict[str, Callable]:
        """
        Returns a dictionary of all methods in this class that have the @kernel_function annotation.
        This function itself is not annotated with @kernel_function.

        Returns:
            Dict[str, Callable]: Dictionary with function names as keys and function objects as values
        """
        kernel_functions = {}

        # Get all class methods
        for name, method in inspect.getmembers(cls, predicate=inspect.isfunction):
            # Skip this method itself and any private/special methods
            if name.startswith("_") or name == "get_all_kernel_functions":
                continue

            # Check if the method has the kernel_function annotation
            # by looking at its __annotations__ attribute
            method_attrs = getattr(method, "__annotations__", {})
            if hasattr(method, "__kernel_function__") or "kernel_function" in str(
                method_attrs
            ):
                kernel_functions[name] = method

        return kernel_functions

    @classmethod
    def generate_tools_json_doc(cls) -> str:
        """
        Generate a JSON document containing information about all methods in the class.

        Returns:
            str: JSON string containing the methods' information
        """

        tools_list = []

        # Get all methods from the class that have the kernel_function annotation
        for name, method in inspect.getmembers(cls, predicate=inspect.isfunction):
            # Skip this method itself and any private methods
            if name.startswith("_") or name == "generate_tools_json_doc":
                continue

            # Check if the method has the kernel_function annotation
            if hasattr(method, "__kernel_function__"):
                # Get method description from docstring or kernel_function description
                description = ""
                if hasattr(method, "__doc__") and method.__doc__:
                    description = method.__doc__.strip()

                # Get kernel_function description if available
                if hasattr(method, "__kernel_function__") and getattr(
                    method.__kernel_function__, "description", None
                ):
                    description = method.__kernel_function__.description

                # Get argument information by introspection
                sig = inspect.signature(method)
                args_dict = {}

                # Get type hints if available
                type_hints = get_type_hints(method)

                # Process parameters
                for param_name, param in sig.parameters.items():
                    # Skip first parameter 'cls' for class methods (though we're using staticmethod now)
                    if param_name in ["cls", "self"]:
                        continue

                    # Get parameter type
                    param_type = "string"  # Default type
                    if param_name in type_hints:
                        type_obj = type_hints[param_name]
                        # Convert type to string representation
                        if hasattr(type_obj, "__name__"):
                            param_type = type_obj.__name__.lower()
                        else:
                            # Handle complex types like List, Dict, etc.
                            param_type = str(type_obj).lower()
                            if "int" in param_type:
                                param_type = "int"
                            elif "float" in param_type:
                                param_type = "float"
                            elif "bool" in param_type:
                                param_type = "boolean"
                            else:
                                param_type = "string"

                    # Create parameter description
                    # param_desc = param_name.replace("_", " ")
                    args_dict[param_name] = {
                        "description": param_name,
                        "title": param_name.replace("_", " ").title(),
                        "type": param_type,
                    }

                # Add the tool information to the list
                tool_entry = {
                    "agent": cls.agent_name,  # Use HR agent type
                    "function": name,
                    "description": description,
                    "arguments": json.dumps(args_dict).replace('"', "'"),
                }

                tools_list.append(tool_entry)

        # Return the JSON string representation
        return json.dumps(tools_list, ensure_ascii=False, indent=2)
