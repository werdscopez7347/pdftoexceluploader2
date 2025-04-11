import sys
import openpyxl
from openai import OpenAI

# Get filename from command line args
if len(sys.argv) < 2:
    print("Usage: python convert.py <excel_file>")
    sys.exit(1)

excel_file = sys.argv[1]

# Load the Excel file
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Extract values from column D starting from D2
text_blocks = []
for row in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
    cell_value = row[0].value
    if cell_value:
        text_blocks.append(f'text: "{cell_value}"')

# Build prompt
prompt_content = '''Split the following text into Name, Location (including street), and Phone Number. Follow these rules strictly:

1. **Capitalize names properly** (e.g., "AHMADMASOUD" → "Ahmad Masoud").
2. **Format the location correctly** for Google Maps:
   - Include proper punctuation (commas) between different parts of the address (e.g., "Al Awir 1, Dubai, UAE").
   - If the location has multiple parts like street name, block, or apartment number, ensure they are separated properly (e.g., "Nasayem Avenue, Block D, Apt 501, Dubai").
3. If a **Makani number** (a 10-digit number) is present, include it in the **Location** as:  
   **Makani Number: [10-digit-number]**.  
   ⚠️ If the Makani number is split across two parts (e.g., 566658776 and 7), **combine them** to form 5666587767.
4. **Normalize phone numbers** to the format: +971 50 XXX XXXX. There may be multiple formats of the number, such as:  
   - 971506679882  
   - 971 506679882  
   - 506679882  
   Ensure phone numbers appear in the **[Number]** section, not in the **[Location]** section.
5. Remove all spaces first so u can analyze the text properly then add spaces accordingly.
6. If there are phone numbers which are incomplete , enclose the number with parentheses to indicate it is incomplete.when i mean incomplete i mean the number is less than 10 digits. or the number is not a valid number.

**Important Clarifications:**
- The **name** should be separated from the **location**. If the name contains multiple parts, ensure they are capitalized properly, and do not merge them into the location.
- If the text contains a **street address** like "Nasayem Avenue," "Block D," or "Apt 501," ensure these parts are correctly identified and formatted in the location.

**Process the following text:**\n\n'''

# Add the extracted text sections
prompt_content += '\n\n'.join(text_blocks)

# Initialize OpenAI client
client = OpenAI(api_key="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")

# Make request
completion = client.chat.completions.create(
    model="gpt-4o-mini",
    store=True,
    messages=[{
        "role": "user",
        "content": prompt_content
    }]
)

# Print result
print(completion.choices[0].message.content)
