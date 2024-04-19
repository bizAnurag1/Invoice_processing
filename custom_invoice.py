from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from azure.storage.blob import BlobServiceClient
import pandas as pd
 
# Azure Form Recognizer credentials
form_recognizer_endpoint = "https://doc-intelligence-inv-pro.cognitiveservices.azure.com/"
form_recognizer_key = "6510688e806d4e44b1a0067f321574b0"
model_id = "Custom_Domestic_Invoice"
 
# Azure Blob Storage credentials
storage_connection_string = "DefaultEndpointsProtocol=https;AccountName=adlstilnvoiceinput;AccountKey=iCycvf2YkC96SU9teDiTswyf6RXxJGqE8fdznzsDw1iLiN7TzVoEzdvkSW4LFMGNr53AfzRBZxOl+AStQsaZKw==;EndpointSuffix=core.windows.net"
container_name = "test"
 
# Create clients
form_recognizer_client = DocumentAnalysisClient(
    endpoint=form_recognizer_endpoint,
    credential=AzureKeyCredential(form_recognizer_key)
)
blob_service_client = BlobServiceClient.from_connection_string(storage_connection_string)
container_client = blob_service_client.get_container_client(container_name)
 
# Function to analyze a document
def analyze_document(blob_url):
    poller = form_recognizer_client.begin_analyze_document_from_url("Custom_Domestic_Invoice", blob_url)
    result = poller.result()
    return result
 
# Collect results in a list
results = []
for blob in container_client.list_blobs():
    blob_url = f"https://{blob_service_client.account_name}.blob.core.windows.net/{container_name}/{blob.name}"
    print(f"Analyzing invoice from blob: {blob.name}")
    analysis_result = analyze_document(blob_url)
    for doc in analysis_result.documents:
        doc_fields = {field_name: field.value for field_name, field in doc.fields.items()}
        results.append(doc_fields)
 
# Convert results to a pandas DataFrame
df = pd.DataFrame(results)

# Post-Processing the DataFrame
for column in df.columns:
    df[column] = df[column].map(lambda x: x.strip(':') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip('|') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip('!') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip('*') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip('#') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip('(') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.lstrip(')') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x[1:] if isinstance(x, str) and x.startswith(".") else x)
    df[column] = df[column].map(lambda x: x.strip('"') if isinstance(x, str) else x)
    df[column] = df[column].map(lambda x: x.strip(',') if isinstance(x, str) else x)
 
# Specify your desired Excel file name
excel_file_name = "sample_Invoice_Output.xlsx"
 
# Save the DataFrame to an Excel file
df.to_excel(excel_file_name, index=False)
 
print(f"DataFrame has been saved to {excel_file_name}")