#In[]: 
import xml.etree.ElementTree as ET
import pandas as pd

# Define the namespace
namespace = {'ns': 'urn:schemas-microsoft-com:office:spreadsheet'}

# Parse the XML file
tree = ET.parse(r"C:\Users\aru\Downloads\13122779.xml")  # Replace 'your_file.xml' with the actual file path
root = tree.getroot()

# Extract column names from the first row
columns = [cell.text.strip() for cell in root.findall('.//ns:Row[1]/ns:Cell/ns:Data', namespace)]

# Extract data from each row
data_rows = root.findall('.//ns:Row', namespace)[1:]
data = [[cell.text.strip() if cell.find('./ns:Data', namespace) is not None else None for cell in row.findall('./ns:Cell', namespace)] for row in data_rows]

# Create a Pandas DataFrame
df = pd.DataFrame(data, columns=columns)

# Display the DataFrame
print(df)



#In[]: