import pandas as pd
from normalization_utils import process_excel

import pandas as pd
from normalization_utils import process_excel

def preprocess_input_file(input_path: str) -> pd.DataFrame:
    """
    Reads the input Excel file and performs preprocessing:
      - Processes the 'Value' column using the normalization utility.
      - Renames "UpdatedValue" to "Value_Normalized".
      - Preserves the original "Value" in "Value_E".
      - Overwrites "Value" with the normalized data.
    
    Returns:
      A DataFrame with a smooth "Value" column (normalized) and the original values in "Value_E".
    """
    # Read the original Excel file to confirm it has a 'Value' column
    df = pd.read_excel(input_path)
    if "Value" not in df.columns:
        raise ValueError("Input file must contain a 'Value' column.")
    
    # Process the file with the normalization utilities (this returns a DataFrame with extra columns)
    df_processed = process_excel(input_path)
    
    # Rename UpdatedValue to Value_Normalized for clarity
    df_processed.rename(columns={"UpdatedValue": "Value_Normalized"}, inplace=True)
    
    # Preserve the original value in Value_E
    df_processed["Value_E"] = df_processed["Value"]
    
    # Overwrite the 'Value' column with the normalized data
    df_processed["Value"] = df_processed["Value_Normalized"]
    
    return df_processed
