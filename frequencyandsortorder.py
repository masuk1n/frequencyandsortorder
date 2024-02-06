import pandas as pd

def calculate_item_frequencies(input_file, output_file):
    # Read the input Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Calculate item frequencies
    item_frequencies = df.iloc[:, 0].value_counts().reset_index()
    item_frequencies.columns = ['Item', 'Frequency']

    # Save the result in a new Excel file
    item_frequencies.to_excel(output_file, index=False)

if __name__ == "__main__":
    input_excel_file = r"D:\TEST\input2.xlsx"
    output_excel_file = r"D:\TEST\result2.xlsx"

    calculate_item_frequencies(input_excel_file, output_excel_file)
