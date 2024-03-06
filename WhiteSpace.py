import pandas as pd

def clean_and_save_csv(input_file, output_file):
    df = pd.read_csv(input_file, encoding='latin-1')
    df.columns = df.columns.str.strip()
    df = df.apply(lambda x: x.map(lambda y: y.strip() if isinstance(y, str) else y))
    df.to_csv(output_file, index=False)

if __name__ == "__main__":
    input_file = "Analysis of Slum Publications.csv"
    output_file = 'SlumPublications.csv'

    clean_and_save_csv(input_file, output_file)
