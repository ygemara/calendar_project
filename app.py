# Install Flask using pip: pip install flask
from flask import Flask, render_template, request, send_file, redirect
import pandas as pd
import numpy as np
import re

app = Flask(__name__)

# Define the route for the home page
@app.route('/')
def index():
    return render_template('index.html')

# Define the route for handling the file upload and processing
@app.route('/process', methods=['POST'])
def process():
    # Get the uploaded files from the request
    file1 = request.files['file1']
    #file2 = request.files['file2']

    # Save the uploaded files to the server
    file1_path = 'uploads/' + file1.filename
    #file2_path = 'uploads/' + file2.filename
    file1.save(file1_path)
    #file2.save(file2_path)

    # Process the Excel files using a Python script
    output_path1, output_path2, failed_rows_df = process_files(file1_path)

    # Return the download link for the output file and render the template
    return render_template('result.html', output_path1=output_path1,output_path2=output_path2, failed_rows=failed_rows_df.to_dict(orient='records'))

# Process the Excel files and generate the output file
def process_files(file1_path):
    # Read the Excel files using pandas
    df_old = pd.read_excel(file1_path, 1)
    df_new = pd.read_excel('uploads/Calendar Dates - 5784.xlsx', 1)
    df_old.columns = [col.strip() for col in df_old.columns]
    # Perform your desired data transformations here
    # ...
    pattern = r"(Line \d+)"
    df_old['Line Value'] = df_old['Date and Line'].str.extract(pattern)
    df_new['Line Value'] = df_new['Date and Line'].str.extract(pattern)
    df_new['Hebrew'] = df_new['Hebrew'].replace(r'^Adar.*', 'Adar', regex=True)
    df_new['row_value'] = list(range(len(df_new)))
    merged_df = df_old.merge(df_new, on=['Hebrew', 'Date', 'Line Value'], suffixes=['_5783', '_5784'], how='right')

    final_df = merged_df[['row_value', 'Hebrew', 'Date', 'Date and Line_5784', 'Text_5783']]
    final_df = final_df.sort_values(by=['Date and Line_5784', 'Text_5783'])
    final_df = final_df.drop_duplicates(subset=['Hebrew', 'Date', 'Date and Line_5784'], keep='first')
    final_df['char_limit'] = final_df['Date and Line_5784'].str.strip().str[-3:-1].astype(float)
    final_df['text_length'] = final_df.Text_5783.str.len().fillna(0)
    final_df = final_df.sort_values(by='row_value')




    risk_rows_table = final_df[final_df.char_limit < final_df.text_length]
    risk_rows_table['Reason'] = 'Character Limit Exceeded'
    risk_rows_table = risk_rows_table[['Hebrew', 'Date', 'Date and Line_5784', 'Text_5783', 'Reason']]

    set_old = set(df_old[~df_old.Text.isnull()]['Text'])
    set_new = set(final_df[~final_df.Text_5783.isnull()]["Text_5783"])
    missing_set = (set_old - set_new)

    missing_table = df_old[df_old.Text.isin(list(missing_set))]
    missing_table['Reason'] = 'Date Compatibility Issue'
    missing_table = missing_table[['Hebrew', 'Date', 'Date and Line', 'Text', 'Reason']]

    failed_rows_df = pd.DataFrame(np.vstack([missing_table, risk_rows_table]), columns=missing_table.columns)
    failed_rows = failed_rows_df.values.tolist()

    # Write the modified data to a new Excel file
    output_path1 = 'uploads/output.xlsx'
    output_path2 = 'uploads/errors.xlsx'

    #failed_rows_df.to_html(index=False)
    print(failed_rows)
    final_df.drop(['row_value','text_length'], axis=1, inplace=True)
    final_df.columns = ['Hebrew', 'Date', 'Date and Line', 'Text', 'char_limit']
    final_df.to_excel(output_path1, index=False)
    failed_rows_df.to_excel(output_path2, index=False)
    return output_path1, output_path2, failed_rows_df


# Define the route for file download
@app.route('/download/<file_name>')
def download(file_name):
    output_path = 'uploads/' +file_name
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run()

