"""
This module explores spreadsheet automation 



"""
import pandas as pd
import seaborn as sns
import logging
import os
from utils import get_credentials, create_google_spreadsheet, write_to_sheet

GCP_SERVICE_ACCOUNT_EMAIL = 'example-service-account@gorgon-city-90.iam.gserviceaccount.com'
CREATE_NEW_SPREADSHEET = False


def order_columns(df):
    columns_first = [
        df.columns[idx] for idx, dtype in enumerate(df.dtypes) if str(dtype) == 'object'
    ]
    columns_last = [col for col in df.columns if col not in columns_first]
    df = df[columns_first + columns_last]
    return df


def download_data():
    df = sns.load_dataset('iris')
    logger.info(df.shape)
    return df


if __name__ == '__main__':
    logging.basicConfig()
    logger = logging.getLogger('logger')
    logger.setLevel(logging.INFO)
    df = download_data()
    df = order_columns(df)

    # create or fetch spreadsheet
    if CREATE_NEW_SPREADSHEET:
        sheet_id, gc, link = create_google_spreadsheet(
            title='DFP AutoUpdating Sheet',
            service_account_email=GCP_SERVICE_ACCOUNT_EMAIL,
            credential_password_env_var_name='SA_PASSWORD',
            email_message="DFP is sharing its cool autoupdating sheet with you, enjoy!",
            send_notification=True,
            share_domains=['amagalhaes@civisanalytics.com'],
        )

    else:
        sheet_id = '1vMGxaHnomqF0HWSAbONFTVZEgwcUInKpzEWpsZGK3iY'
        link = f'\nhttps://docs.google.com/spreadsheets/d/{sheet_id}'
        _, gc = get_credentials(credential_password_env_var_name='SA_PASSWORD')

    # write df to sheet
    write_to_sheet(
        gc,
        sheet_id,
        'floral_data',
        df,
        percent_cols=[],
        column_width_dict={0: 250},
        no_borders=True,
        wrap_cols=[],
    )
    logger.info("Spreadsheet is ready!!!")
    logger.info('\n' + link)
    cols_to_colors = (
        {
            'sepal_length': 'green',
            'sepal_width': 'grey',
            'petal_length': 'blue',
            'petal_width': 'red',
        },
    )
