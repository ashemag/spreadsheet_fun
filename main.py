import pandas as pd
import logging
import os
from utils import get_credentials, create_google_spreadsheet, write_to_sheet
from column_color_codings import COLS_TO_COLORS

GCP_SERVICE_ACCOUNT_EMAIL = 'example-service-account@gorgon-city-90.iam.gserviceaccount.com'
CREATE_NEW_SPREADSHEET = False


def download_data():
    df = pd.read_csv('data/blast_500.csv')
    columns = [
        'District',
        'DMA',
        'Channel',
        'Week',
        'Dem GRPs',
        'Rep GRPs',
        'Net GRPs',
        'Value Per Dollar',
        'Change In DMA Voteshare',
        'Cost Per Impression',
    ]
    df = df.sort_values(by=['Value Per Dollar'], ascending=False)

    logger.info(df.shape)
    return df[columns]


if __name__ == '__main__':
    logging.basicConfig()
    logger = logging.getLogger('logger')
    logger.setLevel(logging.INFO)
    df = download_data()

    # create or fetch spreadsheet
    if CREATE_NEW_SPREADSHEET:
        sheet_id, gc, link = create_google_spreadsheet(
            title='DFP [AutoUpdating] Sheet',
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
        'grp blast',
        df,
        percent_cols=[],
        column_width_dict={0: 250},
        no_borders=True,
        wrap_cols=[],
        cols_to_colors=COLS_TO_COLORS,
    )
    logger.info("Spreadsheet is ready!!!")
    logger.info('\n' + link)
