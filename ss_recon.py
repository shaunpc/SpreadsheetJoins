import pandas as pd

# Read in the data sets
xls = 'ss_recon_data.xlsx'
df_left = pd.read_excel(xls, 'LEFT');
df_right_main = pd.read_excel(xls, 'RIGHT-MAIN');
df_right_sub1 = pd.read_excel(xls, 'RIGHT-SUB1');

# convert to dictionaries
# gives in form: {'KEY': {'col1': 'BOB', 'col2': 'EUR', 'col3': 1200, 'col4': 'DE'}}
# dict_left = df_left.set_index('ID ').T.to_dict()
# dict_right_main = df_right_main.set_index('ID ').T.to_dict()
# dict_right_sub1 = df_right_sub1.set_index('ID ').T.to_dict()

# but if we need to handle multiple rows with the same key, then need to avoid using the 'ID' as index on sub-tables
# {0: {'ID ': 'KEY', 'col1': 'CREDIT', 'col2': 'BOB', 'col3': 'DE', 'col4': 1.5},
#  1: {'ID ': 'KEY', 'col1': 'DEBIT', 'col2': 'STEVE', 'col3': 'DE', 'col4': 22.8}}
dict_left = df_left.T.to_dict()


# function to return key for any value
def get_key(my_dict, val):
    for key, value in my_dict.items():
        if val == value:
            return key
    return "key doesn't exist"


for l_row_id, l_row_values in dict_left.items():
    pkey = l_row_values['ID']
    print('Processing row: ', l_row_id, ' / ', l_row_values)

    # Restrict the RIGHT_MAIN dictionary to only those rows with matching PK
    df_right_main_pkey = df_right_main.loc[df_right_main['ID'] == pkey]
    dict_right_main = df_right_main_pkey.T.to_dict()
    print(dict_right_main)
    # Restrict the RIGHT_SUB1 dictionary to only those rows with matching PK
    df_right_sub1_pkey = df_right_sub1.loc[df_right_main['ID'] == pkey]
    dict_right_sub1 = df_right_sub1_pkey.T.to_dict()
    print(dict_right_sub1)

    # Find the ID in the RIGHT_MAIN side
    l_row_found = False
    for r_row_id, r_row_values in dict_right_main.items():
        if pkey == r_row_values.get('ID'):
            # Found the same ID, let's see which columns can be found
            print('Found: ', r_row_id, ' / ', r_row_values)
            l_row_found = True
            r_
            for l_col_id, l_col_value in l_row_values.items():
                r_found_key = get_key(r_row_values, l_col_value)
                print('    Mapped LEFT ', l_col_id, '/', l_col_value, 'to RIGHT_MAIN ', r_found_key)

