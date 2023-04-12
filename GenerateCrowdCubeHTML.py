import pandas as pd
import datetime
import shutil

# DEPENDENCIES - reading excel needs:
# pip install xlrd
# pip install openpyxl
# pip install --upgrade pandas


xl_file = r"C:\Users\shaun\OneDrive\Documents\Personal\Lucy-Flat-Income (NEW).xlsx"
xl_sheet = r"CrowdCube"

html_file = r"C:\Users\shaun\OneDrive\Documents\My Web Sites\WebSite1\Investments-AUTO.html"

print("Copying master CSS Style Sheet to webpage directory")
dest_css_file = r"C:\Users\shaun\OneDrive\Documents\My Web Sites\WebSite1\Investments-AUTO.css"
src_css_file = r"Investments-AUTO.css"
shutil.copyfile(src_css_file, dest_css_file)

print("Loading data spreadsheet:", xl_file )
df = pd.read_excel(xl_file, sheet_name=xl_sheet, header=0)

time_now = datetime.datetime.now()
print('\nStarting: Angel Investment webpage generation @ ' + time_now.strftime("%H:%M:%S") + '\n')
# Need to set encoding for the £ sign inclusion
out_file = open(html_file, "w", encoding='utf-8')
out_file.write('<!DOCTYPE html>\n')
out_file.write('<html>\n')
out_file.write('<head>\n')
out_file.write('   <meta charset="utf-8" />\n')
out_file.write('   <title>Angel Investments</title>\n')
out_file.write('   <link rel="stylesheet" href="Investments-AUTO.css">\n')
out_file.write('</head>\n')
out_file.write('<body>\n')
out_file.write('\n')
out_file.write('<!-- AUTOGENERATED CODE  -->\n')
out_file.write('\n')
out_file.write('<div class="page-grid">\n')
out_file.write('    <div><h1>Cotter Family Angel Investments</h1> (generated: ' + time_now.strftime("%d-%B-%Y") + ')</div>\n')
out_file.write('\n')
out_file.write('    <div class="outer-grid" >\n')
out_file.write('\n')
out_file.write('    <!-- REPEATED CODE FOR EACH INVESTMENT  -->\n')

# DO STUFF HERE
total_value = 0
total_outcome = 0
total_count = 0
df = df.reset_index()  # make sure indexes pair with number of rows
for index, row in df.iterrows():
    r_status = row['Status']
    if pd.isnull(r_status):
        break
    total_count += 1
    r_platform = row['Platform']
    r_platform_website = row['Platform Website']
    r_platform_image = row['Platform Image']
    r_company = row['Company']
    r_company_website = row['Company Website']
    r_company_image = row['Company Image']
    r_holder = row['Who']
    r_invest_date_raw = row['Invest Date']
    r_invest_outcome = row['Outcome']
    r_invest_outcome_date = row['Outcome Date']
    r_invest_outcome_value = row['Outcome Value']
    if pd.isnull(r_invest_date_raw):
        r_invest_date = "Status: " + r_status
    else:
        r_invest_date = r_status + ": " + row['Invest Date'].to_pydatetime().strftime("%B %Y")
    r_invest_value = "Value: £" + "{0:,.2f}".format(row['Value'])
    total_value += row['Value']

    if pd.isna(r_invest_outcome):
        print('  > Processing: ' + r_company)
    else:
        print('  > Processing: ' + r_company + '  \t' + str(r_invest_outcome))
        total_outcome += row['Outcome Value']

    out_file.write('\n')
    # figure out status to determine background cell style
    match r_invest_outcome:
        case 'DEFUNCT':
            out_file.write('        <div class="inner-grid-defunct"> <!-- ' + r_company + ' -->\n')
        case 'CANCELLED':
            out_file.write('        <div class="inner-grid-cancelled"> <!-- ' + r_company + ' -->\n')
        case 'EXITED':
            out_file.write('        <div class="inner-grid-exited"> <!-- ' + r_company + ' -->\n')
        case _:
            out_file.write('        <div class="inner-grid"> <!-- ' + r_company + ' -->\n')
    out_file.write('            <div class="inner-grid-platform">\n')
    if not(pd.isna(r_platform_website)):
        out_file.write('                <a href="' + r_platform_website + '" target="_blank" rel="noopener noreferrer">\n')
    out_file.write('                    <img src="'+ r_platform_image + '" title="' + r_platform + '">\n')
    if not(pd.isna(r_platform_website)):
        out_file.write('                </a>\n')
    out_file.write('            </div>\n')
    out_file.write('            <div class="inner-grid-holder">' + r_holder + '</div>\n')
    out_file.write('            <div class="inner-grid-investment">\n')
    out_file.write('                <a href="' + r_company_website + '" target="_blank" rel="noopener noreferrer">\n')
    out_file.write('                    <img src="' + r_company_image + '" title="' + r_company + '">\n')
    out_file.write('                </a>\n')
    out_file.write('            </div>\n')
    out_file.write('            <div class="inner-grid-since">' + r_invest_date + '</div>\n')
    out_file.write('            <div class="inner-grid-amount">' + r_invest_value + '</div>\n')
    out_file.write('        </div>  <!-- end of inner-grid items -->\n')

out_file.write('\n')
out_file.write('    </div>   <!-- end of outer-grid items -->\n')
out_file.write('\n')
summary_value_line = "{0:,.2f}".format(total_value) + ' Invested / £' + "{0:,.2f}".format(total_outcome) + ' Returned'
out_file.write('    <div><h1>' + str(total_count) + ' Investments : £' + summary_value_line+ '</h1></div>\n')
out_file.write('\n')
out_file.write('</div>   <!-- end of page-grid items -->\n')
out_file.write('</body>\n')
out_file.write('</html>\n')
out_file.close()

print('\nCompleted: ' + str(total_count) + ' Investments - Total Value: £' + "{0:,.2f}".format(total_value) + '\n')
