import psycopg2
import time
import openpyxl

# Add PO  # QC LAB at the end
# make it so you cant add over a certain number
# when finished an email is sent automatically to MPL

def mpl():
    otc = openpyxl.load_workbook('otcmpl.xlsx')
    otc_sheet = otc.get_sheet_by_name('Sheet1')
    non_otc = openpyxl.load_workbook('nonotcmpl.xlsx')
    non_otc_sheet = non_otc.get_sheet_by_name('Sheet1')
    cursor = conn.cursor()
    # Creates a list of batch numbers to be sent for micro
    list_of_batch_numbers = []
    initials = input('Enter your initials: ')
    print('Enter the batch number to be sent to micro\nif you are done type QUIT: ')
    while True:
        batch_id = input('Next batch id:')
        list_of_batch_numbers.append(batch_id.upper())
        if batch_id.upper() == 'QUIT':
            break
        else:
            continue

    # Execute the query --> assign it to a variable using fetchone
    # Fetch one returns a tuple so we select the item at index 0
    otc_query_list = []
    non_otc_list = []
    for batch in list_of_batch_numbers:
        try:
            cursor.execute("""SELECT products.otc, company.company_code, products.formula_number, products.product_name
                        FROM batch
                        JOIN products
                             ON batch.formula_number = products.formula_number
                        JOIN company
                             ON products.company_name = company.company_name
                        JOIN planners
                             ON company.planner_name = planners.planner_name
                            WHERE batch_number = (%s)""", (batch,))
            query_results = list(cursor.fetchone())
            if len(query_results) > 1 and query_results[0] is True:
                query_results.append(batch)
                otc_query_list.append(query_results)
            elif len(query_results) > 1 and query_results[0] is False:
                query_results.append(batch)
                non_otc_list.append(query_results)
        except:
            pass

    print(non_otc_list)
    print(otc_query_list)

    today = time.strftime("%m-%d-%Y")
    otc_row = 9
    non_otc_row = 8
    sample_number = 1

    for item in otc_query_list:
        print(item)
        otc_sheet['A' + str(otc_row)].value = sample_number  # Sample number
        otc_sheet['B' + str(otc_row)].value = item[2] + ': ' + item[1] + ' ' + item[3] # Sample Description
        otc_sheet['M' + str(otc_row)].value = 'B'  # Type of Sample
        otc_sheet['p' + str(otc_row)].value = item[4]  # Batch number
        otc_sheet['R' + str(otc_row)].value = item[2]  # Production Date
        otc_sheet['T' + str(otc_row)].value = initials.upper()  # Prepared by
        otc_sheet['U' + str(otc_row)].value = today  # Date Prepared
        otc_row += 1
        sample_number +=1

    otc_sheet['B' + str(otc_row)].value = 'PO # QC LAB'

    for item in non_otc_list:
        print(item)
        non_otc_sheet['A' + str(non_otc_row)].value = sample_number  # Sample number
        non_otc_sheet['B' + str(non_otc_row)].value = item[2] + ': ' + item[1] + ' ' + item[3] # Sample Description
        non_otc_sheet['M' + str(non_otc_row)].value = 'B'  # Type of Sample
        non_otc_sheet['p' + str(non_otc_row)].value = item[4]  # Batch number
        non_otc_sheet['R' + str(non_otc_row)].value = item[2]  # Production Date
        non_otc_sheet['T' + str(non_otc_row)].value = initials.upper()  # Prepared by
        non_otc_sheet['U' + str(non_otc_row)].value = today  # Date Prepared
        non_otc_row += 1
        sample_number +=1

    non_otc_sheet['B' + str(non_otc_row)].value = 'PO # QC LAB'


    otc.save('MPL OTC ' + today + '.xlsx')
    non_otc.save('MPL NON OTC' + today + '.xlsx')


try:
    conn = psycopg2.connect("dbname='work' host='localhost' password=''")
    print('You successfully connected to the Bentley Labs QC Database.')
    print('-----------------------------------------------------------')
except Exception as error:
    print(error)

mpl()



