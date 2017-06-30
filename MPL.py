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
    query_list = []
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
            query_results = cursor.fetchone()
            if len(query_results) > 1:
                query_list.append(query_results)
        except:
            pass

    today = time.strftime("%m-%d-%Y")
    otc_row = 9
    non_otc_row = 8
    for item in query_list:
        if item[0] is True:
            otc_sheet['A' + str(otc_row)].value = item[2]
            otc_sheet['B' + str(otc_row)].value = item[2]
            otc_sheet['M' + str(otc_row)].value = item[2]
            otc_sheet['p' + str(otc_row)].value = item[2]
            otc_sheet['R' + str(otc_row)].value = item[2]
            otc_sheet['T' + str(otc_row)].value = item[2]
            otc_sheet['U' + str(otc_row)].value = item[2]
            otc_row += 1
        else:
            non_otc_sheet['A' + str(non_otc_row)].value = item[2]
            non_otc_row += 1

    otc.save('MPL OTC ' + today + '.xlsx')
    non_otc.save('nonotctest.xlsx')


try:
    conn = psycopg2.connect("dbname='work' host='localhost' password=''")
    print('You successfully connected to the Bentley Labs QC Database.')
    print('-----------------------------------------------------------')
except Exception as error:
    print(error)

mpl()


