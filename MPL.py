import psycopg2
import time


def mpl():
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
    otc_batches = []
    non_otc_batches = []

    # Execute the query --> assign it to a variable using fetchone
    # Fetch one returns a tuple so we select the item at index 0
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
            otc_bool = cursor.fetchone()
            final_string = otc_bool[2] + ': ' + otc_bool[1] + ' ' + otc_bool[3] + ' B ' + batch + ' DOM ' + initials.upper()+' '+time.strftime("%m-%d-%Y")
            if otc_bool[0] is True:
                otc_batches.append(final_string)
            else:
                non_otc_batches.append(final_string)
        except:
            print('batch # not in database')
    return otc_batches, non_otc_batches


try:
    conn = psycopg2.connect("dbname='work' host='localhost' password=''")
    print('You successfully connected to the Bentley Labs QC Database.')
    print('-----------------------------------------------------------')
except Exception as rofl:
    print(rofl)

print(mpl())
