import xlsxwriter
import random
from datetime import datetime, timedelta

def create_huge_messy_excel():
    workbook = xlsxwriter.Workbook('Messy_Bank_Export_Huge.xlsx')
    worksheet = workbook.add_worksheet('Raw Data Dump')

    # Messy headers
    worksheet.write(0, 0, '**** GHOST ACCOUNT EXPORT ***')
    worksheet.write(1, 0, 'DATE RANGE: 08/01/2023 - 10/31/2023')
    worksheet.write(3, 0, 'POST_DATE')
    worksheet.write(3, 1, 'TXN_REF_NUM')
    worksheet.write(3, 2, 'RAW_MEMO_DESC')
    worksheet.write(3, 3, 'DEBIT_AMT')

    start_date = datetime(2023, 8, 1)
    
    # Pools of messy merchant strings
    coffee_shops = ['SQ *LOCAL COFFEE', 'STARBUCKS STORE #1982', 'PHILZ COFFEE SF', 'PEETS COFFEE 119']
    dining = ['DOORDASH*UBER EATS RESTAURANT DLV', 'UBER EATS HELA J', 'CHIPOTLE 1928', 'MCDONALDS 1922', 'IN N OUT BURGER']
    groceries = ['WHOLEFDS SFO 10452', 'TRADER JOES #102 QPS', 'SAFEWAY GROCERY #991', 'SAFEWAY']
    transport = ['POS DEBIT -- UBER *TRIP 8921 SF CA', 'LYFT *RIDE 9912', 'BART RESIDUAL VALUE', 'UBER *PENDING']
    shopping = ['AMZN MKTP US*8G19 AMZN.COM/BILL WA', 'AMZN.COM/BILL', 'APPLE STORE R102 SAN FRANCISCO', 'BEST BUY MKT 102', 'TARGET T-192']
    utilities = ['CHK 8291 PG&E UTILITY WEB PMT', 'COMCAST INTERNET SVC', 'SF WATER DEPT BILL']

    # Subscriptions (fixed pairs of merchant + amount)
    subscriptions = [
        ('ACH Electronic Debit - NETFLIX.COM', 15.99),
        ('RECURRING PAYMENT SPOTIFY USA NY', 10.99),
        ('AMZN PRIME MEMBERSHIP AMZN.COM/BILL', 14.99),
        ('POS PURCH PLANET FITNESS CLUB #11', 29.99),
        ('APPLE.COM/BILL ICLOUD STORAGE', 2.99),
        ('ADOBE *CREATIVE CLOUD 800-833-6687 CA', 54.99)
    ]

    current_row = 4
    
    # Generate exactly 200 random transactions over 90 days
    # Plus the monthly subscriptions specifically injected.
    
    transactions = []
    
    # 1. Inject recurring subscriptions on specific days each month
    for month in [8, 9, 10]:
        for sub_desc, sub_amt in subscriptions:
            # Pick a somewhat consistent day for subscriptions (e.g. 1st through 5th)
            sub_date = datetime(2023, month, random.randint(1, 10))
            txn_ref = f"REF-{random.randint(1000, 9999)}"
            transactions.append((sub_date, txn_ref, sub_desc, sub_amt))
            
    # 2. Add random day-to-day transactions
    for _ in range(180):
        random_days = random.randint(0, 90)
        txn_date = start_date + timedelta(days=random_days)
        txn_ref = f"REF-{random.randint(10000, 99999)}"
        
        category = random.choices(['coffee', 'dining', 'groceries', 'transport', 'shopping', 'utilities'], 
                                  weights=[30, 20, 15, 20, 10, 5])[0]
        
        if category == 'coffee':
            desc = random.choice(coffee_shops) + f" {random.randint(100, 999)}"
            amt = round(random.uniform(3.50, 12.00), 2)
        elif category == 'dining':
            desc = random.choice(dining) + f" REQ-{random.randint(10, 99)}"
            amt = round(random.uniform(15.00, 65.00), 2)
        elif category == 'groceries':
            desc = random.choice(groceries)
            amt = round(random.uniform(45.00, 210.00), 2)
        elif category == 'transport':
            desc = random.choice(transport)
            amt = round(random.uniform(9.00, 45.00), 2)
        elif category == 'shopping':
            desc = random.choice(shopping)
            amt = round(random.uniform(20.00, 150.00), 2)
        else:
            desc = random.choice(utilities)
            amt = round(random.uniform(50.00, 120.00), 2)
            
        transactions.append((txn_date, txn_ref, desc, amt))

    # Sort transactions chronologically
    transactions.sort(key=lambda x: x[0])

    for txn in transactions:
        # Occasionally add an annoying blank row for extra messiness
        if random.random() < 0.05:
            current_row += 1
            
        worksheet.write(current_row, 0, txn[0].strftime('%Y-%m-%d'))
        worksheet.write(current_row, 1, txn[1])
        worksheet.write(current_row, 2, txn[2])
        worksheet.write(current_row, 3, str(txn[3]))
        current_row += 1

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 50)
    worksheet.set_column('D:D', 12)

    workbook.close()
    print("Huge messy Excel template created.")

if __name__ == '__main__':
    create_huge_messy_excel()
