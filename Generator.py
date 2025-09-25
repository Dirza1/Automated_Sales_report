import pandas as pd

# Voorbeeld data
data = [
    [1001, '2025-09-01', 'Jan de Vries', 'Pen', 10, 1.5, 15],
    [1002, '2025-09-01', 'Maria Smit', 'Notebook', 5, 3.0, 15],
    [None, '2025-09-02', 'Tom Janssen', 'Pen', 8, 1.5, 12],
    [1004, '2025-09-02', None, 'Marker', 12, 2.0, 24],
    [1005, '2025-09-03', 'Anne Bakker', 'Pen', None, 1.5, None],
    [1006, '2025-09-03', 'Lisa de Jong', 'Notebook', 7, 3.0, 21],
    [1007, None, 'Pieter Klaas', 'Pen', 6, None, None],
    [1008, '2025-09-04', 'Jan de Vries', 'Marker', 10, 2.0, 20],
    [1009, '2025-09-04', 'Maria Smit', 'Pen', 15, 1.5, 22.5],
    [1010, '2025-09-05', 'Tom Janssen', 'Notebook', 4, 3.0, 12]
]

# Kolomnamen
columns = ['OrderID', 'Date', 'Curtomer', 'Product', 'Count', 'Price', 'Total']

# DataFrame maken
df = pd.DataFrame(data, columns=columns)

# Opslaan als Excel-bestand
df.to_excel('example_sales.xlsx', index=False)

print("Excel-file 'example_sales.xlsx' is made!")
