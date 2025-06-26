import pandas as pd

# Créer des données d'exemple
data = {
    'Description': [
        'Payment to ADVICEPRO for consulting services AE1602600010153',
        'Office supplies OFFICE 123 PARIS',
        'Regular payment for utilities',
        'ADVICEPRO monthly fee AE1602600010154',
        'Bank transfer for equipment'
    ],
    'Bank account': [
        'EUR-ACCOUNT-001',
        'USD-ACCOUNT-002',
        'EUR-ACCOUNT-003',
        'USD-ACCOUNT-004',
        'EUR-ACCOUNT-005'
    ],
    'Amount': [1500.00, 250.00, 120.00, 1800.00, 3200.00],
    'Date': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19'],
    'Nature': ['', '', '', '', ''],
    'Descrip': ['', '', '', '', ''],
    'Vessel': ['', '', '', '', ''],
    'Service': ['', '', '', '', ''],
    'Reference': ['', '', '', '', '']
}

df = pd.DataFrame(data)
df.to_excel('test_data.xlsx', index=False)
print("Fichier d'exemple créé : test_data.xlsx")

