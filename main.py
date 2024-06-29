# This is a sample Python script.
from excel_transformer import transform_file


if __name__ == '__main__':
    transform_file(
        'input/transactions.csv',
        '/Users/teotoplak/Drive/areas/finances/personal_finances.xlsx',
        '/Users/teotoplak/Drive/areas/finances/backup',
        'revolut'
    )

