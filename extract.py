import inventory
import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser('extract')
    parser.add_argument('spreadsheet')
    parser.add_argument('cachedir')
    args = parser.parse_args()
    
    sheet = inventory.ServproSheet(args.spreadsheet)
    sheet.parse()
    for row in sheet.rows[0:2]:
        sheet.populate_images(row, args.cachedir)
#    sheet.print()
