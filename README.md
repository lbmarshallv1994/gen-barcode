This Ruby program generates an excel spreadsheet of barcodes with their checkdigit.

## Dependencies
Genbarcode uses the spreadsheet gem to write out data to XLS files.

```
sudo gem install spreadsheet
```

## Usage
```
ruby GenBarcode.rb [output file] [barcode begin range] [barcode end range]
```

if end range is smaller than begin range, the end range will become begin range + end range

## Checkdigit Algorithm
Starting from the left most digit, multiply each digit by 2 or 1 alternating each digit. If the result of any multiplication is a two digit number, add the two digits together. Add together all of the resulting 1 digit numbers. The checkdigit is the number to be added to the result that will raise it to the nearest power of 10. 

