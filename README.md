# bloomberg-google-
python script to collect data about a company from bloomberg, google map and google. this project use selenium and lxml to parse data from combination of websites.
it takes input of excel and csv file and output in the same format.
it contains 3 files:
1. config.ini---> this is the file from which the script take inputs
       inside the config.ini , 
       path--> it refers to the path of the input file
       row--> it refers to the number of row from which script start processing 
       separator--> it refers to the separtor used in input file (if input file is a CSV file)
       columns --> it refers to the name of columns from which you want to make query to search on the google
2.log.txt ----> it notes down any error in the script 
3.main.py-----> it refers to the main script which take care of  all the scraping part.
