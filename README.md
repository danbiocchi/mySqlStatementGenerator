# mySqlStatementGenerator
> Windows ONLY
>Purpose of script, convert a excel doc to an insert statement to be inputted into a mySQL database <<<

Converts .xslm (macro enabled excel docs) to insert statements given database column title and column which corresponds to said data, printing an insert statement in the format:

`INSERT INTO 'table_name'(column_1,column_2,...) VALUES (value_1,value_2,...);`

I have addded 1-9 columns with the script using 3 values so far as an example and the other entries commented out. 
To use those entries just un-comment and then change values depending on where the information exists.


TO USE THE SCRIPT >>

>1. Open Excel with relevant data (Make sure developer tab is enabled https://www.techonthenet.com/excel/questions/developer_tab2013.php)
>2. Navigate to the Developer Tab
>3. Click "Visual Basic" on the left-most of the ribbon.
>4. Select the relevant Sheet from the left hand side.
>5. Paste the code provided above
>6. Alter the program settings in VBA Editor 
>7. Run Sub-routine sing F5 or clicking "Run" on the top then selecting "Run Sub-Routine"
>8. Insert statement will be printed to a file on the desktop.
