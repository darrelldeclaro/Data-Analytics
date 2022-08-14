--ERROR UNABLE TO CONVERT STRING(nvarchar, varchar) TO DATE OR TIME HERE IS THE SOLUTION

--1) extract strings from delimeter and creat a new table
SELECT
    SUBSTRING( date, 0, CHARINDEX( '/', date ) ) AS Column1,
    SUBSTRING( date, ( CHARINDEX( '/', date) + 1 )  -- starting position of column 2.
        , CHARINDEX( '/', date,  ( CHARINDEX( '/', date ) + 1 ) ) - ( CHARINDEX( '/', date) + 1 ) -- length of column two is the number of characters between the two delimiters.
    ) AS Column2,
    SUBSTRING(
        date
        , CHARINDEX( '/', date, ( CHARINDEX( '/', date ) + 1 ) ) + 1
        , LEN( date )
    ) AS Column3, 
	* INTO CovidVaccinationsEdited2
FROM CovidVaccinationsEdited 


--2) delete date column
ALTER TABLE CovidVaccinationsEdited2
DROP COLUMN date_


--3) from the new column(month_name), you will identify month 'mmm'. 
SELECT *,
CASE
	WHEN _Month=1 THEN 'Jan'
	WHEN _Month=2 THEN 'Feb'
	WHEN _Month=3 THEN 'Mar'
	WHEN _Month=4 THEN 'Apr'
	WHEN _Month=5 THEN 'May'
	WHEN _Month=6 THEN 'Jun'
	WHEN _Month=7 THEN 'Jul'
	WHEN _Month=8 THEN 'Aug'
	WHEN _Month=9 THEN 'Sep'
	WHEN _Month=10 THEN 'Oct'
	WHEN _Month=11 THEN 'Nov'
	WHEN _Month=12 THEN 'Dec'
END AS Month_Name
INTO CovidVaccinationsEdited3
FROM CovidVaccinationsEdited2;

--4) from the object explorer pane, right click to modify column data type to numeric.


--5) you can now convert the concatenated coluumn values
SELECT *,
	CONVERT(date,CONCAT(_Day,Month_Name,_Year),101) AS newDate
INTO CovidVaccinationsFinal
FROM CovidVaccinationsEdited3;

SELECT *
FROM CovidVaccinationsFinal

--6) remove multiple columns
ALTER TABLE CovidVaccinationsFinal
DROP COLUMN _Day, _Month, _Year, Month_Name

--7) final check of the table
SELECT *
FROM CovidVaccinationsFinal