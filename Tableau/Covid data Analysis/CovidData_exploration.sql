SELECT *
FROM covid_data.dbo.CovidDeaths

SELECT *
FROM covid_data.dbo.CovidVaccinationsFinal

SELECT date

FROM covid_data.dbo.CovidVaccinations

-------------------------------------------------
--total cases VS total deaths
SELECT location, date, population, total_cases,total_deaths, 
	(total_deaths/total_cases)*100 AS DeathByCases_Percent,
	(total_deaths/population)*100 AS DeathByPopulation_Percent
FROM covid_data.dbo.CovidDeaths
WHERE location LIKE '%states%'
	 AND population != 0
	 AND total_cases != 0
ORDER BY total_deaths DESC


-------------------------------------------------
--countries with highest infection rate VS population
SELECT location, population, date,
	MAX(total_cases) AS HighInfectionRate,
	MAX((total_cases/population)*100) AS PercentPopulationInfected
FROM covid_data.dbo.CovidDeathsEdited
WHERE population != 0
GROUP BY location, population,date
ORDER BY PercentPopulationInfected DESC


-------------------------------------------------
--countries with highest death count by population
SELECT location,  
	MAX(total_deaths) AS DeathCount
	--MAX((total_deaths/population)*100) AS Fatality
FROM covid_data.dbo.CovidDeaths
GROUP BY location
ORDER BY DeathCount DESC


-------------------------------------------------
--highest death count by location
SELECT location,  
	MAX(cast(total_deaths AS int)) AS TotalDeathCount
FROM covid_data.dbo.CovidDeaths
WHERE location IN ('Asia', 'North America', 'European Union','World','United States','North America','South America')
--'Upper middle income','Lower middle income','Low income',
GROUP BY location
ORDER BY TotalDeathCount DESC


-------------------------------------------------
--highest death count by IncomeClass
SELECT location,  
	MAX(cast(total_deaths AS int)) AS TotalDeathCount
FROM covid_data.dbo.CovidDeaths
WHERE location IN ('Upper middle income','Lower middle income','Low income')
GROUP BY location
ORDER BY TotalDeathCount DESC


-------------------------------------------------
--highest death count by continent
SELECT continent,  
	MAX(cast(total_deaths AS int)) AS TotalDeathCount
FROM covid_data.dbo.CovidDeaths
WHERE continent != '' --
--WHERE continent IN ('Asia', 'North America', 'European Union','World','United States','North America','South America')
--'Upper middle income','Lower middle income','Low income',
GROUP BY continent
ORDER BY TotalDeathCount DESC


--GLOBAL NUMBERS total cases VS total deaths
SELECT  
	SUM(new_cases) AS newtotal_cases, 
	SUM(new_deaths) AS newtotal_deaths,
	(SUM(new_deaths)/SUM(new_cases))*100 AS DeathPercentage
FROM covid_data.dbo.CovidDeaths
--GROUP BY date
ORDER BY 1,2



--total population VS vaccination
SELECT cod.continent, cod.location, cod.date, cod.population, cov.new_vaccinations
FROM covid_data.dbo.CovidDeaths cod
JOIN covid_data.dbo.CovidVaccinationsFinal cov
	ON cod.location = cov.location
	AND cod.date = cov.newDate
WHERE cod.continent is not NULL
ORDER BY 1,2,3




--looking at continent and location result 
SELECT location, continent
FROM covid_data.dbo.CovidDeaths
WHERE continent=''
--found continent in the location column
--found social classes in the location column

--checking if social classes are related to other columns 
SELECT * 
FROM covid_data.dbo.CovidDeaths
WHERE continent='' AND location IN ('Upper middle income','Lower middle income','Low income')

--if not related to other column then we create another temporary table
SELECT * INTO SocialClasses
FROM covid_data.dbo.CovidDeaths
WHERE continent='' AND location IN ('Upper middle income','Lower middle income','Low income')

--checking new table
SELECT *
FROM SocialClasses


--checking data CovidDeaths
SELECT *  
FROM covid_data.dbo.CovidDeaths
WHERE continent!=''
--result with 195,865 rows

--checking data CovidVaccinations
SELECT *
FROM covid_data.dbo.CovidVaccinations
WHERE continent!=''
--result with 195,865 rows


--created new table for a work around on convertion error, I was not able to modify it thru the object explorer. see the error below
--"Unable to modify table. Conversion failed when converting date and/or time from character string."
--refer to CovidData_Cleaning, the way I solve the date error. took me 2 days researching doing trail and error.
--To continue our data exploration



-- Total Population vs Vaccinations
-- Shows Percentage of Population that has recieved at least one Covid Vaccine

SELECT code.continent, code.location, code.date, code.population, covf.new_vaccinations,
	SUM(covf.new_vaccinations) OVER (Partition by code.Location Order by code.location, code.Date) as RollingPeopleVaccinated
FROM covid_data.dbo.CovidDeathsEdited code
INNER JOIN covid_data.dbo.CovidVaccinationsFinal covf
	ON code.location = covf.location
	AND code.date = covf.newDate
ORDER BY 2,3


-- Using CTE to perform Calculation on Partition By in previous query

With PopvsVac (Continent, Location, Date, Population, New_Vaccinations, RollingPeopleVaccinated)
as
(
SELECT code.continent, code.location, code.date, code.population, covf.new_vaccinations,
	SUM(covf.new_vaccinations) OVER (Partition by code.Location Order by code.location, code.Date) as RollingPeopleVaccinated
FROM covid_data.dbo.CovidDeathsEdited code
INNER JOIN covid_data.dbo.CovidVaccinationsFinal covf
	ON code.location = covf.location
	AND code.date = covf.newDate
)
Select *, (RollingPeopleVaccinated/Population)*100
From PopvsVac
WHERE Population !=0



-- Using Temp Table to perform Calculation on Partition By in previous query

DROP Table if exists #PercentPopulationVaccinated
Create Table #PercentPopulationVaccinated
(
Continent nvarchar(255),
Location nvarchar(255),
Date datetime,
Population numeric,
New_vaccinations numeric,
RollingPeopleVaccinated numeric
)

Insert into #PercentPopulationVaccinated
SELECT code.continent, code.location, code.date, code.population, covf.new_vaccinations,
	SUM(covf.new_vaccinations) OVER (Partition by code.Location Order by code.location, code.Date) as RollingPeopleVaccinated
FROM covid_data.dbo.CovidDeathsEdited code
INNER JOIN covid_data.dbo.CovidVaccinationsFinal covf
	ON code.location = covf.location
	AND code.date = covf.newDate
--where dea.continent is not null 
--order by 2,3

Select *, (RollingPeopleVaccinated/Population)*100
From #PercentPopulationVaccinated
WHERE Population !=0





-- Creating View to store data for later visualizations

Create View PercentPopulationVaccinated as
SELECT code.continent, code.location, code.date, code.population, covf.new_vaccinations,
	SUM(covf.new_vaccinations) OVER (Partition by code.Location Order by code.location, code.Date) as RollingPeopleVaccinated
FROM covid_data.dbo.CovidDeathsEdited code
INNER JOIN covid_data.dbo.CovidVaccinationsFinal covf
	ON code.location = covf.location
	AND code.date = covf.newDate




SELECT continent, date, population, total_cases
FROM covid_data.dbo.CovidDeathsEdited

