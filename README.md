# Geochemistry QAQC

Quality assurance and quality control (QAQC) of geochemical data is an important first step before any interpretation of the data is undertaken. Due to the increasing number of elements that are being reported by laboratories undertaking multi-element analysis, the time to undertake QAQC of the data has increased. In order to alleviate the increasing time constraints of undertaking QAQC this script was developed. This script provides a quick first pass of the data automatically to produce summary statistics and plot of the included standards laboratory duplicates and analytical duplicates. The statistics and plots allow for rapid assessment of geochemical data to discover potential issues with the data and trend though time. It should be noted that no general quality cut-offs have been included within the script as it does not replace the need for an expert examining the data to identify potential issues. 

This script was developed as part of Geoscience Australia’s Exploring for the Future program. Geoscience Australia’s Exploring for the Future program provides precompetitive information to inform decision-making by government, community and industry on the sustainable development of Australia's mineral, energy and groundwater resources. By gathering, analysing and interpreting new and existing precompetitive geoscience data and knowledge, we are building a national picture of Australia’s geology and resource potential. This leads to a strong economy, resilient society and sustainable environment for the benefit of all Australians. This includes supporting Australia’s transition to net zero emissions, strong, sustainable resources and agriculture sectors, and economic opportunities and social benefits for Australia’s regional and remote communities. The Exploring for the Future program, which commenced in 2016, is an eight year, $225m investment by the Australian Government.

## Dependencies
The code was developed with the following dependencies and their versions:
* numpy - 1.13.3
* scipy - 0.19.1
* pandas - 0.25.3
* matplotlib - 2.0.2
* seaborn - 0.9.0
* sklearn - 0.20.1
* matplotlib - 3.5.1
* xlsxwriter - 3.0.3
* openpyxl - 3.0.10

## Running
File Requirements: 
* The files should be excel .xlsx files with limited extraneous where possible. Whilst the script has an in built parser to find the elements unnecessary column may produce errors due to incorrect assignment. Additonally repeats of the same element name in the header will produce an error and the script will be unable to run correctly. Example files can be found in the examples folder.

Run Parameters:
* FILE_NAME - the path to the first data set. The path should be surrounded my quotation marks and preceded by an r e.g.  r"C:\Users\Desktop\Data_Set.xlsx".
* Save_Location – The folder location to save the files.
* Id_Coloumn – The name of the column where the sampels, standards, and duplictaes are named. The column name should be surrounded by quotes e.g. ‘SampleName’.
* STANDARD_CUTOFF – how many times a value in the Id_Coloumn needs to be repeated before it is added to the list of standards.
* DUPLICATE_NAME – The identifier to denote laboratory duplicates, this should include any information including spaces after the sample number, e.g. ' DUP'. If no duplicates of this type exist then it should be set at two quotes with no space.
* REPLICATE_NAME – The identifier to denote analytical duplicates, this should include any information including spaces after the sample number, e.g. ' Rpt'. If no duplicates of this type exist then it should be set at two quotes with no space.
* DEBUG – Used for limited code debugging, leave as False unless experienced with python.
* BATCHED – Boolean option (True or False) as to whether the dataset includes multiple batches of data. If set to True vertical dashed lines denoting each batch (as set by the BATCHES variable) will be added to the element plots.
* BATCH – The file save name.
* BATCHES – the row numbers for the start of each batch of samples within the data. The numbers should be separated by commas and enclosed by square brackets e.g. [328,654,981,1308,1635].

## Output data
* When run the code will produce three new folders, two excel files, and a pdf within the specified save location.
* The three folders are Standards, Analytical_Duplicates, and Lab_Duplicates:
  * The standards folder will contain a plot for each element of the difference standards through time. These standard plots will include a solid line for each standard representing the medium value and two dashed lines representing +/- 10% of the median. Some plots may also contain a black dashed line which represents the detection limit for the element. A pdf containing all the plots and an excel file containing the summary statistics are also generated in the save location. 
  * The Analytical_Duplicates folder contains a linear regression plot of the duplicate pairs for each element. Two sheets are also generated within the duplicates excel sheet containing the duplicates pairs, and the summary statistics.
  * The Lab_Duplicates folder contains a linear regression plot of the duplicate pairs for each element. Two sheets are also generated within the duplicates excel sheet containing the duplicates pairs, and the summary statistics

eCat Id: 147731
DOI: https://dx.doi.org/10.26186/147731
