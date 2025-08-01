# parser-pipeline
 - Upload a workbook at the fastapi endpoint upload_file
 - It converts the workbook data to csv and fixes discrepencies like null values and spaces.
 - Gives the csv files generated for each sheet to azure openai api
 - gets a final_supplier_kpis.json file that contains KPI Wise structured json data where kpis are the top level object -> Company -> Month : Value
 - this is then given to insights.py which summarizes performance company wise and gives a summary of 5 keypoints/ company
 - Then these insights go to general_summary.py where a more generalized and comparative study between all companies is done and 5-10 key points are generated as general-info.json.
 - Then all of that is given as output in the api endpoint response.
