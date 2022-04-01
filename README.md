# internal-inventory-manager
"a script for managing companies internal spreadsheets and forecasting sales using 2 separate methods"


This is an internal inventory managing scipt I built for a past company. It was used to manage the interactions between their inventory related spreadsheets in order to automate tasks, keep transfers of info reliable, and give a format to all entered data
this manager has 4 different functions:

General Inventory: updates a google sheet with our current inventory for each sku, Aggregate, Physical, safety quantities, Minimum Quantities, likely problems with this sku (ERRORS) and notes put on that sku within the google sheet
                   this page recieves its data from exports from the companies POS
                   
                   
                   
Warehouse Transfers: updates put warehouse transfers (and inventories) sheets when we recieve new containers of product to them. After, it updates the Containers logs that the containers have been transfered over and returns a list of updated containers
                     this page recieves its data from an internal spreadsheet that tracks status of inbound containers



Container Logs: allows you to add new or edit existing containers within the companies container logs, updates the live google sheet. 
                also has a sub-page that lets you searc the containers logs for a specific item, and return which containers conatain which quantities
                
                
                
Sales Forecasting: The sales forecasting page is somewhat seperate from the others as its returned info is used alongside the others, but doesnt actually interact with any other google sheets
                   The sales forecasting page uses sales reports from the companies POS to forecast future sales. first it aggregates the messy export into a easily readable dataframe of sales per SKU per day, the removes Out Of Stock periods to simulate regular inventory levels (at managers request), then groups sales into week-long periods 
                   It then forecasts using 2 seperate methods:
                   Avg Delta: for a given forecast length L, the Avg Delta method:
                                                                        forecas periods of 1 week, 2 weeks, 1 month, 3 months, 6 months and 9 months
                                                                        method takes the last 3 period of length L
                                                                        calculates the the average change in sales from period to period
                                                                        then assumes the same growth/decline for the next period L
                                                                        for 6 and 9 months, previous periods of sales overlap to a degree in order to use smaller datasets
                                                                        this method was useful for shorteer forecast periods, or when sales are differing from typical trends
                   Seasonality: using a period of 1.5 yrs of sales data the Seasonality method:
                                                                        generates a dataframe of quarterly sales in relation to the previous quarter, with a value for every week of the year's next quarters sales/last quarters sales
                                                                        this gives us a shape of how the time of year affects a quarter of sales in relation to the last quarter
                                                                        this can be number for the current week can be multiplied with the current last quarters sales to produce a forecast of next quarters sales
                                                                        this forecast for next quarter can be multiplied by the seasonality_dataframe value to produce a forecast of the next quarter after that, and so on, allowing us to forecast longer periods more accurately with seasonal trends
                                                                       
                   The sales forecast then updates the companies forecasting google sheet with the 2 methods forecasts, and sales per period for the previous 5 periods each, this was for checking to verify the forecast seemed to follow the trend of current sales
                   lastly functionality was added to choose which forecast to export, if not both, and to graph the chosen forecast alongside its previous sales, to verify visually, and to allow judgement calls

All these were added into a GUI to allow other coworkers to use this program as well.
