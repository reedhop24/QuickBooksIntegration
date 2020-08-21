# IBQ QuickBooks Customer View

##### IBQ Systems has not been collecting information on cutomer retention until recently; however, we needed a way to access our retention rate of customers from the past. Therefore I built this little api in order to interact with QuickBook's API and return the relevant data in a spreadsheet. I had made an GET to QuickBook's Customers API to receive all of our Customer's. This API consumes the results of that API in JSON data.

##### Once this program receives the Customer JSON data, I begin looping through that data all the while writing the customer's id, name, start date, and end date to an excel spreadsheet. I also make a GET request to their transaction list API to receive their most recent transaction if the type is an invoice or a Payment. 

##### I also have two functions to take the date time received and format it to YYYY-MM-DD. For the date that the customer started with us I had to add a month to it since they were not getting billed until the 1st of the next month.