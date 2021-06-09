http://rpachallenge.com/

Short Video Demo:

https://www.youtube.com/watch?v=ndPgw4SZjQw

RPA (Robotic Process Automation) challenge. The goal was to automate the flow of getting user data (first name, last name, address, etc) from an excel file and input it onto a website with dynamically changing field positions. 

My Python and Selenium solution completes the task in 3.5 seconds average. There is room for optimization but the solution is sufficient for now. This was some of my most elaborate code that I’ve done and it was very fun to research and create. 😊

Some key points of the Python code in operation order:

. Reads the excel file and extracts all the data, cell by cell, left to right, row by row in order into a list data structure.  

. At this point Selenium is launching the web driver and opening the RPA Website.

. Now with the excel file data in memory, a large list of 70 elements is shrunk into a smaller list of 7 elements and that list is fed it into a “data_input” function. 

. data_input() begins “getting” all the needed HTML elements on the website so the code knows where to click and input the list data. 

. The 7 element list is inserted into the website each element at a time. After that’s done the list is cleared and the next 7 elements from the “bigger” list get fed to the small list and the process repeats.
