# pythoncode.py
# importing the required module 
import xlwings as xw
import matplotlib.pyplot as plt 

#grapth function which is to be called from python code
def graph():
        #set the active sheet(The Excel file with the data needs to be open and active)
        sht = xw.sheets.active

        # x axis values 
        month = sht.range('A2:A13').value
        
        # corresponding y axis values 
        retail = sht.range('B2:B13').value

        online = sht.range('C2:C13').value

        vendor = sht.range('D2:D13').value
        
        # plotting the points 
        plt.plot(month, retail,label="Retail") 
        plt.plot(month, online,label="Online")
        plt.plot(month, vendor,label="Vendor")

        #making the legends visible (Retail,Online,Vendor) 
        plt.legend(loc=2, ncol=2)

        # naming the x axis 
        plt.xlabel('Months') 

        # naming the y axis 
        plt.ylabel('Business Type') 

        # giving a title to my graph 
        plt.title('Monthly Sales Report!') 

        # function to show the plot 
        plt.show() 
