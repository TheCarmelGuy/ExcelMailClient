import smtplib
import xlwings as xw
import env
import sys, getopt



#GLOBAL VARIBLES CHANGE THESE TO SET THE COLUMEN AND ROW

#the character can be changed to set a particular column in the data set
#NOTE This is for setting the column of the email account information
emailBase = 'A1' 
emailBound = 'A3' 

#the character can be changed to set a particular column in the data set
#NOTE This is for setting the column of the dues data
duesBase = 'B1'  
duesBound = 'B3' 


#return array of data points in excel sheet
#requires a lower bound, upper bound, and an excel data sheet as inputs

def gen_range(lower_bound, upper_bound,sheet):

	return sheet.range(lower_bound +':'+ upper_bound).value


def smtp_init(server):

    
    #get all the server data that you need to use email client
    server.starttls()
    server.login(env.email_username, env.email_pass);

def emailDispatcher(filename):
  
    #set up server 
    server = smtplib.SMTP('smtp.gmail.com', 587)
    
    smtp_init(server)    

 
    #load proper excel spreadsheet
    wb = xw.Book(filename)
    #assume that here A column contins all the emails and B contain the info that is needed
    datasheet = wb.sheets['emails']

    count = 0
    #iterate with customizable message
    for email, amount in zip(gen_range(emailBase, emailBound, datasheet), gen_range(duesBase,duesBound, datasheet)):
	
        template_msg = str(amount) + " is how much you owe. Thank you!"   
        server.sendmail(env.email_username,str(email), template_msg)
        print("Payment request was sent to " + email + " for the amount of $" + str(amount)) 
        
	server.quit()
	print str(count) + " emails have been proccessed"
	print("Check Log.txt for the information that has been sent out") 


def main():
	
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hf", [])
    except getopt.GetoptError as err:
        # print help information and exit:
        print "Error: " +  str(err)  # will print something like "option -a not recognized"
        sys.exit(2)
    	
    filedata = "emails.xlsx"
    for o, a in opts:
        if o == "-f":
            filedata = a           
        elif o in ("-h", "--help"):
            sys.exit()	
		
    emailDispatcher(filedata)



if __name__=="__main__": 
    main()
